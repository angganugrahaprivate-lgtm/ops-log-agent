require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');
const XLSX = require('xlsx');

const app = express();
app.use(express.json());

// ─── CONFIG ───────────────────────────────────────────────
const ANTHROPIC_KEY      = process.env.ANTHROPIC_API_KEY;
const SHEET_ID           = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB          = process.env.SHEET_TAB_NAME || 'Data Handover';
const MEMORY_TAB         = 'Memory';
const REMINDER_TAB       = 'Reminders';
const UPDATE_LOG_TAB     = 'Update Log';
const FONNTE_TOKEN       = process.env.FONNTE_TOKEN;
const TELEGRAM_TOKEN     = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID   = process.env.TELEGRAM_NOTIF_CHAT_ID;
const YOUR_WA_NUMBER     = process.env.YOUR_WA_NUMBER;
const GROUP_WA_ID        = process.env.GROUP_WA_ID;
const BOT_MENTION        = (process.env.BOT_MENTION_NAME || 'Nyenyenye').toLowerCase();
const BOT_WA_NUMBER      = process.env.BOT_WA_NUMBER || '';
const REMINDER_TARGETS   = (process.env.REMINDER_TARGETS || '').split(',').filter(Boolean);
const PORT               = process.env.PORT || 3000;

console.log('=== ENV CHECK ===');
console.log('ANTHROPIC_API_KEY:', ANTHROPIC_KEY ? 'OK' : 'MISSING');
console.log('GOOGLE_SHEETS_ID:', SHEET_ID ? 'OK' : 'MISSING');
console.log('GROUP_WA_ID:', GROUP_WA_ID || 'NOT SET');
console.log('BOT_MENTION:', BOT_MENTION);
console.log('REMINDER_TARGETS:', REMINDER_TARGETS.length, 'numbers');
console.log('=================');

const anthropic = new Anthropic({ apiKey: ANTHROPIC_KEY });

// ─── COLUMN MAPPING ───────────────────────────────────────
const COL = {
  tanggal:0, cutoff:1, shippingNum:2, koli:3, namaBarang:4,
  partner:5, customer:6, alamat:7, kecamatan:8, kota:9,
  telepon:10, kodePos:11, provinsi:12, ekspedisi:13, resi:14,
  tglPengiriman:15, tglTiba:16, aging:17, requestDate:18,
  remark:19, status:20, sla:21, noOrder:22,
};
const GENERAL_COLS  = [0,2,6,9,13,14,15,16,18,19,20,21,22];
const GENERAL_NAMES = ['Tanggal','Shipping Number','Nama Customer','Kota','Ekspedisi','Resi','Tgl Pengiriman','Tgl Tiba','Request Date','Remark','Status','SLA','No Order'];
const UPDATE_COLS   = [22,6,13,20,14,15,16,19];
const UPDATE_NAMES  = ['No Order','Nama Customer','Ekspedisi','Status','Resi','Tgl Pengiriman','Tgl Tiba','Remark'];

// ─── CACHE (1 JAM) ────────────────────────────────────────
const cache = {};
const CACHE_TTL = 60 * 60 * 1000;
function getCache(key) { const c=cache[key]; if(c&&Date.now()-c.time<CACHE_TTL)return c.data; return null; }
function setCache(key,data){ cache[key]={data,time:Date.now()}; }
function clearCache(key){ if(key)delete cache[key]; else Object.keys(cache).forEach(k=>delete cache[k]); }

// ─── GOOGLE SHEETS ────────────────────────────────────────
let sheetsClient = null;
async function getSheets() {
  if (sheetsClient) return sheetsClient;
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    sheetsClient = google.sheets({ version: 'v4', auth });
    console.log('Google Sheets OK');
    return sheetsClient;
  } catch (e) { console.error('Sheets init error:', e.message); return null; }
}

function toCol(n) { let c=''; while(n>=0){c=String.fromCharCode(65+(n%26))+c;n=Math.floor(n/26)-1;} return c; }
function getToday() { return new Date().toLocaleDateString('sv-SE',{timeZone:'Asia/Jakarta'}); }
function getCutOff() { const h=parseInt(new Date().toLocaleString('en-US',{timeZone:'Asia/Jakarta',hour:'numeric',hour12:false})); return h<14?'1':'2'; }

function parseDate(s) {
  if (!s) return '';
  s = s.toString().trim();
  const mon = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,mei:5,maret:3,juni:6,juli:7,agustus:8,oktober:10,november:11,desember:12};
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m) return s;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  m = s.match(/(\d{1,2})\s*[-\s]+\s*([a-zA-Z]+)\s*[-\s]+\s*(\d{4})/);
  if(m){ const mn=mon[m[2].toLowerCase().trim().substring(0,3)]; if(mn) return `${m[3]}-${String(mn).padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
  return s;
}

function formatDateID(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('id-ID',{day:'2-digit',month:'long',year:'numeric',timeZone:'Asia/Jakarta'});
}

async function getSheetData() {
  const cached = getCache('sheetData');
  if (cached) { console.log('Cache hit: sheetData'); return cached; }
  const s = await getSheets(); if (!s) return null;
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId:SHEET_ID, range:`${SHEET_TAB}!A:AK` });
    const data = res.data.values || [];
    setCache('sheetData', data);
    console.log(`Sheet loaded: ${data.length} rows`);
    return data;
  } catch (e) { console.error('getSheetData error:', e.message); return null; }
}

function selectColumns(rows, colIndices, colNames) {
  return rows.map(row => { const obj={}; colIndices.forEach((idx,i)=>{obj[colNames[i]]=(row[idx]||'');}); return obj; });
}

async function updateCell(sheetTab, rowNum, colIdx, value) {
  const s = await getSheets(); if (!s) return;
  await s.spreadsheets.values.update({ spreadsheetId:SHEET_ID, range:`${sheetTab}!${toCol(colIdx)}${rowNum}`, valueInputOption:'RAW', requestBody:{values:[[value]]} });
  clearCache('sheetData');
}

async function ensureTab(tabName, headers) {
  const s = await getSheets(); if (!s) return;
  try { await s.spreadsheets.values.get({ spreadsheetId:SHEET_ID, range:`${tabName}!A1` }); }
  catch (e) {
    try {
      const { google } = require('googleapis');
      const auth = new google.auth.GoogleAuth({ credentials:JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes:['https://www.googleapis.com/auth/spreadsheets'] });
      const api = google.sheets({ version:'v4', auth });
      await api.spreadsheets.batchUpdate({ spreadsheetId:SHEET_ID, requestBody:{requests:[{addSheet:{properties:{title:tabName}}}]} });
      await api.spreadsheets.values.update({ spreadsheetId:SHEET_ID, range:`${tabName}!A1`, valueInputOption:'RAW', requestBody:{values:[headers]} });
      console.log(`Tab "${tabName}" created`);
    } catch (e2) { console.error('ensureTab error:', e2.message); }
  }
}

// ─── UPDATE LOG ───────────────────────────────────────────
async function logUpdate(updatedBy, noOrder, kolom, nilaiLama, nilaiBaru) {
  await ensureTab(UPDATE_LOG_TAB, ['Timestamp','Oleh','No Order','Kolom','Nilai Lama','Nilai Baru']);
  const s = await getSheets(); if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'});
    await s.spreadsheets.values.append({ spreadsheetId:SHEET_ID, range:`${UPDATE_LOG_TAB}!A:F`, valueInputOption:'RAW', requestBody:{values:[[ts,updatedBy,noOrder,kolom,nilaiLama,nilaiBaru]]} });
  } catch (e) { console.error('logUpdate error:', e.message); }
}

// ─── UPDATE FUNCTIONS ─────────────────────────────────────
async function updateOrderField(noOrder, colIdx, colName, newValue, updatedBy) {
  const data = await getSheetData(); if (!data) return false;
  for (let i=1; i<data.length; i++) {
    if ((data[i][COL.noOrder]||'').trim() === noOrder.trim()) {
      const oldValue = data[i][colIdx] || '';
      await updateCell(SHEET_TAB, i+1, colIdx, newValue);
      await logUpdate(updatedBy, noOrder, colName, oldValue, newValue);
      console.log(`Updated ${colName} for ${noOrder}: ${oldValue} → ${newValue}`);
      return true;
    }
  }
  return false;
}

async function getOrderByNumber(noOrder) {
  const data = await getSheetData(); if (!data || data.length < 2) return null;
  for (let i=1; i<data.length; i++) {
    if ((data[i][COL.noOrder]||'').includes(noOrder) || (data[i][COL.shippingNum]||'').includes(noOrder)) {
      return data[i];
    }
  }
  return null;
}

// ─── MEMORY ───────────────────────────────────────────────
async function getMemory() {
  const cached = getCache('memory'); if (cached) return cached;
  await ensureTab(MEMORY_TAB, ['Timestamp','Category','Content']);
  const s = await getSheets(); if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId:SHEET_ID, range:`${MEMORY_TAB}!A:C` });
    const data = (res.data.values||[]).slice(1).slice(-50);
    setCache('memory', data);
    return data;
  } catch (e) { return []; }
}

async function saveMemory(category, content) {
  await ensureTab(MEMORY_TAB, ['Timestamp','Category','Content']);
  const s = await getSheets(); if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'});
    await s.spreadsheets.values.append({ spreadsheetId:SHEET_ID, range:`${MEMORY_TAB}!A:C`, valueInputOption:'RAW', requestBody:{values:[[ts,category,content]]} });
    clearCache('memory');
    console.log(`Memory: [${category}] ${content}`);
  } catch (e) { console.error('saveMemory error:', e.message); }
}

// ─── REMINDER ─────────────────────────────────────────────
async function getReminders() {
  const cached = getCache('reminders'); if (cached) return cached;
  await ensureTab(REMINDER_TAB, ['Timestamp','Dibuat Oleh','No Order','Tanggal Reminder','Note','Status']);
  const s = await getSheets(); if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId:SHEET_ID, range:`${REMINDER_TAB}!A:F` });
    const data = (res.data.values||[]).slice(1);
    setCache('reminders', data);
    return data;
  } catch (e) { return []; }
}

async function saveReminder(createdBy, noOrders, tanggal, note) {
  // noOrders bisa string tunggal atau array
  const noOrderStr = Array.isArray(noOrders) ? noOrders.join(',') : noOrders;
  await ensureTab(REMINDER_TAB, ['Timestamp','Dibuat Oleh','No Order','Tanggal Reminder','Note','Status']);
  const s = await getSheets(); if (!s) return false;
  try {
    const ts = new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'});
    await s.spreadsheets.values.append({ spreadsheetId:SHEET_ID, range:`${REMINDER_TAB}!A:F`, valueInputOption:'RAW', requestBody:{values:[[ts,createdBy,noOrderStr||'',tanggal,note,'Pending']]} });
    clearCache('reminders');
    console.log(`Reminder saved by ${createdBy}: ${noOrderStr} on ${tanggal}`);
    return true;
  } catch (e) { console.error('saveReminder error:', e.message); return false; }
}

async function markReminderDone(rowIndex) {
  const s = await getSheets(); if (!s) return;
  try { await updateCell(REMINDER_TAB, rowIndex+2, 5, 'Sent'); clearCache('reminders'); }
  catch (e) { console.error('markReminderDone error:', e.message); }
}

// ─── BUILD REMINDER MESSAGE ────────────────────────────────
async function buildReminderMsg(reminder, label) {
  const [ts, createdBy, noOrderStr, tanggal, note] = reminder;
  const noOrders = (noOrderStr||'').split(',').map(s=>s.trim()).filter(Boolean);
  const data = await getSheetData();

  let orders = [];
  if (data && noOrders.length > 0) {
    for (const no of noOrders) {
      const row = data.slice(1).find(r => (r[COL.noOrder]||'').includes(no)||(r[COL.shippingNum]||'').includes(no));
      if (row) orders.push(row);
    }
  }

  const join = (field) => orders.length > 0 ? orders.map(r => r[field]||'—').filter((v,i,a)=>a.indexOf(v)===i).join(' & ') : '—';

  let msg = `${label} - ${formatDateID(tanggal)}\n\n`;
  msg += `📦 No Order  : ${noOrders.join(' & ') || '—'}\n`;
  msg += `🚢 Shipping  : ${join(COL.shippingNum)}\n`;
  msg += `📍 Kota      : ${join(COL.kota)}\n`;
  msg += `🚚 Ekspedisi : ${join(COL.ekspedisi)}\n`;
  msg += `📅 Tgl Order : ${join(COL.tanggal)}\n`;
  msg += `📝 Remark    : ${join(COL.remark)}\n`;
  msg += `\n📌 Note      : ${note}\n`;
  msg += `👤 Dibuat oleh: ${createdBy}`;
  return msg;
}

// ─── SEND TO REMINDER TARGETS ─────────────────────────────
async function sendToTargets(message) {
  const targets = [...REMINDER_TARGETS];
  if (YOUR_WA_NUMBER && !targets.includes(YOUR_WA_NUMBER)) targets.push(YOUR_WA_NUMBER);
  for (const num of targets) {
    await sendWA(num, message);
  }
  if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, message);
}

// ─── EXCEL UPLOAD ─────────────────────────────────────────
async function handleExcelUpload(fileBuffer, senderName, chatId, isWA=false) {
  try {
    const wb = XLSX.read(fileBuffer,{type:'buffer'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws,{header:1});
    if (rows.length < 2) return 'File kosong atau tidak ada data.';
    const headers = rows[0];
    const dataRows = rows.slice(1).filter(r=>r.some(c=>c));
    const numIdx    = headers.indexOf('Number');
    const koliIdx   = headers.indexOf('Koli');
    const barangIdx = headers.indexOf('Nama Barang');
    const partnerIdx= headers.indexOf('Partner');
    const nameIdx   = headers.findIndex(h=>h&&h.toString().includes('First Name'));
    const alamatIdx = headers.findIndex(h=>h&&h.toString().includes('Line1'));
    const kecIdx    = headers.findIndex(h=>h&&h.toString().includes('Line3'));
    const kotaIdx   = headers.findIndex(h=>h&&h.toString().includes('Line4'));
    const phoneIdx  = headers.findIndex(h=>h&&h.toString().includes('Phone'));
    const posIdx    = headers.findIndex(h=>h&&h.toString().includes('Postcode'));
    const stateIdx  = headers.findIndex(h=>h&&h.toString().includes('State'));
    const today=getToday(), cutoff=getCutOff();
    let preview=`📋 PREVIEW DATA EXCEL\nTanggal: ${today} | Cut Off: ${cutoff}\nTotal: ${dataRows.length} order\n\n`;
    dataRows.forEach((r,i)=>{ preview+=`${i+1}. ${r[numIdx]||'—'} - ${r[nameIdx]||'—'} - ${r[partnerIdx]||'—'}\n`; });
    preview+=`\nMasukkan semua ke GSheet? Balas "ya" untuk konfirmasi.`;
    cache[`pendingExcel_${chatId}`]={data:dataRows,time:Date.now(),today,cutoff,numIdx,koliIdx,barangIdx,partnerIdx,nameIdx,alamatIdx,kecIdx,kotaIdx,phoneIdx,posIdx,stateIdx,senderName,chatId,isWA};
    return preview;
  } catch (e) { console.error('handleExcelUpload error:', e.message); return 'Gagal membaca file Excel.'; }
}

async function insertExcelToSheet(senderName, chatId) {
  const pendingKey=`pendingExcel_${chatId}`;
  const pending=cache[pendingKey];
  if (!pending) return 'Tidak ada data Excel yang menunggu konfirmasi.';
  const s=await getSheets(); if (!s) return 'Gagal konek ke GSheet.';
  try {
    const {data,today,cutoff,numIdx,koliIdx,barangIdx,partnerIdx,nameIdx,alamatIdx,kecIdx,kotaIdx,phoneIdx,posIdx,stateIdx}=pending;
    const rows=data.map(r=>{
      const row=Array(36).fill('');
      row[COL.tanggal]=today; row[COL.cutoff]=cutoff;
      row[COL.shippingNum]=r[numIdx]||''; row[COL.koli]=r[koliIdx]||'';
      row[COL.namaBarang]=r[barangIdx]||''; row[COL.partner]=r[partnerIdx]||'';
      row[COL.customer]=r[nameIdx]||''; row[COL.alamat]=r[alamatIdx]||'';
      row[COL.kecamatan]=r[kecIdx]||''; row[COL.kota]=r[kotaIdx]||'';
      row[COL.telepon]=r[phoneIdx]||''; row[COL.kodePos]=r[posIdx]||'';
      row[COL.provinsi]=r[stateIdx]||'';
      return row;
    });
    await s.spreadsheets.values.append({ spreadsheetId:SHEET_ID, range:`${SHEET_TAB}!A:AK`, valueInputOption:'RAW', requestBody:{values:rows} });
    await logUpdate(senderName,'BULK INSERT','Excel Upload','-',`${rows.length} order baru`);
    clearCache('sheetData'); delete cache[pendingKey];
    return `✅ ${rows.length} order berhasil dimasukkan!\nTanggal: ${today} | Cut Off: ${cutoff}`;
  } catch (e) { console.error('insertExcelToSheet error:', e.message); return 'Gagal insert: '+e.message; }
}

async function getFonnteFile(fileUrl) {
  try {
    const res=await axios.get(fileUrl,{responseType:'arraybuffer',headers:{Authorization:FONNTE_TOKEN}});
    return Buffer.from(res.data);
  } catch (e) { console.error('getFonnteFile error:', e.message); return null; }
}

// ─── SMART FILTER ─────────────────────────────────────────
const chatHistory = {};
const NO_DATA_INTENTS=['greeting','help','out_of_scope','save_instruction'];

function detectIntent(message, senderId) {
  const msg=(message||'').toLowerCase().trim();

  if (/^(halo|hai|hi|hello|selamat|pagi|siang|sore|malam|hey)/.test(msg)) return 'greeting';
  if (/^(oke|ok|tidak|ga|gak|siap|done|sip|noted|thanks|makasih)$/.test(msg)) return 'greeting';
  if (/kamu bisa|fitur apa|help|bantuan|cara pakai/.test(msg)) return 'help';
  if (/^(catat|ingat|note)\s*:/i.test(msg)) return 'save_instruction';
  if (/refresh|sync data|reload data/.test(msg)) return 'refresh';
  if (/log hari ini|history update|apa.*diupdate/.test(msg)) return 'log_today';

  // Excel confirm per sender
  const hasPending = senderId && cache[`pendingExcel_${senderId}`];
  if (/^(ya|yes|iya|yep|yup|konfirmasi|insert|masukkan)$/.test(msg) && hasPending) return 'confirm_excel';
  if (/upload.*(excel|xlsx|file)|kirim.*(excel|file)/.test(msg)) return 'prompt_excel';

  // Format 1: JNL update
  if (/no order\s*:/i.test(msg) && /tgl kirim\s*:/i.test(msg)) return 'format1_update';
  // Format 2: Remark
  if (/no order\s*:/i.test(msg) && /remark\s*:/i.test(msg)) return 'format2_remark';
  // Format 3: Reminder terstruktur
  if (/^reminder/i.test(msg) && /no order\s*:/i.test(msg) && /tgl\s*:/i.test(msg)) return 'format3_reminder';

  // Updates
  if (/update resi|input resi|tambah resi/.test(msg)) return 'update_resi';
  if (/update.*tgl.*kirim|update.*tanggal.*kirim|tgl pengiriman/.test(msg)) return 'update_tgl_kirim';
  if (/update.*tgl.*tiba|update.*tanggal.*tiba|sudah tiba/.test(msg)) return 'update_tgl_tiba';
  if (/update.*remark|tambah.*remark|isi.*remark|remark.*order/.test(msg)) return 'update_remark';

  // Monitoring
  if (/pengiriman hari ini|shipped hari ini|dikirim hari ini/.test(msg)) return 'shipped_today';
  if (/^(reminder|ingatkan|set reminder)/.test(msg)) return 'reminder';
  if (/cek reminder|list reminder|reminder apa/.test(msg)) return 'list_reminder';
  if (/overdue|telat|terlambat|lewat sla/.test(msg)) return 'overdue';
  if (/sla|deadline|mau deadline|mendekati|urgent/.test(msg)) return 'sla_alert';
  if (/pending|waiting|belum dikirim/.test(msg)) return 'pending';
  if (/belum.*resi|tanpa resi/.test(msg)) return 'no_resi';
  if (/summary|rangkum|rekap|laporan harian/.test(msg)) return 'summary';
  if (/performa|analisa.*ekspedisi|laporan.*ekspedisi/.test(msg)) return 'analytics';
  if (/jne|j&t|jnt|sicepat|anteraja|ninja|tiki|lion|jnl|deliveree|sentral/.test(msg)) return 'ekspedisi';
  if (/\d{6,}/.test(msg)) return 'specific_order';
  if (/nama|customer|cari.*nama|order.*untuk/.test(msg)) return 'customer_search';

  // Conversation-aware fallback
  if (senderId && chatHistory[senderId] && chatHistory[senderId].length > 0 && msg.length < 80) {
    return 'general';
  }

  return 'out_of_scope';
}

function filterData(data, intent, message) {
  if (!data || data.length < 2) return null;
  const rows=data.slice(1);
  const today=getToday();
  const UPDATE_INTENTS=['update_resi','update_tgl_kirim','update_tgl_tiba','update_remark','format1_update','format2_remark'];
  if (UPDATE_INTENTS.includes(intent)) {
    const filtered=rows.filter(r=>(r[COL.status]||'').toLowerCase()!=='received');
    return {names:UPDATE_NAMES, rows:selectColumns(filtered,UPDATE_COLS,UPDATE_NAMES)};
  }
  let filtered;
  switch(intent) {
    case 'pending': filtered=rows.filter(r=>/pending|waiting/i.test(r[COL.status]||'')); break;
    case 'no_resi': filtered=rows.filter(r=>!(r[COL.resi]||'').trim()&&(r[COL.status]||'').toLowerCase()!=='received'); break;
    case 'overdue': filtered=rows.filter(r=>{const s=r[COL.sla];return s&&s<today&&(r[COL.status]||'').toLowerCase()!=='received';}); break;
    case 'sla_alert': filtered=rows.filter(r=>{const s=r[COL.sla];if(!s||(r[COL.status]||'').toLowerCase()==='received')return false;const d=Math.round((new Date(s)-new Date(today))/86400000);return d>=-1&&d<=2;}); break;
    case 'shipped_today': filtered=rows.filter(r=>(r[COL.tglPengiriman]||'').startsWith(today)); break;
    case 'specific_order':
      const orderNum=(message.match(/\d{6,}/)||[])[0];
      filtered=orderNum?rows.filter(r=>(r[COL.noOrder]||'').includes(orderNum)||(r[COL.shippingNum]||'').includes(orderNum)):rows.slice(-50); break;
    case 'customer_search':
      const words=message.split(/\s+/).filter(w=>w.length>3);
      filtered=rows.filter(r=>words.some(w=>(r[COL.customer]||'').toLowerCase().includes(w.toLowerCase()))); break;
    case 'ekspedisi':
      const ekspKeywords=['jne','j&t','jnt','sicepat','anteraja','ninja','tiki','lion','jnl','deliveree','sentral'];
      const matchEksp=ekspKeywords.find(k=>message.toLowerCase().includes(k));
      filtered=matchEksp?rows.filter(r=>(r[COL.ekspedisi]||'').toLowerCase().includes(matchEksp)):rows.slice(-100); break;
    case 'reminder': case 'list_reminder': case 'format3_reminder': return null;
    default: filtered=rows.slice(-200); break;
  }
  console.log(`Filter [${intent}]: ${filtered.length}/${rows.length} rows`);
  return {names:GENERAL_NAMES, rows:selectColumns(filtered,GENERAL_COLS,GENERAL_NAMES)};
}

// ─── MONITORING ───────────────────────────────────────────
function computeSLAAlerts(data) {
  if (!data||data.length<2) return null;
  const today=getToday();
  const result={urgent:[],warning:[],attention:[],overdue:[]};
  data.slice(1).forEach(row=>{
    const status=(row[COL.status]||'').toLowerCase();
    if(status==='received') return;
    const slaStr=row[COL.sla]; if(!slaStr) return;
    const diff=Math.round((new Date(slaStr)-new Date(today))/86400000);
    const info={no_order:row[COL.noOrder],customer:row[COL.customer],kota:row[COL.kota],status:row[COL.status],sla:slaStr,eksp:row[COL.ekspedisi],resi:row[COL.resi],diff};
    if(diff<0)result.overdue.push(info);
    else if(diff===0)result.urgent.push(info);
    else if(diff===1)result.warning.push(info);
    else if(diff<=2)result.attention.push(info);
  });
  return result;
}

function computeDailySummary(data) {
  if (!data||data.length<2) return null;
  const today=getToday();
  const rows=data.slice(1);
  const kotaCount={};
  rows.filter(r=>(r[COL.status]||'').toLowerCase()!=='received').forEach(r=>{const k=r[COL.kota]||'Unknown';kotaCount[k]=(kotaCount[k]||0)+1;});
  return {
    receivedToday:rows.filter(r=>(r[COL.tanggal]||'').startsWith(today)).length,
    shippedToday:rows.filter(r=>(r[COL.tglPengiriman]||'').startsWith(today)).length,
    waiting:rows.filter(r=>/waiting|pending/i.test(r[COL.status]||'')).length,
    noResi:rows.filter(r=>!(r[COL.resi]||'').trim()&&(r[COL.status]||'').toLowerCase()!=='received').length,
    overdue:rows.filter(r=>{const s=r[COL.sla];return s&&s<today&&(r[COL.status]||'').toLowerCase()!=='received';}).length,
    topKota:Object.entries(kotaCount).sort((a,b)=>b[1]-a[1]).slice(0,3),
  };
}

function computeEkspReport(data) {
  if (!data||data.length<2) return null;
  const today=getToday(); const report={};
  data.slice(1).forEach(row=>{
    const eksp=row[COL.ekspedisi]||'Unknown';
    if(!report[eksp])report[eksp]={total:0,ontime:0,late:0,pending:0};
    report[eksp].total++;
    const status=(row[COL.status]||'').toLowerCase(), sla=row[COL.sla]||'';
    if(status==='received'){(sla&&row[COL.tglTiba]&&row[COL.tglTiba]<=sla)?report[eksp].ontime++:report[eksp].late++;}
    else if(sla&&sla<today)report[eksp].late++;
    else report[eksp].pending++;
  });
  return report;
}

// ─── SYSTEM PROMPT ────────────────────────────────────────
const SYSTEM_PROMPT = `
Kamu adalah OPS Agent untuk Warehouse & Logistik Palembang (nama: Nyenyenye).
Bahasa: Indonesia casual. Nada: Santai tapi profesional.

## KEMAMPUAN
1. Cek & monitor order (pending, SLA, overdue, ekspedisi, summary)
2. Update Resi → ACTION:UPDATE_RESI:[no_order]:[nilai]
3. Update Tgl Pengiriman → ACTION:UPDATE_TGL_KIRIM:[no_order]:[YYYY-MM-DD]
4. Update Tgl Tiba → ACTION:UPDATE_TGL_TIBA:[no_order]:[YYYY-MM-DD]
5. Update Remark → ACTION:UPDATE_REMARK:[no_order]:[teks remark]
6. Set Reminder → ACTION:SAVE_REMINDER:[no_order(s)]:[YYYY-MM-DD]:[note]
7. Simpan instruksi/memory → ACTION:SAVE_MEMORY:[category]:[content]
8. Cek log update, pengiriman hari ini, upload Excel
9. Sapaan & pertanyaan ringan

JIKA di luar list → balas TEPAT: "Hmmm gatau sih, diluar konteks itu keknya 😅"

## FORMAT PESAN YANG DIKENALI

### Format 1 — Update JNL/Palembang
Ketika menerima pesan dengan field: HARI, TGL, JAM, NO ORDER, EKSPEDISI, Driver, TGL KIRIM, TGL TIBA, KET
→ Parse semua field, tampilkan konfirmasi ringkas, lalu execute:
ACTION:UPDATE_RESI:[no_order]:[nama_driver]
ACTION:UPDATE_TGL_KIRIM:[no_order]:[YYYY-MM-DD]
ACTION:UPDATE_TGL_TIBA:[no_order]:[YYYY-MM-DD]
Catatan: KET diabaikan (tidak diupdate). Konversi tanggal "12 - May - 2026" → "2026-05-12"

### Format 2 — Remark Order
Ketika menerima: NO ORDER + REMARK
→ ACTION:UPDATE_REMARK:[no_order]:[isi remark]

### Format 3 — Set Reminder
Ketika menerima: REMINDER + NO ORDER + TGL + NOTE
→ NO ORDER bisa berisi lebih dari satu (dipisah "dan", ",", "&")
→ ACTION:SAVE_REMINDER:[no_order1,no_order2]:[YYYY-MM-DD]:[note]

## INSTRUKSI MEMORY
Ketika user ketik "Catat: ..." atau "Ingat: ..."
→ ACTION:SAVE_MEMORY:instruksi:[content]
→ Konfirmasi: "✅ Dicatat! Aku akan ingat ini."

## ATURAN PENTING
- ACTION hanya di akhir pesan, tidak ditampilkan ke user
- Selalu konfirmasi sebelum update (kecuali Format 1, 2, 3 yang sudah jelas)
- Jika order tidak ditemukan → beritahu dengan jelas
- Gunakan data dan instruksi yang tersedia di konteks
`.trim();

// ─── CLAUDE CALL ──────────────────────────────────────────
async function callClaude(senderId, senderName, userMessage, imageBase64=null, imageMime='image/jpeg') {
  const today = getToday();

  // Jalur foto
  if (imageBase64) {
    if (!chatHistory[senderId]) chatHistory[senderId]=[];
    const messages=[...chatHistory[senderId]];
    messages.push({role:'user',content:[
      {type:'image',source:{type:'base64',media_type:imageMime,data:imageBase64}},
      {type:'text',text:`Tanggal: ${today}\nUser: ${senderName}\n\n${userMessage||'Tolong baca nomor resi dari foto ini.'}`}
    ]});
    const resp=await anthropic.messages.create({model:'claude-haiku-4-5-20251001',max_tokens:1000,system:[{type:'text',text:SYSTEM_PROMPT,cache_control:{type:'ephemeral'}}],messages});
    const reply=resp.content[0].text;
    chatHistory[senderId].push({role:'user',content:userMessage||'[foto]'});
    chatHistory[senderId].push({role:'assistant',content:reply});
    if(chatHistory[senderId].length>12)chatHistory[senderId]=chatHistory[senderId].slice(-12);
    await parseActions(reply,senderId,senderName);
    return reply.replace(/ACTION:[^\n]+/g,'').trim();
  }

  // Jalur teks
  const intent = detectIntent(userMessage||'', senderId);
  if (intent==='out_of_scope') return 'Hmmm gatau sih, diluar konteks itu keknya 😅';
  if (intent==='confirm_excel') return await insertExcelToSheet(senderName, senderId);
  if (intent==='prompt_excel') return '📎 Silakan kirim file Excel-nya langsung ke sini ya!';
  if (intent==='refresh') { clearCache(); return '🔄 Data berhasil di-refresh!'; }
  if (intent==='log_today') return await getLogToday();
  if (intent==='save_instruction') {
    const content=(userMessage||'').replace(/^(catat|ingat|note)\s*:/i,'').trim();
    await saveMemory('instruksi', content);
    return '✅ Dicatat! Aku akan ingat instruksi ini.';
  }

  // Load data
  let rawData=null, filteredResult=null;
  if (!NO_DATA_INTENTS.includes(intent)) {
    rawData=await getSheetData();
    filteredResult=filterData(rawData,intent,userMessage||'');
  }

  const [memories,reminders]=await Promise.all([getMemory(),getReminders()]);
  const instruksi=memories.filter(m=>m[1]==='instruksi');
  const memoryOther=memories.filter(m=>m[1]!=='instruksi');
  const todayReminders=reminders.filter(r=>r[3]===today&&r[5]==='Pending');
  const upcomingReminders=reminders.filter(r=>r[3]>=today&&r[5]==='Pending');
  const slaAlerts=['sla_alert','overdue','summary'].includes(intent)?computeSLAAlerts(rawData):null;
  const dailySummary=intent==='summary'?computeDailySummary(rawData):null;
  const ekspReport=intent==='analytics'?computeEkspReport(rawData):null;

  // Build dynamic context
  let context=`Tanggal hari ini (WIB): ${today}\nUser: ${senderName}\n\n`;

  if (instruksi.length>0) {
    context+=`=== ⚡ INSTRUKSI WAJIB DIIKUTI ===\n`;
    instruksi.forEach(m=>{context+=`• ${m[2]}\n`;});
    context+='\n';
  }
  if (memoryOther.length>0) {
    context+=`=== MEMORY ===\n`;
    memoryOther.forEach(m=>{context+=`[${m[1]}] ${m[2]}\n`;});
    context+='\n';
  }
  if (todayReminders.length>0) {
    context+=`=== REMINDER HARI INI ===\n`;
    todayReminders.forEach(r=>{context+=`Oleh: ${r[1]} | Order: ${r[2]||'-'} | ${r[4]}\n`;});
    context+='\n';
  }
  if (upcomingReminders.length>0) {
    context+=`=== UPCOMING REMINDERS ===\n`;
    upcomingReminders.forEach(r=>{context+=`[${r[3]}] Oleh: ${r[1]} | Order: ${r[2]||'-'} | ${r[4]}\n`;});
    context+='\n';
  }
  if (slaAlerts) {
    context+=`=== SLA ALERTS ===\nOverdue:${slaAlerts.overdue.length} Urgent:${slaAlerts.urgent.length} Warning:${slaAlerts.warning.length} Attention:${slaAlerts.attention.length}\n`;
    if(slaAlerts.overdue.length)context+=`Overdue:${JSON.stringify(slaAlerts.overdue)}\n`;
    if(slaAlerts.urgent.length)context+=`Urgent:${JSON.stringify(slaAlerts.urgent)}\n`;
    if(slaAlerts.warning.length)context+=`Warning:${JSON.stringify(slaAlerts.warning)}\n`;
    context+='\n';
  }
  if (dailySummary) context+=`=== DAILY SUMMARY ===\n${JSON.stringify(dailySummary)}\n\n`;
  if (ekspReport) context+=`=== EKSPEDISI REPORT ===\n${JSON.stringify(ekspReport)}\n\n`;
  if (filteredResult&&filteredResult.rows.length>0) context+=`=== DATA ORDER [${intent}, ${filteredResult.rows.length} rows] ===\n${JSON.stringify(filteredResult.rows)}\n`;
  else if (filteredResult) context+=`=== DATA ORDER ===\nTidak ada data sesuai filter.\n`;

  if (!chatHistory[senderId]) chatHistory[senderId]=[];
  const messages=[...chatHistory[senderId]];
  messages.push({role:'user',content:`${context}\n\n${userMessage}`});

  const response=await anthropic.messages.create({
    model:'claude-haiku-4-5-20251001',
    max_tokens:1000,
    system:[{type:'text',text:SYSTEM_PROMPT,cache_control:{type:'ephemeral'}}],
    messages,
  });
  const usage=response.usage;
  console.log(`[Tokens] in:${usage.input_tokens} out:${usage.output_tokens} cache_write:${usage.cache_creation_input_tokens||0} cache_read:${usage.cache_read_input_tokens||0}`);

  const reply=response.content[0].text;
  chatHistory[senderId].push({role:'user',content:userMessage});
  chatHistory[senderId].push({role:'assistant',content:reply});
  if(chatHistory[senderId].length>12)chatHistory[senderId]=chatHistory[senderId].slice(-12);
  await parseActions(reply,senderId,senderName);
  return reply.replace(/ACTION:[^\n]+/g,'').trim();
}

async function parseActions(text, senderId, senderName) {
  for (const line of text.split('\n')) {
    const t=line.trim();
    try {
      if (t.startsWith('ACTION:UPDATE_RESI:')) {
        const rest=t.replace('ACTION:UPDATE_RESI:',''),idx=rest.indexOf(':');
        if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.resi,'Resi',rest.substring(idx+1),senderName);
      }
      if (t.startsWith('ACTION:UPDATE_TGL_KIRIM:')) {
        const rest=t.replace('ACTION:UPDATE_TGL_KIRIM:',''),idx=rest.indexOf(':');
        if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.tglPengiriman,'Tgl Pengiriman',parseDate(rest.substring(idx+1)),senderName);
      }
      if (t.startsWith('ACTION:UPDATE_TGL_TIBA:')) {
        const rest=t.replace('ACTION:UPDATE_TGL_TIBA:',''),idx=rest.indexOf(':');
        if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.tglTiba,'Tgl Tiba',parseDate(rest.substring(idx+1)),senderName);
      }
      if (t.startsWith('ACTION:UPDATE_REMARK:')) {
        const rest=t.replace('ACTION:UPDATE_REMARK:',''),idx=rest.indexOf(':');
        if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.remark,'Remark',rest.substring(idx+1),senderName);
      }
      if (t.startsWith('ACTION:SAVE_MEMORY:')) {
        const rest=t.replace('ACTION:SAVE_MEMORY:',''),idx=rest.indexOf(':');
        if(idx>-1)await saveMemory(rest.substring(0,idx),rest.substring(idx+1));
      }
      if (t.startsWith('ACTION:SAVE_REMINDER:')) {
        const rest=t.replace('ACTION:SAVE_REMINDER:','');
        const parts=rest.split(':');
        if(parts.length>=3){
          const noOrders=parts[0].split(/[,\s]+dan\s+|,\s*|&\s*/).map(s=>s.trim()).filter(Boolean);
          await saveReminder(senderName,noOrders,parseDate(parts[1]),parts.slice(2).join(':'));
        }
      }
    } catch (e) { console.error(`parseActions error: ${e.message}`); }
  }
}

// ─── LOG TODAY ────────────────────────────────────────────
async function getLogToday() {
  await ensureTab(UPDATE_LOG_TAB,['Timestamp','Oleh','No Order','Kolom','Nilai Lama','Nilai Baru']);
  const s=await getSheets(); if(!s) return 'Gagal ambil log.';
  try {
    const res=await s.spreadsheets.values.get({spreadsheetId:SHEET_ID,range:`${UPDATE_LOG_TAB}!A:F`});
    const rows=(res.data.values||[]).slice(1);
    const today=getToday();
    const todayStr=today.split('-').reverse().join('/');
    const todayLogs=rows.filter(r=>(r[0]||'').includes(todayStr.replace(/^0/,'').replace('/0','/')));
    if(!todayLogs.length) return '📋 Belum ada update hari ini.';
    let msg=`📋 LOG UPDATE HARI INI\n\n`;
    todayLogs.forEach(r=>{msg+=`🕐 ${r[0]} | ${r[1]}\n📦 ${r[2]} | ${r[3]}: ${r[4]||'—'} → ${r[5]}\n\n`;});
    return msg;
  } catch (e) { return 'Gagal ambil log: '+e.message; }
}

// ─── MESSAGING ────────────────────────────────────────────
async function sendTelegram(chatId, message) {
  try { await axios.post(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`,{chat_id:chatId,text:message}); }
  catch (e) { console.error('sendTelegram error:', e.response?.data||e.message); }
}

async function sendWA(target, message) {
  try { await axios.post('https://api.fonnte.com/send',{target,message},{headers:{Authorization:FONNTE_TOKEN}}); }
  catch (e) { console.error('sendWA error:', e.message); }
}

async function getTelegramFile(fileId) {
  const fileRes=await axios.get(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
  const fileUrl=`https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${fileRes.data.result.file_path}`;
  const res=await axios.get(fileUrl,{responseType:'arraybuffer'});
  return Buffer.from(res.data);
}

// ─── GROUP WA HANDLER ─────────────────────────────────────
async function handleGroupMessage(groupId, rawMessage, senderName) {
  // Strip @Nyenyenye dari pesan
  const cleanMsg = rawMessage.replace(new RegExp(`@${BOT_MENTION}`,'gi'),'').trim();
  console.log(`Group [${senderName}]: ${cleanMsg.substring(0,80)}`);

  // Update tgl kirim: update tgl kirim 9228972 12/05/2026
  const updateMatch = cleanMsg.match(/update\s+tgl\s+kirim\s+(\d+)\s+(.+)/i);
  if (updateMatch) {
    const noOrder=updateMatch[1], tanggal=parseDate(updateMatch[2].trim());
    const ok=await updateOrderField(noOrder,COL.tglPengiriman,'Tgl Pengiriman',tanggal,senderName);
    const reply=ok
      ? `✅ Tgl Kirim order *${noOrder}* → *${tanggal}*\nOleh: ${senderName}`
      : `⚠️ Order *${noOrder}* tidak ditemukan di sheet.`;
    await sendWA(groupId, reply);
    return;
  }

  // Cek order: cek order 9228972
  const cekMatch = cleanMsg.match(/cek\s+order\s+(\d+)/i);
  if (cekMatch) {
    const noOrder=cekMatch[1];
    const row=await getOrderByNumber(noOrder);
    if (row) {
      let reply=`📦 ORDER ${row[COL.noOrder]} - ${row[COL.customer]}\n`;
      reply+=`📍 ${row[COL.kota]} | 🚚 ${row[COL.ekspedisi]||'—'}\n`;
      reply+=`🚢 Shipping: ${row[COL.shippingNum]||'—'}\n`;
      reply+=`🔖 Resi: ${row[COL.resi]||'Kosong ⚠️'}\n`;
      reply+=`📅 Tgl Kirim: ${row[COL.tglPengiriman]||'—'}\n`;
      reply+=`📅 Tgl Tiba: ${row[COL.tglTiba]||'—'}\n`;
      reply+=`📊 Status: ${row[COL.status]||'—'}\n`;
      reply+=`⏱ SLA: ${row[COL.sla]||'—'}`;
      await sendWA(groupId, reply);
    } else {
      await sendWA(groupId, `⚠️ Order *${noOrder}* tidak ditemukan.`);
    }
    return;
  }

  // Reminder inline: reminder 9228972 15/05/2026 follow up
  const reminderMatch = cleanMsg.match(/reminder\s+(\d+)\s+(\S+)\s+(.+)/i);
  if (reminderMatch) {
    const noOrder=reminderMatch[1], tanggal=parseDate(reminderMatch[2]), note=reminderMatch[3];
    await saveReminder(senderName,[noOrder],tanggal,note);
    await sendWA(groupId, `📌 Reminder disimpan!\n📦 Order: ${noOrder}\n📅 Tanggal: ${tanggal}\n📌 Note: ${note}`);
    return;
  }

  // Fallback: info cara pakai
  await sendWA(groupId,
    `Halo! Cara pakai @${BOT_MENTION} di grup:\n\n` +
    `• *Cek order:* @${BOT_MENTION} cek order [no_order]\n` +
    `• *Update tgl kirim:* @${BOT_MENTION} update tgl kirim [no_order] [tanggal]\n` +
    `• *Reminder:* @${BOT_MENTION} reminder [no_order] [tanggal] [note]`
  );
}

// ─── WEBHOOKS ─────────────────────────────────────────────
app.get('/health',(_, res)=>res.json({status:'ok',model:'claude-haiku-4-5-20251001',time_wib:new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'})}));
app.get('/webhook/telegram',(_, res)=>res.sendStatus(200));
app.get('/webhook/wa',(_, res)=>res.sendStatus(200));

// Telegram
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  const msg=req.body.message||req.body.edited_message; if(!msg) return;
  const chatId=msg.chat.id;
  const firstName=msg.from?.first_name||msg.chat?.first_name||'User';
  const senderId=`tg_${chatId}`;
  let text=msg.text||msg.caption||'', imageBase64=null;
  if (msg.document) {
    const fname=msg.document.file_name||'';
    if (fname.match(/\.(xlsx|xls)$/i)) {
      try { const buf=await getTelegramFile(msg.document.file_id); await sendTelegram(chatId,await handleExcelUpload(buf,firstName,senderId,false)); }
      catch(e){ await sendTelegram(chatId,'Gagal baca Excel: '+e.message); }
      return;
    }
  }
  if (msg.photo&&msg.photo.length>0) {
    try { const buf=await getTelegramFile(msg.photo[msg.photo.length-1].file_id); imageBase64=buf.toString('base64'); if(!text)text='Tolong baca nomor resi dari foto ini.'; }
    catch(e){ console.error('Photo error:', e.message); }
  }
  if (!text&&!imageBase64) return;
  console.log(`TG [${firstName}]: ${text.substring(0,80)}`);
  try { await sendTelegram(chatId,await callClaude(senderId,firstName,text,imageBase64)); }
  catch(e){ console.error('TG error:', e.message); await sendTelegram(chatId,'Maaf, error: '+e.message); }
});

// WhatsApp
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const {sender,message,name,file,mimetype,member}=req.body;
  if (!sender) return;
  const isGroup=sender.includes('@g.us');
  const senderName=name||member||sender;
  const senderId=`wa_${sender}`;
  console.log(`WA [${senderName}${isGroup?'/GROUP':''}]: "${(message||'').substring(0,60)}" file=${!!file}`);

  try {
    // GROUP MESSAGE
    if (isGroup) {
      if (!message) return;
      // Deteksi mention: via teks @Nyenyenye ATAU via WA proper mention @628xxx
      const isMentioned = message.toLowerCase().includes(`@${BOT_MENTION}`) ||
                          (BOT_WA_NUMBER && message.includes(`@${BOT_WA_NUMBER}`));
      if (!isMentioned) return;
      await handleGroupMessage(sender, message, senderName);
      return;
    }

    // PRIVATE - Excel
    if (file&&mimetype&&/spreadsheet|excel|xlsx|xls/i.test(mimetype)) {
      const buf=await getFonnteFile(file);
      if (!buf){ await sendWA(sender,'Gagal download file.'); return; }
      await sendWA(sender,await handleExcelUpload(buf,senderName,senderId,true));
      return;
    }

    // PRIVATE - Foto
    if (file&&mimetype&&/image/i.test(mimetype)) {
      const buf=await getFonnteFile(file);
      if (!buf){ await sendWA(sender,'Gagal download foto.'); return; }
      const imgBase64=buf.toString('base64');
      const caption=message||'Tolong baca nomor resi dari foto ini.';
      await sendWA(sender,await callClaude(senderId,senderName,caption,imgBase64,mimetype));
      return;
    }

    // PRIVATE - Teks
    if (!message) return;
    await sendWA(sender,await callClaude(senderId,senderName,message));

  } catch(e){ console.error('WA error:', e.message); await sendWA(sender,'Maaf, error: '+e.message); }
});

// ─── SCHEDULERS ───────────────────────────────────────────

// 08:00 WIB — H-1 Pending + SLA Alert harian
cron.schedule('0 8 * * *', async () => {
  console.log('=== SCHEDULER 08:00 WIB ===');
  try {
    clearCache('sheetData');
    const data=await getSheetData();
    const today=getToday();
    const tomorrow=new Date(); tomorrow.setDate(tomorrow.getDate()+1);
    const tomorrowStr=tomorrow.toLocaleDateString('sv-SE',{timeZone:'Asia/Jakarta'});

    if (data) {
      // H-1 Pending Request
      const h1=data.slice(1).filter(r=>(r[COL.requestDate]||'').trim()===tomorrowStr&&(r[COL.status]||'').toLowerCase()!=='received');
      if (h1.length>0) {
        let msg=`🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
        h1.forEach(r=>{ msg+=`📦 ${r[COL.noOrder]} — ${r[COL.customer]} (${r[COL.ekspedisi]||'—'})\n`; msg+=r[COL.resi]?`  ✅ Resi: ${r[COL.resi]}\n`:`  ❌ Resi belum diinput!\n`; });
        msg+=`\nTotal: ${h1.length} order`;
        await sendToTargets(msg);
      }

      // SLA Alert
      const sla=computeSLAAlerts(data);
      if (sla&&(sla.urgent.length>0||sla.overdue.length>0)) {
        let msg=`⚠️ ALERT SLA - ${today}\n\n`;
        if(sla.overdue.length){msg+=`🚨 OVERDUE (${sla.overdue.length}):\n`;sla.overdue.forEach(o=>{msg+=`• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`;});msg+='\n';}
        if(sla.urgent.length){msg+=`🔴 URGENT H-0 (${sla.urgent.length}):\n`;sla.urgent.forEach(o=>{msg+=`• ${o.no_order} - ${o.customer}\n`;});msg+='\n';}
        if(sla.warning.length){msg+=`🟡 WARNING H-1 (${sla.warning.length}):\n`;sla.warning.forEach(o=>{msg+=`• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`;}); }
        await sendToTargets(msg);
      }
    }
  } catch(e){ console.error('Scheduler 08:00 error:', e.message); }
}, {timezone:'Asia/Jakarta'});

// 17:00 WIB — Reminder H-1 (besok)
cron.schedule('0 17 * * *', async () => {
  console.log('=== SCHEDULER 17:00 WIB — Reminder H-1 ===');
  try {
    clearCache('reminders');
    const reminders=await getReminders();
    const tomorrow=new Date(); tomorrow.setDate(tomorrow.getDate()+1);
    const tomorrowStr=tomorrow.toLocaleDateString('sv-SE',{timeZone:'Asia/Jakarta'});
    const h1Reminders=reminders.filter(r=>r[3]===tomorrowStr&&r[5]==='Pending');
    for (let i=0; i<h1Reminders.length; i++) {
      const msg=await buildReminderMsg(h1Reminders[i],`⏰ REMINDER BESOK`);
      await sendToTargets(msg);
    }
    console.log(`H-1 reminders sent: ${h1Reminders.length}`);
  } catch(e){ console.error('Scheduler 17:00 error:', e.message); }
}, {timezone:'Asia/Jakarta'});

// 07:00 WIB — Reminder H-0 (hari ini)
cron.schedule('0 7 * * *', async () => {
  console.log('=== SCHEDULER 07:00 WIB — Reminder H-0 ===');
  try {
    clearCache('reminders');
    const reminders=await getReminders();
    const today=getToday();
    const todayReminders=reminders.filter(r=>r[3]===today&&r[5]==='Pending');
    for (let i=0; i<todayReminders.length; i++) {
      const msg=await buildReminderMsg(todayReminders[i],`🔔 REMINDER HARI INI`);
      await sendToTargets(msg);
      await markReminderDone(reminders.indexOf(todayReminders[i]));
    }
    console.log(`H-0 reminders sent: ${todayReminders.length}`);
  } catch(e){ console.error('Scheduler 07:00 error:', e.message); }
}, {timezone:'Asia/Jakarta'});

// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS Agent (Claude Haiku) running on port ${PORT}`);
  console.log(`Group WA: ${GROUP_WA_ID||'NOT SET'} | Mention: @${BOT_MENTION}`);
  console.log(`Reminder targets: ${REMINDER_TARGETS.length} numbers`);
});
