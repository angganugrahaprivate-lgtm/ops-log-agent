require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');
const XLSX = require('xlsx');
const { wrapper }   = require('axios-cookiejar-support');
const { CookieJar } = require('tough-cookie');
const cheerio       = require('cheerio');

const app = express();
app.use(express.json());

// ─── CONFIG ───────────────────────────────────────────────
const ANTHROPIC_KEY    = process.env.ANTHROPIC_API_KEY;
const SHEET_ID         = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB        = process.env.SHEET_TAB_NAME || 'Data Handover';
const MEMORY_TAB       = 'Memory';
const REMINDER_TAB     = 'Reminders';
const UPDATE_LOG_TAB   = 'Update Log';
const FONNTE_TOKEN     = process.env.FONNTE_TOKEN;
const TELEGRAM_TOKEN   = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_NOTIF_CHAT_ID;
const YOUR_WA_NUMBER   = process.env.YOUR_WA_NUMBER;
const GROUP_WA_ID      = process.env.GROUP_WA_ID;
const BOT_MENTION      = (process.env.BOT_MENTION_NAME || 'Nyenyenye').toLowerCase();
const BOT_WA_NUMBER    = process.env.BOT_WA_NUMBER || '';
const REMINDER_TARGETS = (process.env.REMINDER_TARGETS || '').split(',').filter(Boolean);
const PORT             = process.env.PORT || 3000;

console.log('=== ENV CHECK ===');
console.log('ANTHROPIC_API_KEY:', ANTHROPIC_KEY ? 'OK' : 'MISSING');
console.log('GOOGLE_SHEETS_ID:', SHEET_ID ? 'OK' : 'MISSING');
console.log('GROUP_WA_ID:', GROUP_WA_ID || 'NOT SET');
console.log('BOT_MENTION:', BOT_MENTION);
console.log('BOT_WA_NUMBER:', BOT_WA_NUMBER || 'NOT SET');
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
function getCache(key) { const c = cache[key]; if (c && Date.now() - c.time < CACHE_TTL) return c.data; return null; }
function setCache(key, data) { cache[key] = { data, time: Date.now() }; }
function clearCache(key) { if (key) delete cache[key]; else Object.keys(cache).forEach(k => delete cache[k]); }

// ─── GOOGLE SHEETS ────────────────────────────────────────
let sheetsClient = null;
let sheetsInitPromise = null;

async function getSheets() {
  if (sheetsClient) return sheetsClient;
  if (!sheetsInitPromise) {
    sheetsInitPromise = (async () => {
      try {
        const { google } = require('googleapis');
        const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        sheetsClient = google.sheets({ version: 'v4', auth });
        console.log('Google Sheets OK');
        return sheetsClient;
      } catch (e) {
        sheetsInitPromise = null; // reset agar bisa retry
        console.error('Sheets init error:', e.message);
        return null;
      }
    })();
  }
  return sheetsInitPromise;
}

function toCol(n) { let c = ''; while (n >= 0) { c = String.fromCharCode(65 + (n % 26)) + c; n = Math.floor(n / 26) - 1; } return c; }
function getToday() { return new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' }); }
function getCutOff() { const h = parseInt(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta', hour: 'numeric', hour12: false })); return h < 14 ? '1' : '2'; }

function parseDate(s) {
  if (!s) return '';
  s = s.toString().trim();
  const mon = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,mei:5,maret:3,april:4,juni:6,juli:7,agustus:8,oktober:10,november:11,desember:12 };
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if (m) return s;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  m = s.match(/(\d{1,2})\s*[-\s]+\s*([a-zA-Z]+)\s*[-\s]+\s*(\d{4})/);
  if (m) { const mn = mon[m[2].toLowerCase().trim().substring(0,3)]; if (mn) return `${m[3]}-${String(mn).padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
  return s;
}

function parseDateFromCommand(command) {
  const today = getToday();
  const cmd = (command || '').toLowerCase();
  const addDays = (base, n) => { const d = new Date(base + 'T00:00:00'); d.setDate(d.getDate() + n); return d.toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' }); };
  if (cmd.includes('hari ini')) return today;
  if (cmd.includes('besok')) return addDays(today, 1);
  if (cmd.includes('lusa')) return addDays(today, 2);
  if (cmd.includes('kemarin')) return addDays(today, -1);
  const mon = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,mei:5,maret:3,april:4,juni:6,juli:7,agustus:8,oktober:10,november:11,desember:12 };
  let m = cmd.match(/(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{4}))?/);
  if (m) { const yr = m[3] || new Date().getFullYear(); return `${yr}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
  m = cmd.match(/(\d{1,2})\s+([a-z]+)(?:\s+(\d{4}))?/);
  if (m) { const mn = mon[m[2].substring(0,3)]; if (mn) { const yr = m[3] || new Date().getFullYear(); return `${yr}-${String(mn).padStart(2,'0')}-${m[1].padStart(2,'0')}`; } }
  return today;
}

function formatDateID(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('id-ID', { day:'2-digit', month:'long', year:'numeric', timeZone:'Asia/Jakarta' });
}

async function getSheetData() {
  const cached = getCache('sheetData');
  if (cached) { console.log('Cache hit: sheetData'); return cached; }
  const s = await getSheets(); if (!s) return null;
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!A:AK` });
    const data = res.data.values || [];
    setCache('sheetData', data);
    console.log(`Sheet loaded: ${data.length} rows`);
    return data;
  } catch (e) { console.error('getSheetData error:', e.message); return null; }
}

function selectColumns(rows, colIndices, colNames) {
  return rows.map(row => { const obj = {}; colIndices.forEach((idx, i) => { obj[colNames[i]] = (row[idx] || ''); }); return obj; });
}

async function updateCell(sheetTab, rowNum, colIdx, value) {
  const s = await getSheets(); if (!s) return;
  await s.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${sheetTab}!${toCol(colIdx)}${rowNum}`, valueInputOption: 'RAW', requestBody: { values: [[value]] } });
  clearCache('sheetData');
}

async function ensureTab(tabName, headers) {
  const s = await getSheets(); if (!s) return;
  try { await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${tabName}!A1` }); }
  catch (e) {
    try {
      const { google } = require('googleapis');
      const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
      const api = google.sheets({ version: 'v4', auth });
      await api.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] } });
      await api.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${tabName}!A1`, valueInputOption: 'RAW', requestBody: { values: [headers] } });
      console.log(`Tab "${tabName}" created`);
    } catch (e2) { console.error('ensureTab error:', e2.message); }
  }
}

// ─── UPDATE LOG ───────────────────────────────────────────
async function logUpdate(updatedBy, noOrder, kolom, nilaiLama, nilaiBaru) {
  await ensureTab(UPDATE_LOG_TAB, ['Timestamp','Oleh','No Order','Kolom','Nilai Lama','Nilai Baru']);
  const s = await getSheets(); if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${UPDATE_LOG_TAB}!A:F`, valueInputOption: 'RAW', requestBody: { values: [[ts, updatedBy, noOrder, kolom, nilaiLama, nilaiBaru]] } });
  } catch (e) { console.error('logUpdate error:', e.message); }
}

// ─── UPDATE FUNCTIONS ─────────────────────────────────────
async function updateOrderField(noOrder, colIdx, colName, newValue, updatedBy) {
  const data = await getSheetData(); if (!data) return false;
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL.noOrder] || '').trim() === noOrder.trim()) {
      const oldValue = data[i][colIdx] || '';
      await updateCell(SHEET_TAB, i + 1, colIdx, newValue);
      await logUpdate(updatedBy, noOrder, colName, oldValue, newValue);
      return true;
    }
  }
  return false;
}

async function getOrderByNumber(noOrder) {
  const data = await getSheetData(); if (!data || data.length < 2) return null;
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL.noOrder] || '').includes(noOrder) || (data[i][COL.shippingNum] || '').includes(noOrder)) return data[i];
  }
  return null;
}

// ─── MEMORY ───────────────────────────────────────────────
async function getMemory() {
  const cached = getCache('memory'); if (cached) return cached;
  await ensureTab(MEMORY_TAB, ['Timestamp','Category','Content']);
  const s = await getSheets(); if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${MEMORY_TAB}!A:C` });
    const data = (res.data.values || []).slice(1).slice(-50);
    setCache('memory', data);
    return data;
  } catch (e) { return []; }
}

async function saveMemory(category, content) {
  await ensureTab(MEMORY_TAB, ['Timestamp','Category','Content']);
  const s = await getSheets(); if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${MEMORY_TAB}!A:C`, valueInputOption: 'RAW', requestBody: { values: [[ts, category, content]] } });
    clearCache('memory');
  } catch (e) { console.error('saveMemory error:', e.message); }
}

// ─── REMINDER ─────────────────────────────────────────────
async function getReminders() {
  const cached = getCache('reminders'); if (cached) return cached;
  await ensureTab(REMINDER_TAB, ['Timestamp','Dibuat Oleh','No Order','Tanggal Reminder','Note','Status']);
  const s = await getSheets(); if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:F` });
    const data = (res.data.values || []).slice(1);
    setCache('reminders', data);
    return data;
  } catch (e) { return []; }
}

async function saveReminder(createdBy, noOrders, tanggal, note) {
  const noOrderStr = Array.isArray(noOrders) ? noOrders.join(',') : noOrders;
  await ensureTab(REMINDER_TAB, ['Timestamp','Dibuat Oleh','No Order','Tanggal Reminder','Note','Status']);
  const s = await getSheets(); if (!s) return false;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:F`, valueInputOption: 'RAW', requestBody: { values: [[ts, createdBy, noOrderStr || '', tanggal, note, 'Pending']] } });
    clearCache('reminders');
    return true;
  } catch (e) { console.error('saveReminder error:', e.message); return false; }
}

async function markReminderDone(rowIndex) {
  const s = await getSheets(); if (!s) return;
  try { await updateCell(REMINDER_TAB, rowIndex + 2, 5, 'Sent'); clearCache('reminders'); }
  catch (e) { console.error('markReminderDone error:', e.message); }
}

async function buildReminderMsg(reminder, label) {
  const [, createdBy, noOrderStr, tanggal, note] = reminder;
  const noOrders = (noOrderStr || '').split(',').map(s => s.trim()).filter(Boolean);
  const data = await getSheetData();
  const orders = [];
  if (data && noOrders.length > 0) {
    for (const no of noOrders) {
      const row = data.slice(1).find(r => (r[COL.noOrder] || '').includes(no) || (r[COL.shippingNum] || '').includes(no));
      if (row) orders.push(row);
    }
  }
  const join = (field) => orders.length > 0 ? orders.map(r => r[field] || '—').filter((v, i, a) => a.indexOf(v) === i).join(' & ') : '—';
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

async function sendToTargets(message) {
  const targets = [...REMINDER_TARGETS];
  if (YOUR_WA_NUMBER && !targets.includes(YOUR_WA_NUMBER)) targets.push(YOUR_WA_NUMBER);
  for (const num of targets) await sendWA(num, message);
  if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, message);
}

// ─── LIST PENGIRIMAN ──────────────────────────────────────
async function buildListPengiriman(command, filterPalembang = false) {
  const targetDate = parseDateFromCommand(command);
  const data = await getSheetData();
  if (!data || data.length < 2) return { msg1: 'Tidak ada data.', msg2: null };

  const rows = data.slice(1).filter(r => {
    const tgl = (r[COL.tglPengiriman] || '').toString().trim();
    if (!tgl.startsWith(targetDate)) return false;
    if (filterPalembang && !(r[COL.kota] || '').toLowerCase().includes('palembang')) return false;
    return true;
  });

  const dateLabel = formatDateID(targetDate);
  if (rows.length === 0) return { msg1: `Tidak ada pengiriman pada ${dateLabel}.`, msg2: null };

  const byEksp = {};
  rows.forEach(r => {
    const eksp = r[COL.ekspedisi] || 'Lainnya';
    if (!byEksp[eksp]) byEksp[eksp] = [];
    byEksp[eksp].push(r);
  });

  let msg1 = filterPalembang
    ? `📦 LIST PENGIRIMAN PALEMBANG - ${dateLabel}\n`
    : `📦 LIST PENGIRIMAN - ${dateLabel}\n`;
  msg1 += `Total: ${rows.length} order\n`;

  for (const [eksp, orders] of Object.entries(byEksp)) {
    msg1 += `\n━━ 🚚 ${eksp} (${orders.length} order) ━━\n`;
    orders.forEach(r => {
      msg1 += `• ${r[COL.noOrder] || '—'} | ${r[COL.customer] || '—'}\n`;
      msg1 += `  📍 ${r[COL.alamat] || '—'}\n`;
      const rmk = (r[COL.remark] || '').trim();
      msg1 += `  📝 ${rmk || '-'}\n`;
    });
  }

  let msg2 = `📋 No. Shipping - ${dateLabel}\n`;
  rows.forEach(r => { msg2 += `${r[COL.shippingNum] || '—'}\n`; });

  return { msg1: msg1.trim(), msg2: msg2.trim() };
}

// ─── EXCEL UPLOAD ─────────────────────────────────────────
async function handleExcelUpload(fileBuffer, senderName, chatId, isWA = false) {
  try {
    const wb = XLSX.read(fileBuffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 2) return 'File kosong atau tidak ada data.';
    const headers = rows[0];
    const dataRows = rows.slice(1).filter(r => r.some(c => c));
    const numIdx    = headers.indexOf('Number');
    const koliIdx   = headers.indexOf('Koli');
    const barangIdx = headers.indexOf('Nama Barang');
    const partnerIdx= headers.indexOf('Partner');
    const nameIdx   = headers.findIndex(h => h && h.toString().includes('First Name'));
    const alamatIdx = headers.findIndex(h => h && h.toString().includes('Line1'));
    const kecIdx    = headers.findIndex(h => h && h.toString().includes('Line3'));
    const kotaIdx   = headers.findIndex(h => h && h.toString().includes('Line4'));
    const phoneIdx  = headers.findIndex(h => h && h.toString().includes('Phone'));
    const posIdx    = headers.findIndex(h => h && h.toString().includes('Postcode'));
    const stateIdx  = headers.findIndex(h => h && h.toString().includes('State'));
    const today = getToday(), cutoff = getCutOff();
    let preview = `📋 PREVIEW DATA EXCEL\nTanggal: ${today} | Cut Off: ${cutoff}\nTotal: ${dataRows.length} order\n\n`;
    dataRows.forEach((r, i) => { preview += `${i+1}. ${r[numIdx]||'—'} - ${r[nameIdx]||'—'} - ${r[partnerIdx]||'—'}\n`; });
    preview += `\nMasukkan semua ke GSheet? Balas "ya" untuk konfirmasi.`;
    cache[`pendingExcel_${chatId}`] = { data:dataRows,time:Date.now(),today,cutoff,numIdx,koliIdx,barangIdx,partnerIdx,nameIdx,alamatIdx,kecIdx,kotaIdx,phoneIdx,posIdx,stateIdx,senderName,chatId,isWA };
    return preview;
  } catch (e) { console.error('handleExcelUpload error:', e.message); return 'Gagal membaca file Excel.'; }
}

async function insertExcelToSheet(senderName, chatId) {
  const pendingKey = `pendingExcel_${chatId}`;
  const pending = cache[pendingKey];
  if (!pending) return 'Tidak ada data Excel yang menunggu konfirmasi.';
  const s = await getSheets(); if (!s) return 'Gagal konek ke GSheet.';
  try {
    const { data,today,cutoff,numIdx,koliIdx,barangIdx,partnerIdx,nameIdx,alamatIdx,kecIdx,kotaIdx,phoneIdx,posIdx,stateIdx } = pending;
    const rows = data.map(r => {
      const row = Array(36).fill('');
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
    await logUpdate(senderName, 'BULK INSERT', 'Excel Upload', '-', `${rows.length} order baru`);
    clearCache('sheetData'); delete cache[pendingKey];
    return `✅ ${rows.length} order berhasil dimasukkan!\nTanggal: ${today} | Cut Off: ${cutoff}`;
  } catch (e) { return 'Gagal insert: ' + e.message; }
}

async function getFonnteFile(fileUrl) {
  try {
    const res = await axios.get(fileUrl, { responseType:'arraybuffer', headers:{Authorization:FONNTE_TOKEN} });
    return Buffer.from(res.data);
  } catch (e) { console.error('getFonnteFile error:', e.message); return null; }
}

// ─── SMART FILTER ─────────────────────────────────────────
const chatHistory = {};
const HISTORY_MAX_SENDERS = 200;
const NO_DATA_INTENTS = ['greeting','help','out_of_scope','save_instruction'];

function detectIntent(message, senderId) {
  const msg = (message || '').toLowerCase().trim();
  if (/^(halo|hai|hi|hello|selamat|pagi|siang|sore|malam|hey)/.test(msg)) return 'greeting';
  if (/^(oke|ok|tidak|ga|gak|siap|done|sip|noted|thanks|makasih)$/.test(msg)) return 'greeting';
  if (/kamu bisa|fitur apa|help|bantuan/.test(msg)) return 'help';
  if (/^(catat|ingat|note)\s*:/i.test(msg)) return 'save_instruction';
  if (/dry.?run|test.?jt|test.?report/.test(msg)) return 'dryrun_jt';  // ← PATCH: dry run intent
  if (/dry.?run.*sentral|test.*sentral|sentral.*dry.?run/.test(msg)) return 'dryrun_sentral';
  if (/refresh|sync data|reload data/.test(msg)) return 'refresh';
  if (/log hari ini|history update|apa.*diupdate/.test(msg)) return 'log_today';
  if (/list pengiriman|daftar pengiriman|pengiriman (hari ini|besok|lusa|kemarin|\d)/.test(msg)) return 'list_pengiriman';
  const hasPending = senderId && cache[`pendingExcel_${senderId}`];
  if (/^(ya|yes|iya|yep|yup|konfirmasi|insert|masukkan)$/.test(msg) && hasPending) return 'confirm_excel';
  if (/upload.*(excel|xlsx|file)|kirim.*(excel|file)/.test(msg)) return 'prompt_excel';
  if (/no order\s*:/i.test(msg) && /tgl kirim\s*:/i.test(msg)) return 'format1_update';
  if (/no order\s*:/i.test(msg) && /remark\s*:/i.test(msg)) return 'format2_remark';
  if (/^reminder/i.test(msg) && /no order\s*:/i.test(msg) && /tgl\s*:/i.test(msg)) return 'format3_reminder';
  if (/update resi|input resi|tambah resi/.test(msg)) return 'update_resi';
  if (/update.*tgl.*kirim|update.*tanggal.*kirim|tgl pengiriman/.test(msg)) return 'update_tgl_kirim';
  if (/update.*tgl.*tiba|update.*tanggal.*tiba|sudah tiba/.test(msg)) return 'update_tgl_tiba';
  if (/update.*remark|tambah.*remark|isi.*remark/.test(msg)) return 'update_remark';
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
  if (/nama|customer|cari.*nama/.test(msg)) return 'customer_search';
  if (senderId && chatHistory[senderId] && chatHistory[senderId].length > 0 && msg.length < 80) return 'general';
  return 'out_of_scope';
}

function filterData(data, intent, message) {
  if (!data || data.length < 2) return null;
  const rows = data.slice(1);
  const today = getToday();
  const UPDATE_INTENTS = ['update_resi','update_tgl_kirim','update_tgl_tiba','update_remark','format1_update','format2_remark'];
  if (UPDATE_INTENTS.includes(intent)) {
    const filtered = rows.filter(r => (r[COL.status] || '').toLowerCase() !== 'received');
    return { names: UPDATE_NAMES, rows: selectColumns(filtered, UPDATE_COLS, UPDATE_NAMES) };
  }
  let filtered;
  switch (intent) {
    case 'pending': filtered = rows.filter(r => /pending|waiting/i.test(r[COL.status] || '')); break;
    case 'no_resi': filtered = rows.filter(r => !(r[COL.resi] || '').trim() && (r[COL.status] || '').toLowerCase() !== 'received'); break;
    case 'overdue': filtered = rows.filter(r => { const s = r[COL.sla]; return s && s < today && (r[COL.status] || '').toLowerCase() !== 'received'; }); break;
    case 'sla_alert': filtered = rows.filter(r => { const s = r[COL.sla]; if (!s || (r[COL.status] || '').toLowerCase() === 'received') return false; const d = Math.round((new Date(s) - new Date(today)) / 86400000); return d >= -1 && d <= 2; }); break;
    case 'shipped_today': filtered = rows.filter(r => (r[COL.tglPengiriman] || '').startsWith(today)); break;
    case 'specific_order': const on = (message.match(/\d{6,}/) || [])[0]; filtered = on ? rows.filter(r => (r[COL.noOrder] || '').includes(on) || (r[COL.shippingNum] || '').includes(on)) : rows.slice(-50); break;
    case 'customer_search': const ws = message.split(/\s+/).filter(w => w.length > 3); filtered = rows.filter(r => ws.some(w => (r[COL.customer] || '').toLowerCase().includes(w.toLowerCase()))); break;
    case 'ekspedisi': const ek = ['jne','j&t','jnt','sicepat','anteraja','ninja','tiki','lion','jnl','deliveree','sentral'].find(k => message.toLowerCase().includes(k)); filtered = ek ? rows.filter(r => (r[COL.ekspedisi] || '').toLowerCase().includes(ek)) : rows.slice(-100); break;
    case 'reminder': case 'list_reminder': case 'format3_reminder': return null;
    default: filtered = rows.slice(-200); break;
  }
  console.log(`Filter [${intent}]: ${filtered.length}/${rows.length} rows`);
  return { names: GENERAL_NAMES, rows: selectColumns(filtered, GENERAL_COLS, GENERAL_NAMES) };
}

// ─── MONITORING ───────────────────────────────────────────
function computeSLAAlerts(data) {
  if (!data || data.length < 2) return null;
  const today = getToday();
  const result = { urgent:[], warning:[], attention:[], overdue:[] };
  data.slice(1).forEach(row => {
    const status = (row[COL.status] || '').toLowerCase();
    if (status === 'received') return;
    const slaStr = row[COL.sla]; if (!slaStr) return;
    const diff = Math.round((new Date(slaStr) - new Date(today)) / 86400000);
    const info = { no_order:row[COL.noOrder],customer:row[COL.customer],kota:row[COL.kota],status:row[COL.status],sla:slaStr,eksp:row[COL.ekspedisi],resi:row[COL.resi],diff };
    if (diff < 0) result.overdue.push(info);
    else if (diff === 0) result.urgent.push(info);
    else if (diff === 1) result.warning.push(info);
    else if (diff <= 2) result.attention.push(info);
  });
  return result;
}

function computeDailySummary(data) {
  if (!data || data.length < 2) return null;
  const today = getToday(), rows = data.slice(1);
  const kotaCount = {};
  rows.filter(r => (r[COL.status] || '').toLowerCase() !== 'received').forEach(r => { const k = r[COL.kota] || 'Unknown'; kotaCount[k] = (kotaCount[k] || 0) + 1; });
  return {
    receivedToday: rows.filter(r => (r[COL.tanggal] || '').startsWith(today)).length,
    shippedToday: rows.filter(r => (r[COL.tglPengiriman] || '').startsWith(today)).length,
    waiting: rows.filter(r => /waiting|pending/i.test(r[COL.status] || '')).length,
    noResi: rows.filter(r => !(r[COL.resi] || '').trim() && (r[COL.status] || '').toLowerCase() !== 'received').length,
    overdue: rows.filter(r => { const s = r[COL.sla]; return s && s < today && (r[COL.status] || '').toLowerCase() !== 'received'; }).length,
    topKota: Object.entries(kotaCount).sort((a, b) => b[1] - a[1]).slice(0, 3),
  };
}

function computeEkspReport(data) {
  if (!data || data.length < 2) return null;
  const today = getToday(), report = {};
  data.slice(1).forEach(row => {
    const eksp = row[COL.ekspedisi] || 'Unknown';
    if (!report[eksp]) report[eksp] = { total:0, ontime:0, late:0, pending:0 };
    report[eksp].total++;
    const status = (row[COL.status] || '').toLowerCase(), sla = row[COL.sla] || '';
    if (status === 'received') { (sla && row[COL.tglTiba] && row[COL.tglTiba] <= sla) ? report[eksp].ontime++ : report[eksp].late++; }
    else if (sla && sla < today) report[eksp].late++;
    else report[eksp].pending++;
  });
  return report;
}

// ─── SYSTEM PROMPT ────────────────────────────────────────
const SYSTEM_PROMPT = `
Kamu adalah Nyenyenye — asisten operasional logistik Warehouse Palembang.
Bahasa: Indonesia casual. Nada: Santai, friendly, tapi tetap profesional.
Kamu punya akses ke data order, tracking resi, dan bisa update sheet langsung.

## KEPRIBADIAN
- Ngobrol ringan itu oke, tapi tetap fokus ke konteks ops logistik
- Kalau ditanya kondisi hari ini, kasih summary singkat yang informatif
- Proaktif: kalau lihat ada yang perlu diperhatikan, sebutin
- Kalau butuh data atau tracking → gunakan tools yang tersedia
- Jangan bilang "saya tidak bisa" kalau ada tool yang bisa bantu

## KEMAMPUAN (via tools)
- Cek & monitor order: summary, SLA, overdue, per ekspedisi
- Tracking resi J&T Cargo realtime
- Tracking resi Sentral Cargo realtime
- Update data: resi, tgl kirim, tgl tiba, remark, status
- Set reminder, simpan instruksi
- Cari order by nama customer atau nomor

## FORMAT UPDATE (tanpa tool, langsung parse)

### Format 1 — Update JNL/Palembang
Field: HARI, TGL, JAM, NO ORDER, EKSPEDISI, Driver, TGL KIRIM, TGL TIBA, KET
→ ACTION:UPDATE_RESI:[no_order]:[driver]
→ ACTION:UPDATE_TGL_KIRIM:[no_order]:[YYYY-MM-DD]
→ ACTION:UPDATE_TGL_TIBA:[no_order]:[YYYY-MM-DD]
Konversi: "12 - May - 2026" → "2026-05-12"

### Format 2 — Remark
NO ORDER + REMARK → ACTION:UPDATE_REMARK:[no_order]:[remark]

### Format 3 — Reminder
→ ACTION:SAVE_REMINDER:[no_order1,no_order2]:[YYYY-MM-DD]:[note]

### Memory
"Catat/Ingat: ..." → ACTION:SAVE_MEMORY:instruksi:[content]

## ATURAN
- ACTION hanya di akhir pesan, tidak ditampilkan ke user
- Kalau di luar konteks ops logistik → "Hmmm gatau sih, diluar konteks itu keknya 😅"
- Konfirmasi sebelum update kecuali Format 1,2,3 yang sudah jelas
`.trim();

// ─── TOOL DEFINITIONS ─────────────────────────────────────
const TOOLS = [
  {
    name: 'get_order_data',
    description: 'Ambil data order dari Google Sheet. Bisa filter by intent: summary, overdue, sla_alert, no_resi, pending, shipped_today, specific_order, customer_search, ekspedisi, atau all untuk semua data.',
    input_schema: {
      type: 'object',
      properties: {
        intent: {
          type: 'string',
          description: 'Jenis filter data yang dibutuhkan',
          enum: ['summary', 'overdue', 'sla_alert', 'no_resi', 'pending', 'shipped_today', 'specific_order', 'customer_search', 'ekspedisi', 'all'],
        },
        query: {
          type: 'string',
          description: 'Query tambahan: nomor order, nama customer, atau nama ekspedisi',
        },
      },
      required: ['intent'],
    },
  },
  {
    name: 'track_resi_jt',
    description: 'Cek status tracking resi J&T Cargo secara realtime via API. Gunakan untuk tahu posisi paket, apakah sudah diterima, atau estimasi tiba.',
    input_schema: {
      type: 'object',
      properties: {
        waybill_no: {
          type: 'string',
          description: 'Nomor resi / waybill J&T Cargo',
        },
      },
      required: ['waybill_no'],
    },
  },
  {
    name: 'track_resi_sentral',
    description: 'Cek status tracking resi Sentral Cargo secara realtime. Gunakan untuk tahu posisi paket, apakah sudah diterima, atau riwayat perjalanan.',
    input_schema: {
      type: 'object',
      properties: {
        waybill_no: {
          type: 'string',
          description: 'Nomor resi Sentral Cargo',
        },
      },
      required: ['waybill_no'],
    },
  },
  {
    name: 'update_order',
    description: 'Update data order di Google Sheet. Bisa update resi, tgl kirim, tgl tiba, remark, atau status.',
    input_schema: {
      type: 'object',
      properties: {
        no_order: { type: 'string', description: 'Nomor order yang akan diupdate' },
        field:    { type: 'string', description: 'Field yang diupdate', enum: ['resi', 'tgl_kirim', 'tgl_tiba', 'remark', 'status'] },
        value:    { type: 'string', description: 'Nilai baru' },
      },
      required: ['no_order', 'field', 'value'],
    },
  },
  {
    name: 'get_sla_status',
    description: 'Ambil ringkasan status SLA semua order: berapa yang overdue, urgent (H-0), warning (H-1), dan attention (H-2).',
    input_schema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  {
    name: 'get_daily_summary',
    description: 'Ambil ringkasan operasional hari ini: total order masuk, dikirim, waiting, belum resi, overdue, dan top kota tujuan.',
    input_schema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
];

// ─── TOOL EXECUTOR ────────────────────────────────────────
async function executeTool(toolName, toolInput, senderId, senderName) {
  console.log(`[Tool] ${toolName}:`, JSON.stringify(toolInput));
  try {
    switch (toolName) {

      case 'get_order_data': {
        const rawData = await getSheetData();
        if (!rawData) return { error: 'Gagal ambil data sheet' };
        const result = filterData(rawData, toolInput.intent, toolInput.query || '');
        if (!result) return { rows: [], total: 0 };
        return { rows: result.rows.slice(0, 50), total: result.rows.length };
      }

      case 'track_resi_jt': {
        const result = await cekResiJTCargo(toolInput.waybill_no);
        if (!result.ok) return { error: result.msg };
        return {
          waybill_no     : result.waybillNo,
          status         : result.status,
          is_delivered   : result.isDelivered,
          posisi         : result.posisi,
          update_terakhir: result.updateTerakhir,
          next_stop      : result.nextStop,
          narasi         : result.narasi,
          kota_asal      : result.kotaAsal,
          kota_tujuan    : result.kotaTujuan,
          collect_time   : result.collectTime,
        };
      }

      case 'track_resi_sentral': {
        const result = await cekResiSentralCargo(toolInput.waybill_no);
        if (!result.ok) return { error: result.msg };
        return {
          waybill_no     : result.waybillNo,
          status         : result.status,
          is_delivered   : result.isDelivered,
          posisi         : result.posisi,
          update_terakhir: result.updateTerakhir,
          narasi         : result.narasi,
          kota_asal      : result.kotaAsal,
          kota_tujuan    : result.kotaTujuan,
          collect_time   : result.collectTime,
        };
      }

      case 'update_order': {
        const fieldMap = {
          resi      : { col: COL.resi,          name: 'Resi' },
          tgl_kirim : { col: COL.tglPengiriman,  name: 'Tgl Pengiriman' },
          tgl_tiba  : { col: COL.tglTiba,        name: 'Tgl Tiba' },
          remark    : { col: COL.remark,         name: 'Remark' },
          status    : { col: COL.status,         name: 'Status' },
        };
        const f = fieldMap[toolInput.field];
        if (!f) return { error: 'Field tidak valid' };
        const value = ['tgl_kirim','tgl_tiba'].includes(toolInput.field)
          ? parseDate(toolInput.value)
          : toolInput.value;
        const ok = await updateOrderField(toolInput.no_order, f.col, f.name, value, senderName);
        return ok
          ? { success: true, no_order: toolInput.no_order, field: toolInput.field, value }
          : { error: `Order ${toolInput.no_order} tidak ditemukan` };
      }

      case 'get_sla_status': {
        const rawData = await getSheetData();
        if (!rawData) return { error: 'Gagal ambil data' };
        const alerts = computeSLAAlerts(rawData);
        return alerts || { urgent: [], warning: [], attention: [], overdue: [] };
      }

      case 'get_daily_summary': {
        const rawData = await getSheetData();
        if (!rawData) return { error: 'Gagal ambil data' };
        return computeDailySummary(rawData) || {};
      }

      default:
        return { error: `Tool tidak dikenal: ${toolName}` };
    }
  } catch (e) {
    console.error(`[Tool Error] ${toolName}:`, e.message);
    return { error: e.message };
  }
}

// ─── CLAUDE CALL (with tool use) ──────────────────────────
async function callClaude(senderId, senderName, userMessage, imageBase64 = null, imageMime = 'image/jpeg') {
  const today = getToday();

  // ── Foto: langsung ke Claude tanpa tool ──────────────────
  if (imageBase64) {
    if (!chatHistory[senderId]) chatHistory[senderId] = [];
    const messages = [...chatHistory[senderId]];
    messages.push({ role: 'user', content: [
      { type: 'image', source: { type: 'base64', media_type: imageMime, data: imageBase64 } },
      { type: 'text',  text: `Tanggal: ${today}\nUser: ${senderName}\n\n${userMessage || 'Tolong baca nomor resi dari foto ini.'}` },
    ]});
    const resp = await anthropic.messages.create({
      model     : 'claude-haiku-4-5-20251001',
      max_tokens: 1000,
      system    : [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
      messages,
    });
    const reply = resp.content[0].text;
    chatHistory[senderId].push({ role: 'user',      content: userMessage || '[foto]' });
    chatHistory[senderId].push({ role: 'assistant', content: reply });
    if (chatHistory[senderId].length > 12) chatHistory[senderId] = chatHistory[senderId].slice(-12);
    await parseActions(reply, senderId, senderName);
    return reply.replace(/ACTION:[^\n]+/g, '').trim();
  }

  // ── Shortcut tanpa Claude ─────────────────────────────────
  const intent = detectIntent(userMessage || '', senderId);
  if (intent === 'confirm_excel') return await insertExcelToSheet(senderName, senderId);
  if (intent === 'prompt_excel')  return '📎 Silakan kirim file Excel-nya langsung ke sini ya!';
  if (intent === 'refresh')       { clearCache(); return '🔄 Data berhasil di-refresh!'; }
  if (intent === 'log_today')     return await getLogToday();
  if (intent === 'save_instruction') {
    const cnt = (userMessage || '').replace(/^(catat|ingat|note)\s*:/i, '').trim();
    await saveMemory('instruksi', cnt);
    return '✅ Dicatat! Aku akan ingat instruksi ini.';
  }
  if (intent === 'list_pengiriman') {
    const result = await buildListPengiriman(userMessage, false);
    return result;
  }
  if (intent === 'out_of_scope') return 'Hmmm gatau sih, diluar konteks itu keknya 😅';

  // ── Siapkan context (memory + reminder) ──────────────────
  const [memories, reminders] = await Promise.all([getMemory(), getReminders()]);
  const instruksi        = memories.filter(m => m[1] === 'instruksi');
  const memoryOther      = memories.filter(m => m[1] !== 'instruksi');
  const todayReminders   = reminders.filter(r => r[3] === today && r[5] === 'Pending');
  const upcomingReminders= reminders.filter(r => r[3] >= today && r[5] === 'Pending');

  let context = `Tanggal hari ini (WIB): ${today}\nUser: ${senderName}\n\n`;
  if (instruksi.length)        { context += `=== ⚡ INSTRUKSI ===\n`; instruksi.forEach(m => { context += `• ${m[2]}\n`; }); context += '\n'; }
  if (memoryOther.length)      { context += `=== MEMORY ===\n`; memoryOther.forEach(m => { context += `[${m[1]}] ${m[2]}\n`; }); context += '\n'; }
  if (todayReminders.length)   { context += `=== REMINDER HARI INI ===\n`; todayReminders.forEach(r => { context += `Oleh: ${r[1]} | Order: ${r[2]||'-'} | ${r[4]}\n`; }); context += '\n'; }
  if (upcomingReminders.length){ context += `=== UPCOMING REMINDERS ===\n`; upcomingReminders.forEach(r => { context += `[${r[3]}] Oleh: ${r[1]} | Order: ${r[2]||'-'} | ${r[4]}\n`; }); context += '\n'; }

  // ── Tool use loop ─────────────────────────────────────────
  if (!chatHistory[senderId]) chatHistory[senderId] = [];
  const messages = [...chatHistory[senderId]];
  messages.push({ role: 'user', content: `${context}\n\n${userMessage}` });

  let finalReply = '';
  let loopCount  = 0;
  const MAX_LOOPS = 5;

  while (loopCount < MAX_LOOPS) {
    loopCount++;

    const response = await anthropic.messages.create({
      model     : 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      system    : [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
      tools     : TOOLS,
      messages,
    });

    const usage = response.usage;
    console.log(`[Tokens L${loopCount}] in:${usage.input_tokens} out:${usage.output_tokens} cache_write:${usage.cache_creation_input_tokens||0} cache_read:${usage.cache_read_input_tokens||0}`);

    // ── End turn: Claude selesai jawab ──────────────────────
    if (response.stop_reason === 'end_turn') {
      const textBlock = response.content.find(b => b.type === 'text');
      finalReply = textBlock ? textBlock.text : '';
      messages.push({ role: 'assistant', content: response.content });
      break;
    }

    // ── Tool use: Claude minta jalankan tool ────────────────
    if (response.stop_reason === 'tool_use') {
      messages.push({ role: 'assistant', content: response.content });

      // Eksekusi semua tool yang diminta (bisa paralel kalau >1)
      const toolUseBlocks = response.content.filter(b => b.type === 'tool_use');
      const toolResults   = await Promise.all(
        toolUseBlocks.map(async (block) => {
          const result = await executeTool(block.name, block.input, senderId, senderName);
          return {
            type        : 'tool_result',
            tool_use_id : block.id,
            content     : JSON.stringify(result),
          };
        })
      );

      messages.push({ role: 'user', content: toolResults });
      continue;
    }

    // Fallback
    break;
  }

  // ── Simpan history ────────────────────────────────────────
  chatHistory[senderId].push({ role: 'user',      content: userMessage });
  chatHistory[senderId].push({ role: 'assistant', content: finalReply  });
  if (chatHistory[senderId].length > 12) chatHistory[senderId] = chatHistory[senderId].slice(-12);
  if (Object.keys(chatHistory).length > HISTORY_MAX_SENDERS) {
    delete chatHistory[Object.keys(chatHistory)[0]];
  }

  await parseActions(finalReply, senderId, senderName);
  return finalReply.replace(/ACTION:[^\n]+/g, '').trim();
}

async function parseActions(text, senderId, senderName) {
  for (const line of text.split('\n')) {
    const t = line.trim();
    try {
      if (t.startsWith('ACTION:UPDATE_RESI:')) { const rest=t.replace('ACTION:UPDATE_RESI:',''),idx=rest.indexOf(':'); if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.resi,'Resi',rest.substring(idx+1),senderName); }
      if (t.startsWith('ACTION:UPDATE_TGL_KIRIM:')) { const rest=t.replace('ACTION:UPDATE_TGL_KIRIM:',''),idx=rest.indexOf(':'); if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.tglPengiriman,'Tgl Pengiriman',parseDate(rest.substring(idx+1)),senderName); }
      if (t.startsWith('ACTION:UPDATE_TGL_TIBA:')) { const rest=t.replace('ACTION:UPDATE_TGL_TIBA:',''),idx=rest.indexOf(':'); if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.tglTiba,'Tgl Tiba',parseDate(rest.substring(idx+1)),senderName); }
      if (t.startsWith('ACTION:UPDATE_REMARK:')) { const rest=t.replace('ACTION:UPDATE_REMARK:',''),idx=rest.indexOf(':'); if(idx>-1)await updateOrderField(rest.substring(0,idx),COL.remark,'Remark',rest.substring(idx+1),senderName); }
      if (t.startsWith('ACTION:SAVE_MEMORY:')) { const rest=t.replace('ACTION:SAVE_MEMORY:',''),idx=rest.indexOf(':'); if(idx>-1)await saveMemory(rest.substring(0,idx),rest.substring(idx+1)); }
      if (t.startsWith('ACTION:SAVE_REMINDER:')) {
        const rest=t.replace('ACTION:SAVE_REMINDER:',''), parts=rest.split(':');
        if (parts.length >= 3) {
          const noOrders = parts[0].split(/\s+dan\s+|,\s*|&\s*/).map(s=>s.trim()).filter(Boolean);
          await saveReminder(senderName, noOrders, parseDate(parts[1]), parts.slice(2).join(':'));
        }
      }
    } catch (e) { console.error(`parseActions error: ${e.message}`); }
  }
}

// ─── LOG TODAY ────────────────────────────────────────────
async function getLogToday() {
  await ensureTab(UPDATE_LOG_TAB, ['Timestamp','Oleh','No Order','Kolom','Nilai Lama','Nilai Baru']);
  const s = await getSheets(); if (!s) return 'Gagal ambil log.';
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId:SHEET_ID, range:`${UPDATE_LOG_TAB}!A:F` });
    const rows = (res.data.values || []).slice(1);
    const today = getToday();
    const [yr, mo, dy] = today.split('-');
    const todayLogs = rows.filter(r => { const ts = r[0]||''; return ts.includes(`${parseInt(dy)}/${parseInt(mo)}/${yr}`) || ts.includes(`${dy}/${mo}/${yr}`); });
    if (!todayLogs.length) return '📋 Belum ada update hari ini.';
    let msg = `📋 LOG UPDATE HARI INI\n\n`;
    todayLogs.forEach(r => { msg += `🕐 ${r[0]} | ${r[1]}\n📦 ${r[2]} | ${r[3]}: ${r[4]||'—'} → ${r[5]}\n\n`; });
    return msg;
  } catch (e) { return 'Gagal ambil log: ' + e.message; }
}

// ─── MESSAGING ────────────────────────────────────────────
async function sendTelegram(chatId, message) {
  try { await axios.post(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, { chat_id:chatId, text:message }); }
  catch (e) { console.error('sendTelegram error:', e.response?.data || e.message); }
}

async function sendWA(target, message, mentions = []) {
  try {
    const payload = { target, message };
    if (mentions.length > 0) payload.mentions = mentions.join(',');
    await axios.post('https://api.fonnte.com/send', payload, { headers: { Authorization: FONNTE_TOKEN } });
  } catch (e) { console.error('sendWA error:', e.message); }
}

async function getTelegramFile(fileId) {
  const fileRes = await axios.get(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
  const fileUrl = `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${fileRes.data.result.file_path}`;
  const res = await axios.get(fileUrl, { responseType:'arraybuffer' });
  return Buffer.from(res.data);
}

// ─── HELPER: CEK MENTION BOT ──────────────────────────────
function isBotMentioned(message) {
  const msg = message || '';
  const msgLower = msg.toLowerCase();
  if (msgLower.includes(`@${BOT_MENTION}`)) return true;
  if (BOT_WA_NUMBER) {
    const num = BOT_WA_NUMBER.replace(/\D/g, '');
    const variants = [
      num,
      num.replace(/^62/, '0'),
      num.replace(/^62/, ''),
      num.replace(/^0/, '62'),
    ];
    for (const v of variants) {
      if (v && msg.includes(`@${v}`)) return true;
    }
  }
  return false;
}

// ─── GROUP WA HANDLER ─────────────────────────────────────
async function handleGroupMessage(groupId, rawMessage, senderName, senderPhone) {
  let cleanMsg = rawMessage.replace(new RegExp(`@${BOT_MENTION}`, 'gi'), '');
  if (BOT_WA_NUMBER) {
    const num = BOT_WA_NUMBER.replace(/\D/g, '');
    const variants = [num, num.replace(/^62/, '0'), num.replace(/^62/, '')];
    for (const v of variants) {
      if (v) cleanMsg = cleanMsg.replace(new RegExp(`@${v}`, 'g'), '');
    }
  }
  cleanMsg = cleanMsg.trim();
  console.log(`Group [${senderName}]: ${cleanMsg.substring(0, 80)}`);

  const mentionPrefix = senderPhone ? `@${senderPhone} ` : '';
  const mentionArr = senderPhone ? [senderPhone] : [];

  if (/list pengiriman|daftar pengiriman|pengiriman (hari ini|besok|lusa|kemarin|\d)/.test(cleanMsg.toLowerCase())) {
    const result = await buildListPengiriman(cleanMsg, true);
    await sendWA(groupId, mentionPrefix + result.msg1, mentionArr);
    if (result.msg2) await sendWA(groupId, result.msg2);
    return;
  }

  const updateMatch = cleanMsg.match(/update\s+tgl\s+kirim\s+(\d+)\s+(.+)/i);
  if (updateMatch) {
    const noOrder = updateMatch[1], tanggal = parseDate(updateMatch[2].trim());
    const ok = await updateOrderField(noOrder, COL.tglPengiriman, 'Tgl Pengiriman', tanggal, senderName);
    const reply = ok
      ? `✅ Tgl Kirim order *${noOrder}* → *${tanggal}*\nOleh: ${senderName}`
      : `⚠️ Order *${noOrder}* tidak ditemukan di sheet.`;
    await sendWA(groupId, mentionPrefix + reply, mentionArr);
    return;
  }

  const cekMatch = cleanMsg.match(/cek\s+order\s+(\d+)/i);
  if (cekMatch) {
    const noOrder = cekMatch[1];
    const row = await getOrderByNumber(noOrder);
    if (row) {
      let reply = `📦 ORDER ${row[COL.noOrder]} - ${row[COL.customer]}\n`;
      reply += `📍 ${row[COL.kota]} | 🚚 ${row[COL.ekspedisi] || '—'}\n`;
      reply += `🚢 Shipping: ${row[COL.shippingNum] || '—'}\n`;
      reply += `🔖 Resi: ${row[COL.resi] || 'Kosong ⚠️'}\n`;
      reply += `📅 Tgl Kirim: ${row[COL.tglPengiriman] || '—'}\n`;
      reply += `📅 Tgl Tiba: ${row[COL.tglTiba] || '—'}\n`;
      reply += `📊 Status: ${row[COL.status] || '—'} | SLA: ${row[COL.sla] || '—'}`;
      await sendWA(groupId, mentionPrefix + reply, mentionArr);
    } else {
      await sendWA(groupId, mentionPrefix + `⚠️ Order *${noOrder}* tidak ditemukan.`, mentionArr);
    }
    return;
  }

  const reminderMatch = cleanMsg.match(/reminder\s+(\d+)\s+(\S+)\s+(.+)/i);
  if (reminderMatch) {
    const noOrder = reminderMatch[1], tanggal = parseDate(reminderMatch[2]), note = reminderMatch[3];
    await saveReminder(senderName, [noOrder], tanggal, note);
    await sendWA(groupId, mentionPrefix + `📌 Reminder disimpan!\n📦 Order: ${noOrder}\n📅 Tanggal: ${tanggal}\n📌 Note: ${note}`, mentionArr);
    return;
  }

  await sendWA(groupId,
    mentionPrefix +
    `Cara pakai @${BOT_MENTION} di grup:\n\n` +
    `• *Cek order:* @${BOT_MENTION} cek order [no]\n` +
    `• *Update tgl kirim:* @${BOT_MENTION} update tgl kirim [no] [tgl]\n` +
    `• *Reminder:* @${BOT_MENTION} reminder [no] [tgl] [note]\n` +
    `• *List pengiriman:* @${BOT_MENTION} list pengiriman besok`,
    mentionArr
  );
}

// ─── WEBHOOKS ─────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status:'ok', model:'claude-haiku-4-5-20251001', time_wib:new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'}) }));
app.get('/webhook/telegram', (_, res) => res.sendStatus(200));
app.get('/webhook/wa', (_, res) => res.sendStatus(200));

// Telegram
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  const msg = req.body.message || req.body.edited_message; if (!msg) return;
  const chatId = msg.chat.id;
  const firstName = msg.from?.first_name || msg.chat?.first_name || 'User';
  const senderId = `tg_${chatId}`;
  let text = msg.text || msg.caption || '', imageBase64 = null;
  if (msg.document) {
    const fname = msg.document.file_name || '';
    if (fname.match(/\.(xlsx|xls)$/i)) {
      try { const buf = await getTelegramFile(msg.document.file_id); await sendTelegram(chatId, await handleExcelUpload(buf, firstName, senderId, false)); }
      catch (e) { await sendTelegram(chatId, 'Gagal baca Excel: ' + e.message); }
      return;
    }
  }
  if (msg.photo && msg.photo.length > 0) {
    try { const buf = await getTelegramFile(msg.photo[msg.photo.length-1].file_id); imageBase64 = buf.toString('base64'); if (!text) text = 'Tolong baca nomor resi dari foto ini.'; }
    catch (e) { console.error('Photo error:', e.message); }
  }
  if (!text && !imageBase64) return;
  console.log(`TG [${firstName}]: ${text.substring(0,80)}`);
  try {
    const intent = detectIntent(text, senderId);
    if (intent === 'list_pengiriman') {
      const result = await buildListPengiriman(text, false);
      await sendTelegram(chatId, result.msg1);
      if (result.msg2) await sendTelegram(chatId, result.msg2);
      return;
    }
    await sendTelegram(chatId, await callClaude(senderId, firstName, text, imageBase64));
  } catch (e) { console.error('TG error:', e.message); await sendTelegram(chatId, 'Maaf, error: ' + e.message); }
});

// WhatsApp
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const { sender, message, name, file, mimetype, member } = req.body;
  if (!sender) return;
  const isGroup = sender.includes('@g.us');
  const senderName = name || member || sender;
  const senderPhone = member || null;
  const senderId = `wa_${sender}`;
  console.log(`WA [${senderName}${isGroup?'/GROUP':''}]: "${(message||'').substring(0,60)}" file=${!!file}`);

  try {
    // GROUP MESSAGE
    if (isGroup) {
      if (!message) return;
      if (!isBotMentioned(message)) return;
      await handleGroupMessage(sender, message, senderName, senderPhone);
      return;
    }

    // PRIVATE - Excel
    if (file && mimetype && /spreadsheet|excel|xlsx|xls/i.test(mimetype)) {
      const buf = await getFonnteFile(file);
      if (!buf) { await sendWA(sender, 'Gagal download file.'); return; }
      await sendWA(sender, await handleExcelUpload(buf, senderName, senderId, true));
      return;
    }

    // PRIVATE - Foto
    if (file && mimetype && /image/i.test(mimetype)) {
      const buf = await getFonnteFile(file);
      if (!buf) { await sendWA(sender, 'Gagal download foto.'); return; }
      const caption = message || 'Tolong baca nomor resi dari foto ini.';
      await sendWA(sender, await callClaude(senderId, senderName, caption, buf.toString('base64'), mimetype));
      return;
    }

    // PRIVATE - List pengiriman (2 pesan)
    if (!message) return;
    const intent = detectIntent(message, senderId);
    if (intent === 'list_pengiriman') {
      const result = await buildListPengiriman(message, false);
      await sendWA(sender, result.msg1);
      if (result.msg2) await sendWA(sender, result.msg2);
      return;
    }

    // PRIVATE - Dry run J&T
    if (intent === 'dryrun_jt') {
      await sendWA(sender, '⏳ Dry run J&T dimulai...\nHasilnya dikirim bertahap, tunggu sebentar ya.');
      runDryRunWA(sender).catch(e => {
        console.error('runDryRunWA error:', e.message);
        sendWA(sender, '❌ Error saat dry run: ' + e.message);
      });
      return;
    }

    // PRIVATE - Dry run Sentral Cargo
    if (intent === 'dryrun_sentral') {
      await sendWA(sender, '⏳ Dry run Sentral Cargo dimulai...\nHasilnya dikirim bertahap, tunggu sebentar ya.');
      runDryRunWASentral(sender).catch(e => {
        console.error('runDryRunWASentral error:', e.message);
        sendWA(sender, '❌ Error saat dry run Sentral Cargo: ' + e.message);
      });
      return;
    }

    // PRIVATE - Teks biasa
    await sendWA(sender, await callClaude(senderId, senderName, message));

  } catch (e) { console.error('WA error:', e.message); await sendWA(sender, 'Maaf, error: ' + e.message); }
});

// ─── SCHEDULERS ───────────────────────────────────────────

// 08:00 WIB — H-1 Pending + SLA Alert
cron.schedule('0 8 * * *', async () => {
  console.log('=== SCHEDULER 08:00 WIB ===');
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    const today = getToday();
    const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toLocaleDateString('sv-SE', { timeZone:'Asia/Jakarta' });
    if (data) {
      const h1 = data.slice(1).filter(r => (r[COL.requestDate]||'').trim() === tomorrowStr && (r[COL.status]||'').toLowerCase() !== 'received');
      if (h1.length > 0) {
        let msg = `🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
        h1.forEach(r => { msg += `📦 ${r[COL.noOrder]} — ${r[COL.customer]} (${r[COL.ekspedisi]||'—'})\n`; msg += r[COL.resi] ? `  ✅ Resi: ${r[COL.resi]}\n` : `  ❌ Resi belum diinput!\n`; });
        msg += `\nTotal: ${h1.length} order`;
        await sendToTargets(msg);
      }
      const sla = computeSLAAlerts(data);
      if (sla && (sla.urgent.length > 0 || sla.overdue.length > 0)) {
        let msg = `⚠️ ALERT SLA - ${today}\n\n`;
        if (sla.overdue.length) { msg += `🚨 OVERDUE (${sla.overdue.length}):\n`; sla.overdue.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`; }); msg += '\n'; }
        if (sla.urgent.length) { msg += `🔴 URGENT H-0 (${sla.urgent.length}):\n`; sla.urgent.forEach(o => { msg += `• ${o.no_order} - ${o.customer}\n`; }); msg += '\n'; }
        if (sla.warning.length) { msg += `🟡 WARNING H-1 (${sla.warning.length}):\n`; sla.warning.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`; }); }
        await sendToTargets(msg);
      }
    }
  } catch (e) { console.error('Scheduler 08:00 error:', e.message); }
}, { timezone:'Asia/Jakarta' });

// 17:00 WIB — Reminder H-1
cron.schedule('0 17 * * *', async () => {
  console.log('=== SCHEDULER 17:00 WIB — Reminder H-1 ===');
  try {
    clearCache('reminders');
    const reminders = await getReminders();
    const tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toLocaleDateString('sv-SE', { timeZone:'Asia/Jakarta' });
    const h1 = reminders.filter(r => r[3] === tomorrowStr && r[5] === 'Pending');
    for (const r of h1) { await sendToTargets(await buildReminderMsg(r, '⏰ REMINDER BESOK')); }
    console.log(`H-1 reminders sent: ${h1.length}`);
  } catch (e) { console.error('Scheduler 17:00 error:', e.message); }
}, { timezone:'Asia/Jakarta' });

// 07:00 WIB — Reminder H-0
cron.schedule('0 7 * * *', async () => {
  console.log('=== SCHEDULER 07:00 WIB — Reminder H-0 ===');
  try {
    clearCache('reminders');
    const reminders = await getReminders();
    const today = getToday();
    const todayR = reminders.filter(r => r[3] === today && r[5] === 'Pending');
    for (let i = 0; i < todayR.length; i++) {
      await sendToTargets(await buildReminderMsg(todayR[i], '🔔 REMINDER HARI INI'));
      await markReminderDone(reminders.indexOf(todayR[i]));
    }
    console.log(`H-0 reminders sent: ${todayR.length}`);
  } catch (e) { console.error('Scheduler 07:00 error:', e.message); }
}, { timezone:'Asia/Jakarta' });

// ═══════════════════════════════════════════════════════════
//  J&T CARGO TRACKING + DAILY REPORT
// ═══════════════════════════════════════════════════════════

// ─── J&T CARGO API ────────────────────────────────────────
const JT_BASE          = 'https://office.jtcargo.co.id/official/waybill/';
const JT_VALIDATE_CODE = '6251';
const JT_HEADERS       = {
  'Content-Type': 'application/json',
  'Origin'      : 'https://www.jtcargo.id',
  'Referer'     : 'https://www.jtcargo.id/',
  'language'    : 'ID',
};

async function cekResiJTCargo(waybillNo) {
  try {
    const r1 = await axios.post(JT_BASE + 'trackingIsNotEmpty',
      { waybillNo },
      { headers: JT_HEADERS, timeout: 10000 }
    );
    if (!r1.data?.succ || !r1.data?.data) {
      return { ok: false, msg: 'Resi tidak ditemukan' };
    }

    const r2 = await axios.post(JT_BASE + 'trackingValidate',
      { waybillNo, validateCode: JT_VALIDATE_CODE, searchWaybillOrCustomerOrderId: '1' },
      { headers: JT_HEADERS, timeout: 10000 }
    );
    if (!r2.data?.succ) {
      return { ok: false, msg: 'Validasi gagal' };
    }

    const r3 = await axios.post(JT_BASE + 'trackingCustomerByWaybillNo',
      { waybillNo, langType: 'ID', searchWaybillOrCustomerOrderId: '1' },
      { headers: JT_HEADERS, timeout: 10000 }
    );
    if (!r3.data?.succ || !r3.data?.data?.length) {
      return { ok: false, msg: 'Data tracking kosong' };
    }

    const trk    = r3.data.data[0];
    const latest = trk.details[0];

    return {
      ok            : true,
      waybillNo,
      statusCode    : latest.code,
      status        : latest.status,
      isDelivered   : latest.code === 100,
      posisi        : `${latest.scanNetworkName} — ${latest.scanNetworkCity}`,
      updateTerakhir: latest.scanTime,
      nextStop      : latest.nextStopName || '-',
      narasi        : latest.customerTracking || '-',
      kotaAsal      : trk.senderCityName  || '-',
      kotaTujuan    : trk.receiverCityName || '-',
      collectTime   : trk.collectTime     || '-',
      details       : trk.details,
    };

  } catch (e) {
    return { ok: false, msg: e.message || 'Error tidak diketahui' };
  }
}

// ─── ANALISA SLA VIA CLAUDE ───────────────────────────────
async function analisaSLAJTCargo(orderInfo, trackingData) {
  try {
    const prompt =
      `Kamu adalah analis logistik pengiriman Indonesia. Analisa pengiriman ini.\n\n` +
      `DATA:\n` +
      `- Ekspedisi  : J&T Cargo\n` +
      `- Resi       : ${trackingData.waybillNo}\n` +
      `- Asal       : ${trackingData.kotaAsal}\n` +
      `- Tujuan     : ${trackingData.kotaTujuan}\n` +
      `- Pickup     : ${trackingData.collectTime}\n` +
      `- Hari jalan : ${orderInfo.hariJalan} hari\n` +
      `- SLA        : ${orderInfo.sla} hari\n` +
      `- Sisa SLA   : ${orderInfo.sisaSLA} hari\n` +
      `- Status     : ${trackingData.status}\n` +
      `- Posisi     : ${trackingData.posisi}\n` +
      `- Next stop  : ${trackingData.nextStop}\n\n` +
      `RIWAYAT 5 SCAN TERAKHIR:\n` +
      trackingData.details.slice(0, 5)
        .map(d => `${d.scanTime} — ${d.scanTypeName} di ${d.scanNetworkName} (${d.scanNetworkCity})`)
        .join('\n') +
      `\n\nBerikan analisa SINGKAT (max 2 kalimat) dalam Bahasa Indonesia:\n` +
      `Apakah berisiko telat? Apa rekomendasinya? Langsung ke intinya.`;

    const response = await anthropic.messages.create({
      model     : 'claude-haiku-4-5-20251001',
      max_tokens: 150,
      messages  : [{ role: 'user', content: prompt }],
    });

    return response.content[0]?.text || '-';
  } catch (e) {
    console.error('analisaSLA error:', e.message);
    return 'Analisa tidak tersedia';
  }
}

// ─── DAILY TRACKING REPORT ────────────────────────────────
async function dailyJTTrackingReport() {
  console.log('=== SCHEDULER 07:05 WIB — J&T Tracking Report ===');
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) {
      console.log('Tidak ada data sheet');
      return;
    }

    const today    = getToday();
    const diterima  = [];
    const berisiko  = [];
    const overdue   = [];
    const errors    = [];
    let   totalPantau = 0;

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const eks    = (row[COL.ekspedisi]     || '').trim();
      const resi   = (row[COL.resi]          || '').trim();
      const status = (row[COL.status]        || '').trim().toLowerCase();
      const slaStr = (row[COL.sla]           || '').trim();
      const tglKirim = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder  = (row[COL.noOrder]     || '').trim();
      const customer = (row[COL.customer]    || '-').trim();
      const shippingNum = (row[COL.shippingNum] || '-').trim();

      // PATCH: gunakan regex agar match J&T Cargo maupun JNT Cargo
      if (!eks.match(/j[n&]t\s*cargo/i)) continue;
      if (['delivered', 'received', 'return'].includes(status)) continue;
      if (!resi) continue;

      const hariJalan = tglKirim
        ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000)
        : 0;
      const sla = slaStr
        ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)
        : null;
      const sisaSLA = sla !== null ? sla : null;

      totalPantau++;
      await new Promise(r => setTimeout(r, 600));

      console.log(`Cek resi [${totalPantau}]: ${resi} (${customer})`);
      const result = await cekResiJTCargo(resi);

      if (!result.ok) {
        errors.push(`⚠️ ${shippingNum}\n   👤 ${customer}\n   ❌ ${result.msg}`);
        continue;
      }

      if (result.isDelivered) {
        try {
          const tglTibaVal = result.updateTerakhir.split(' ')[0];
          await updateCell(SHEET_TAB, i + 1, COL.tglTiba, tglTibaVal);
          await updateCell(SHEET_TAB, i + 1, COL.status,  'Delivered');
          await logUpdate('BOT_JT', noOrder, 'Status+TglTiba', status, `Delivered | ${tglTibaVal}`);
          console.log(`  ✅ Auto-update Delivered: ${noOrder}`);
        } catch (e) {
          console.error(`  Gagal update baris ${i+1}:`, e.message);
        }

        diterima.push(
          `✅ \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   📅 Diterima: ${result.updateTerakhir}`
        );
        continue;
      }

      if (sisaSLA !== null && sisaSLA < 0) {
        const overDays = Math.abs(sisaSLA);
        overdue.push(
          `🔴 \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   ⏱️ Telat *${overDays} hari* | Jalan: ${hariJalan} hari\n` +
          `   💬 ${result.narasi}`
        );
        continue;
      }

      if (sisaSLA !== null && sisaSLA <= 1) {
        const orderInfo = { hariJalan, sla: slaStr ? Math.round((new Date(slaStr) - new Date(tglKirim)) / 86400000) : '-', sisaSLA };
        const analisa   = await analisaSLAJTCargo(orderInfo, result);

        berisiko.push(
          `⚠️ \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   ⏱️ Sisa SLA: *${sisaSLA} hari* | Jalan: ${hariJalan} hari\n` +
          `   🤖 _${analisa}_`
        );
      }
    }

    const lines = [
      `🏭 *LAPORAN HARIAN OPS PALEMBANG*`,
      `📅 ${formatDateID(today)} — 07.00 WIB`,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
    ];

    lines.push(`\n✅ *DITERIMA HARI INI — ${diterima.length} resi*`);
    if (diterima.length) { lines.push(''); diterima.forEach(x => lines.push(x)); }
    else lines.push('_Belum ada._');

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(`\n⚠️ *BERISIKO TELAT — ${berisiko.length} resi*`);
    if (berisiko.length) { lines.push(''); berisiko.forEach(x => lines.push(x)); }
    else lines.push('_Tidak ada._');

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(`\n🔴 *SUDAH LEWAT SLA — ${overdue.length} resi*`);
    if (overdue.length) { lines.push(''); overdue.forEach(x => lines.push(x)); }
    else lines.push('_Tidak ada._');

    if (errors.length) {
      lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
      lines.push(`\n🚫 *GAGAL CEK — ${errors.length} resi*\n`);
      errors.forEach(x => lines.push(x));
    }

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(
      `📊 Total dipantau : *${totalPantau}* resi\n` +
      `✅ Diterima        : *${diterima.length}*\n` +
      `⚠️ Berisiko telat  : *${berisiko.length}*\n` +
      `🔴 Lewat SLA       : *${overdue.length}*\n` +
      `🚫 Gagal cek       : *${errors.length}*\n` +
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
      `🤖 _Powered by Claude AI_`
    );

    const msg = lines.join('\n');
    await sendToTargets(msg);
    console.log(`J&T Report selesai — Diterima: ${diterima.length}, Berisiko: ${berisiko.length}, Overdue: ${overdue.length}, Error: ${errors.length}`);

  } catch (e) {
    console.error('dailyJTTrackingReport error:', e.message);
    await sendToTargets(`❌ Gagal generate laporan J&T Cargo:\n${e.message}`);
  }
}

// ─── DRY RUN J&T VIA WA ───────────────────────────────────
async function runDryRunWA(waTarget) {
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) {
      await sendWA(waTarget, '❌ Tidak ada data sheet.');
      return;
    }

    const today    = getToday();
    const diterima = [];
    const berisiko = [];
    const overdue  = [];
    const errors   = [];
    const skipped  = [];
    let   totalPantau = 0;

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      const eks         = (row[COL.ekspedisi]      || '').trim();
      const resi        = (row[COL.resi]           || '').trim();
      const status      = (row[COL.status]         || '').trim().toLowerCase();
      const slaStr      = (row[COL.sla]            || '').trim();
      const tglKirim    = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder     = (row[COL.noOrder]        || '').trim();
      const customer    = (row[COL.customer]       || '-').trim();
      const shippingNum = (row[COL.shippingNum]    || '-').trim();

      if (!eks.match(/j[n&]t\s*cargo/i)) {
        skipped.push(`${noOrder || '-'} — bukan JNT (${eks || 'kosong'})`);
        continue;
      }
      if (['delivered', 'received', 'return'].includes(status)) {
        skipped.push(`${noOrder} — ${status}`);
        continue;
      }
      if (!resi) {
        skipped.push(`${noOrder} — resi kosong`);
        continue;
      }

      const hariJalan = tglKirim
        ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000)
        : 0;
      const sisaSLA = slaStr
        ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)
        : null;

      totalPantau++;

      // Progress update setiap 5 resi
      if (totalPantau % 5 === 0) {
        await sendWA(waTarget, `⏳ Memproses resi ke-${totalPantau}...`);
      }

      await new Promise(r => setTimeout(r, 600));

      console.log(`[DRY RUN WA] Cek resi [${totalPantau}]: ${resi} (${customer})`);
      const result = await cekResiJTCargo(resi);

      if (!result.ok) {
        errors.push(`❌ ${shippingNum}\n   👤 ${customer}\n   ⚠️ ${result.msg}`);
        continue;
      }

      const base =
        `📦 ${shippingNum}\n` +
        `👤 ${customer}\n` +
        `📍 ${result.posisi}\n` +
        `🕐 ${result.updateTerakhir}`;

      if (result.isDelivered) {
        diterima.push(`✅ ${base}\n   📅 Diterima: ${result.updateTerakhir}`);
      } else if (sisaSLA !== null && sisaSLA < 0) {
        overdue.push(`🔴 ${base}\n   ⏱️ Telat *${Math.abs(sisaSLA)} hari* | Jalan: ${hariJalan} hari\n   💬 ${result.narasi}`);
      } else if (sisaSLA !== null && sisaSLA <= 1) {
        berisiko.push(`⚠️ ${base}\n   ⏱️ Sisa SLA: *${sisaSLA} hari* | Jalan: ${hariJalan} hari\n   💬 ${result.narasi}`);
      } else {
        skipped.push(`${noOrder} — aman (sisa SLA: ${sisaSLA !== null ? sisaSLA + ' hari' : 'N/A'})`);
      }
    }

    // Kirim summary
    await sendWA(waTarget,
      `🧪 *DRY RUN SELESAI*\n` +
      `📅 ${formatDateID(today)}\n` +
      `━━━━━━━━━━━━━━━\n` +
      `📊 Dipantau : *${totalPantau}*\n` +
      `✅ Diterima  : *${diterima.length}*\n` +
      `⚠️ Berisiko  : *${berisiko.length}*\n` +
      `🔴 Overdue   : *${overdue.length}*\n` +
      `❌ Error     : *${errors.length}*\n` +
      `⏭️ Skip      : *${skipped.length}*\n` +
      `━━━━━━━━━━━━━━━\n` +
      `_Sheet tidak diubah_`
    );

    if (diterima.length) await sendWA(waTarget, `✅ *DITERIMA (${diterima.length})*\n\n` + diterima.join('\n\n'));
    if (berisiko.length) await sendWA(waTarget, `⚠️ *BERISIKO TELAT (${berisiko.length})*\n\n` + berisiko.join('\n\n'));
    if (overdue.length)  await sendWA(waTarget, `🔴 *SUDAH LEWAT SLA (${overdue.length})*\n\n` + overdue.join('\n\n'));
    if (errors.length)   await sendWA(waTarget, `❌ *GAGAL CEK (${errors.length})*\n\n` + errors.join('\n\n'));

    if (skipped.length) {
      const preview = skipped.slice(0, 15).join('\n');
      const more    = skipped.length > 15 ? `\n...+${skipped.length - 15} lainnya` : '';
      await sendWA(waTarget, `⏭️ *SKIP (${skipped.length})*\n\n${preview}${more}`);
    }

  } catch (e) {
    console.error('runDryRunWA error:', e.message);
    await sendWA(waTarget, `❌ Dry run gagal: ${e.message}`);
  }
}

// ─── SCHEDULER 07:05 WIB — J&T Tracking ──────────────────
cron.schedule('5 7 * * *', async () => {
  await dailyJTTrackingReport();
}, { timezone: 'Asia/Jakarta' });

// ─── DEBUG ENDPOINTS ──────────────────────────────────────
app.get('/health', (_, res) => res.json({ status:'ok', model:'claude-haiku-4-5-20251001', time_wib:new Date().toLocaleString('id-ID',{timeZone:'Asia/Jakarta'}) }));

app.get('/debug/jt-report', async (req, res) => {
  res.json({ ok: true, msg: 'Report sedang diproses, cek WA/Telegram dalam beberapa menit...' });
  await dailyJTTrackingReport();
});

app.get('/debug/jt-cek', async (req, res) => {
  const resi = req.query.resi;
  if (!resi) return res.json({ error: 'Query ?resi= diperlukan' });
  const result = await cekResiJTCargo(resi);
  res.json(result);
});

app.get('/debug/jt-dryrun', async (req, res) => {
  console.log('=== DRY RUN J&T Tracking Report ===');
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) {
      return res.json({ ok: false, msg: 'Tidak ada data sheet' });
    }

    const today     = getToday();
    const diterima  = [];
    const berisiko  = [];
    const overdue   = [];
    const errors    = [];
    const skipped   = [];
    let   totalPantau = 0;

    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const eks      = (row[COL.ekspedisi]      || '').trim();
      const resi     = (row[COL.resi]           || '').trim();
      const status   = (row[COL.status]         || '').trim().toLowerCase();
      const slaStr   = (row[COL.sla]            || '').trim();
      const tglKirim = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder  = (row[COL.noOrder]        || '').trim();
      const customer = (row[COL.customer]       || '-').trim();
      const shippingNum = (row[COL.shippingNum] || '-').trim();

      if (!eks.match(/j[n&]t\s*cargo/i)) {
        skipped.push({ noOrder, customer, eks, reason: 'bukan JNT Cargo' });
        continue;
      }
      if (['delivered', 'received', 'return'].includes(status)) {
        skipped.push({ noOrder, customer, eks, status, reason: 'status sudah selesai' });
        continue;
      }
      if (!resi) {
        skipped.push({ noOrder, customer, eks, reason: 'resi kosong' });
        continue;
      }

      const hariJalan = tglKirim
        ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000)
        : 0;
      const sisaSLA = slaStr
        ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)
        : null;

      totalPantau++;
      await new Promise(r => setTimeout(r, 600));

      const result = await cekResiJTCargo(resi);

      if (!result.ok) {
        errors.push({ shippingNum, customer, noOrder, resi, error: result.msg });
        continue;
      }

      const base = { shippingNum, noOrder, customer, resi, hariJalan, sisaSLA, slaStr, statusTracking: result.status, posisi: result.posisi, updateTerakhir: result.updateTerakhir, nextStop: result.nextStop, narasi: result.narasi };

      if (result.isDelivered) {
        diterima.push({ ...base, tglDiterima: result.updateTerakhir });
      } else if (sisaSLA !== null && sisaSLA < 0) {
        overdue.push({ ...base, hariTelat: Math.abs(sisaSLA) });
      } else if (sisaSLA !== null && sisaSLA <= 1) {
        const orderInfo = { hariJalan, sla: slaStr ? Math.round((new Date(slaStr) - new Date(tglKirim)) / 86400000) : '-', sisaSLA };
        const analisa = await analisaSLAJTCargo(orderInfo, result);
        berisiko.push({ ...base, analisa });
      } else {
        skipped.push({ ...base, reason: 'aman, sisa SLA > 1 hari' });
      }
    }

    res.json({
      ok: true, dryRun: true, tanggal: today, totalDipantau: totalPantau,
      summary: { diterima: diterima.length, berisiko: berisiko.length, overdue: overdue.length, errors: errors.length, skipped: skipped.length },
      diterima, berisiko, overdue, errors, skipped,
    });

  } catch (e) {
    console.error('DRY RUN error:', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
//  SENTRAL CARGO TRACKING
// ═══════════════════════════════════════════════════════════

const SENTRAL_BASE_URL    = 'https://www.sentralcargo.co.id';
const SENTRAL_CEKRESI_URL = `${SENTRAL_BASE_URL}/cekresi`;
const SENTRAL_SUBMIT_URL  = `${SENTRAL_BASE_URL}/resi/data/tracking/submit`;

const SENTRAL_HEADERS = {
  'User-Agent'     : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept'         : 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.9,id;q=0.8',
};

function extractSentralToken(html) {
  const $ = cheerio.load(html);
  // Fallback: coba beberapa selector agar tidak breakjika Sentral update UI
  const token =
    $('#cekResiForm input[name="_token"]').attr('value') ||
    $('form input[name="_token"]').first().attr('value') ||
    $('input[name="_token"]').first().attr('value');
  if (!token) throw new Error('Gagal menemukan _token di halaman /cekresi.');
  return token;
}

function extractSentralTrackingData(html) {
  // Regex utama
  let match = html.match(/setTracking\(\s*'[^']*'\s*,\s*(\[[\s\S]*?\])\s*,\s*'(\d+)'/);

  // Fallback regex lebih longgar kalau format sedikit berbeda
  if (!match) {
    match = html.match(/setTracking\([^,]*,\s*(\[[\s\S]*?\])\s*,\s*['"](\d+)['"]/);
  }

  if (!match) {
    if (html.includes('tidak ditemukan') || html.includes('not found')) return { found: false, reason: 'not_found', data: [] };
    if (html.includes('Verifikasi nomor telepon'))                       return { found: false, reason: 'otp_required', data: [] };
    throw new Error('Gagal menemukan data tracking (setTracking) di halaman hasil.');
  }

  const [, rawArray, resiFromPage] = match;
  let data;
  try {
    data = JSON.parse(rawArray);
  } catch (err) {
    throw new Error(`Gagal parse JSON data tracking: ${err.message}`);
  }

  data = data.filter(item => item !== null);
  return { found: data.length > 0, reason: data.length > 0 ? 'ok' : 'empty', data, resi: resiFromPage };
}

async function cekResiSentralCargo(noResi) {
  const jar    = new CookieJar();
  const client = wrapper(axios.create({
    jar,
    withCredentials: true,
    headers        : SENTRAL_HEADERS,
    maxRedirects   : 0,
    validateStatus : status => status >= 200 && status < 400,
    timeout        : 15000,
  }));

  try {
    // Step 1: GET /cekresi → ambil _token
    const initialRes = await client.get(SENTRAL_CEKRESI_URL);
    const token      = extractSentralToken(initialRes.data);

    // Step 2: POST submit form
    const formBody = new URLSearchParams({
      _token      : token,
      nomor_resi  : noResi,
      no_token    : '',
      form_botcheck: '',
    }).toString();

    const submitRes = await client.post(SENTRAL_SUBMIT_URL, formBody, {
      headers: {
        ...SENTRAL_HEADERS,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Referer'     : SENTRAL_CEKRESI_URL,
      },
    });

    if (submitRes.status !== 302) {
      return { ok: false, msg: `Server tidak merespons redirect (status: ${submitRes.status}).` };
    }

    // Step 3: GET redirect → ambil hasil tracking
    const redirectLocation = submitRes.headers.location || SENTRAL_CEKRESI_URL;
    const resultRes = await client.get(redirectLocation, {
      headers     : { ...SENTRAL_HEADERS, Referer: SENTRAL_SUBMIT_URL },
      maxRedirects: 5,
    });

    const tracking = extractSentralTrackingData(resultRes.data);

    if (!tracking.found) {
      const messages = {
        not_found   : 'Resi tidak ditemukan di Sentral Cargo',
        otp_required: 'Verifikasi tambahan diminta oleh server Sentral Cargo',
        empty       : 'Resi ditemukan tetapi belum ada riwayat perjalanan',
      };
      return { ok: false, msg: messages[tracking.reason] || 'Data tidak ditemukan' };
    }

    const latest      = tracking.data[0];
    const isDelivered = latest.PODStat === 'Delivered';

    // Parse tanggal — handle format "2026-05-20 14:30:00" dan "2026-05-20T14:30:00"
    const rawTgl  = latest.PODDt || '';
    const tglTiba = isDelivered ? rawTgl.split(/[T\s]/)[0] : '';

    return {
      ok            : true,
      waybillNo     : noResi,
      statusCode    : latest.StatCode  || '',
      status        : latest.PODStat   || 'ON PROCESS',
      isDelivered,
      posisi        : latest.PODDesc   || '—',
      updateTerakhir: rawTgl           || '—',
      nextStop      : '-',
      narasi        : latest.PODDesc   || '—',
      tglTiba,
      kotaAsal      : '-',
      kotaTujuan    : '-',
      collectTime   : tracking.data[tracking.data.length - 1]?.PODDt || '-',
      details       : tracking.data.map(d => ({
        scanTime        : d.PODDt,
        scanTypeName    : d.PODStat,
        scanNetworkName : (d.PODDesc.match(/\[([^\]]+)\]/) || [])[1] || d.PODDesc,
        scanNetworkCity : '-',
        customerTracking: d.PODDesc,
      })),
    };
  } catch (error) {
    return { ok: false, msg: error.message || 'Error tidak diketahui' };
  }
}

// ─── ANALISA SLA SENTRAL via Claude ──────────────────────────
async function analisaSLASentralCargo(orderInfo, trackingData) {
  try {
    const prompt =
      `Kamu adalah analis logistik pengiriman Indonesia. Analisa pengiriman Sentral Cargo ini.\n\n` +
      `DATA:\n` +
      `- Ekspedisi  : Sentral Cargo\n` +
      `- Resi       : ${trackingData.waybillNo}\n` +
      `- Hari jalan : ${orderInfo.hariJalan} hari\n` +
      `- SLA        : ${orderInfo.sla} hari\n` +
      `- Sisa SLA   : ${orderInfo.sisaSLA} hari\n` +
      `- Status     : ${trackingData.status}\n` +
      `- Posisi     : ${trackingData.posisi}\n\n` +
      `RIWAYAT 5 SCAN TERAKHIR:\n` +
      (trackingData.details || []).slice(0, 5)
        .map(d => `${d.scanTime} — ${d.scanTypeName} di ${d.scanNetworkName}`)
        .join('\n') +
      `\n\nBerikan analisa SINGKAT (max 2 kalimat) dalam Bahasa Indonesia:\n` +
      `Apakah berisiko telat? Apa rekomendasinya? Langsung ke intinya.`;

    const response = await anthropic.messages.create({
      model     : 'claude-haiku-4-5-20251001',
      max_tokens: 150,
      messages  : [{ role: 'user', content: prompt }],
    });
    return response.content[0]?.text || '-';
  } catch (e) {
    console.error('analisaSLASentral error:', e.message);
    return 'Analisa tidak tersedia';
  }
}

// ─── DAILY SENTRAL TRACKING REPORT ───────────────────────────
async function dailySentralTrackingReport() {
  console.log('=== SCHEDULER 07:30 WIB — Sentral Cargo Tracking Report ===');
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) { console.log('Tidak ada data sheet'); return; }

    const today     = getToday();
    const diterima   = [];
    const berisiko   = [];
    const overdue    = [];
    const errors     = [];
    let   totalPantau = 0;

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      const eks         = (row[COL.ekspedisi]      || '').trim();
      const resi        = (row[COL.resi]           || '').trim();
      const status      = (row[COL.status]         || '').trim().toLowerCase();
      const slaStr      = (row[COL.sla]            || '').trim();
      const tglKirim    = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder     = (row[COL.noOrder]        || '').trim();
      const customer    = (row[COL.customer]       || '-').trim();
      const shippingNum = (row[COL.shippingNum]    || '-').trim();

      if (!eks.match(/sentral\s*cargo/i)) continue;
      if (['delivered', 'received', 'return'].includes(status)) continue;
      if (!resi) continue;

      const hariJalan = tglKirim
        ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000)
        : 0;
      const sisaSLA = slaStr
        ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)
        : null;

      totalPantau++;
      await new Promise(r => setTimeout(r, 600));

      console.log(`Cek resi Sentral [${totalPantau}]: ${resi} (${customer})`);
      const result = await cekResiSentralCargo(resi);

      if (!result.ok) {
        errors.push(`⚠️ ${shippingNum}\n   👤 ${customer}\n   ❌ ${result.msg}`);
        continue;
      }

      if (result.isDelivered) {
        try {
          await updateCell(SHEET_TAB, i + 1, COL.tglTiba, result.tglTiba);
          await updateCell(SHEET_TAB, i + 1, COL.status,  'Delivered');
          await logUpdate('BOT_SENTRAL', noOrder, 'Status+TglTiba', status, `Delivered | ${result.tglTiba}`);
          console.log(`  ✅ Auto-update Delivered: ${noOrder}`);
        } catch (e) { console.error(`  Gagal update baris ${i+1}:`, e.message); }

        diterima.push(
          `✅ \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   📅 Diterima: ${result.updateTerakhir}`
        );
        continue;
      }

      if (sisaSLA !== null && sisaSLA < 0) {
        overdue.push(
          `🔴 \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   ⏱️ Telat *${Math.abs(sisaSLA)} hari* | Jalan: ${hariJalan} hari\n` +
          `   💬 ${result.narasi}`
        );
        continue;
      }

      if (sisaSLA !== null && sisaSLA <= 1) {
        const orderInfo = { hariJalan, sla: slaStr ? Math.round((new Date(slaStr) - new Date(tglKirim)) / 86400000) : '-', sisaSLA };
        const analisa   = await analisaSLASentralCargo(orderInfo, result);
        berisiko.push(
          `⚠️ \`${shippingNum}\`\n` +
          `   👤 ${customer}\n` +
          `   📍 ${result.posisi}\n` +
          `   ⏱️ Sisa SLA: *${sisaSLA} hari* | Jalan: ${hariJalan} hari\n` +
          `   🤖 _${analisa}_`
        );
      }
    }

    const lines = [
      `🏭 *LAPORAN HARIAN SENTRAL CARGO*`,
      `📅 ${formatDateID(today)} — 07.30 WIB`,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
    ];
    lines.push(`\n✅ *DITERIMA HARI INI — ${diterima.length} resi*`);
    if (diterima.length) { lines.push(''); diterima.forEach(x => lines.push(x)); }
    else lines.push('_Belum ada._');

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(`\n⚠️ *BERISIKO TELAT — ${berisiko.length} resi*`);
    if (berisiko.length) { lines.push(''); berisiko.forEach(x => lines.push(x)); }
    else lines.push('_Tidak ada._');

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(`\n🔴 *SUDAH LEWAT SLA — ${overdue.length} resi*`);
    if (overdue.length) { lines.push(''); overdue.forEach(x => lines.push(x)); }
    else lines.push('_Tidak ada._');

    if (errors.length) {
      lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
      lines.push(`\n🚫 *GAGAL CEK — ${errors.length} resi*\n`);
      errors.forEach(x => lines.push(x));
    }

    lines.push(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    lines.push(
      `📊 Total dipantau : *${totalPantau}* resi\n` +
      `✅ Diterima        : *${diterima.length}*\n` +
      `⚠️ Berisiko telat  : *${berisiko.length}*\n` +
      `🔴 Lewat SLA       : *${overdue.length}*\n` +
      `🚫 Gagal cek       : *${errors.length}*\n` +
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n` +
      `🤖 _Powered by Claude AI_`
    );

    await sendToTargets(lines.join('\n'));
    console.log(`Sentral Report — Diterima: ${diterima.length}, Berisiko: ${berisiko.length}, Overdue: ${overdue.length}, Error: ${errors.length}`);
  } catch (e) {
    console.error('dailySentralTrackingReport error:', e.message);
    await sendToTargets(`❌ Gagal generate laporan Sentral Cargo:\n${e.message}`);
  }
}

// ─── DRY RUN SENTRAL via WA ───────────────────────────────────
async function runDryRunWASentral(waTarget) {
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) { await sendWA(waTarget, '❌ Tidak ada data sheet.'); return; }

    const today     = getToday();
    const diterima   = [];
    const berisiko   = [];
    const overdue    = [];
    const errors     = [];
    const skipped    = [];
    let   totalPantau = 0;

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      const eks         = (row[COL.ekspedisi]      || '').trim();
      const resi        = (row[COL.resi]           || '').trim();
      const status      = (row[COL.status]         || '').trim().toLowerCase();
      const slaStr      = (row[COL.sla]            || '').trim();
      const tglKirim    = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder     = (row[COL.noOrder]        || '').trim();
      const customer    = (row[COL.customer]       || '-').trim();
      const shippingNum = (row[COL.shippingNum]    || '-').trim();

      if (!eks.match(/sentral\s*cargo/i)) { skipped.push(`${noOrder||'-'} — bukan Sentral Cargo`); continue; }
      if (['delivered','received','return'].includes(status)) { skipped.push(`${noOrder} — ${status}`); continue; }
      if (!resi) { skipped.push(`${noOrder} — resi kosong`); continue; }

      const hariJalan = tglKirim ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000) : 0;
      const sisaSLA   = slaStr   ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)   : null;

      totalPantau++;
      if (totalPantau % 5 === 0) await sendWA(waTarget, `⏳ Memproses resi Sentral Cargo ke-${totalPantau}...`);
      await new Promise(r => setTimeout(r, 600));

      console.log(`[DRY RUN WA - Sentral] Cek resi [${totalPantau}]: ${resi} (${customer})`);
      const result = await cekResiSentralCargo(resi);

      if (!result.ok) { errors.push(`❌ ${shippingNum}\n   👤 ${customer}\n   ⚠️ ${result.msg}`); continue; }

      const base =
        `📦 ${shippingNum}\n` +
        `👤 ${customer}\n` +
        `📍 ${result.posisi}\n` +
        `🕐 ${result.updateTerakhir}`;

      if (result.isDelivered) {
        diterima.push(`✅ ${base}`);
      } else if (sisaSLA !== null && sisaSLA < 0) {
        overdue.push(`🔴 ${base}\n   ⏱️ Telat *${Math.abs(sisaSLA)} hari* | Jalan: ${hariJalan} hari`);
      } else if (sisaSLA !== null && sisaSLA <= 1) {
        berisiko.push(`⚠️ ${base}\n   ⏱️ Sisa SLA: *${sisaSLA} hari* | Jalan: ${hariJalan} hari`);
      } else {
        skipped.push(`${noOrder} — aman (sisa SLA: ${sisaSLA !== null ? sisaSLA + ' hari' : 'N/A'})`);
      }
    }

    await sendWA(waTarget,
      `🧪 *DRY RUN SENTRAL CARGO SELESAI*\n` +
      `📅 ${formatDateID(today)}\n` +
      `━━━━━━━━━━━━━━━\n` +
      `📊 Dipantau : *${totalPantau}*\n` +
      `✅ Diterima  : *${diterima.length}*\n` +
      `⚠️ Berisiko  : *${berisiko.length}*\n` +
      `🔴 Overdue   : *${overdue.length}*\n` +
      `❌ Error     : *${errors.length}*\n` +
      `⏭️ Skip      : *${skipped.length}*\n` +
      `━━━━━━━━━━━━━━━\n` +
      `_Sheet tidak diubah_`
    );
    if (diterima.length) await sendWA(waTarget, `✅ *DITERIMA (${diterima.length})*\n\n` + diterima.join('\n\n'));
    if (berisiko.length) await sendWA(waTarget, `⚠️ *BERISIKO TELAT (${berisiko.length})*\n\n` + berisiko.join('\n\n'));
    if (overdue.length)  await sendWA(waTarget, `🔴 *SUDAH LEWAT SLA (${overdue.length})*\n\n` + overdue.join('\n\n'));
    if (errors.length)   await sendWA(waTarget, `❌ *GAGAL CEK (${errors.length})*\n\n` + errors.join('\n\n'));
    if (skipped.length) {
      const preview = skipped.slice(0,15).join('\n');
      const more    = skipped.length > 15 ? `\n...+${skipped.length-15} lainnya` : '';
      await sendWA(waTarget, `⏭️ *SKIP (${skipped.length})*\n\n${preview}${more}`);
    }
  } catch (e) {
    console.error('runDryRunWASentral error:', e.message);
    await sendWA(waTarget, `❌ Dry run Sentral Cargo gagal: ${e.message}`);
  }
}

// ─── SCHEDULER 07:30 WIB — Sentral Cargo Tracking ─────────────
cron.schedule('30 7 * * *', async () => {
  await dailySentralTrackingReport();
}, { timezone: 'Asia/Jakarta' });

// ─── DEBUG ENDPOINTS SENTRAL ───────────────────────────────────
app.get('/debug/sentral-report', async (req, res) => {
  res.json({ ok: true, msg: 'Report Sentral sedang diproses, cek WA/Telegram...' });
  await dailySentralTrackingReport();
});

app.get('/debug/sentral-cek', async (req, res) => {
  const resi = req.query.resi;
  if (!resi) return res.json({ error: 'Query ?resi= diperlukan' });
  const result = await cekResiSentralCargo(resi);
  res.json(result);
});

app.get('/debug/sentral-dryrun', async (req, res) => {
  console.log('=== DRY RUN Sentral Cargo ===');
  try {
    clearCache('sheetData');
    const data = await getSheetData();
    if (!data || data.length < 2) return res.json({ ok: false, msg: 'Tidak ada data sheet' });
    const today = getToday();
    const diterima = [], berisiko = [], overdue = [], errors = [], skipped = [];
    let totalPantau = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const eks = (row[COL.ekspedisi] || '').trim();
      const resi = (row[COL.resi] || '').trim();
      const status = (row[COL.status] || '').trim().toLowerCase();
      const slaStr = (row[COL.sla] || '').trim();
      const tglKirim = parseDate((row[COL.tglPengiriman] || '').trim());
      const noOrder = (row[COL.noOrder] || '').trim();
      const customer = (row[COL.customer] || '-').trim();
      const shippingNum = (row[COL.shippingNum] || '-').trim();
      if (!eks.match(/sentral\s*cargo/i)) { skipped.push({ noOrder, reason: 'bukan Sentral Cargo' }); continue; }
      if (['delivered','received','return'].includes(status)) { skipped.push({ noOrder, status, reason: 'sudah selesai' }); continue; }
      if (!resi) { skipped.push({ noOrder, reason: 'resi kosong' }); continue; }
      const hariJalan = tglKirim ? Math.round((new Date(today) - new Date(tglKirim)) / 86400000) : 0;
      const sisaSLA   = slaStr   ? Math.round((new Date(slaStr) - new Date(today)) / 86400000)   : null;
      totalPantau++;
      await new Promise(r => setTimeout(r, 600));
      const result = await cekResiSentralCargo(resi);
      if (!result.ok) { errors.push({ shippingNum, customer, noOrder, resi, error: result.msg }); continue; }
      const base = { shippingNum, noOrder, customer, resi, hariJalan, sisaSLA, statusTracking: result.status, posisi: result.posisi, updateTerakhir: result.updateTerakhir };
      if (result.isDelivered)                              diterima.push({ ...base });
      else if (sisaSLA !== null && sisaSLA < 0)            overdue.push({ ...base, hariTelat: Math.abs(sisaSLA) });
      else if (sisaSLA !== null && sisaSLA <= 1)           berisiko.push({ ...base });
      else                                                  skipped.push({ ...base, reason: 'aman' });
    }
    res.json({ ok: true, dryRun: true, tanggal: today, totalDipantau: totalPantau,
      summary: { diterima: diterima.length, berisiko: berisiko.length, overdue: overdue.length, errors: errors.length, skipped: skipped.length },
      diterima, berisiko, overdue, errors, skipped });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});


// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS Agent (Claude Haiku) running on port ${PORT}`);
  console.log(`Group WA: ${GROUP_WA_ID||'NOT SET'} | Mention: @${BOT_MENTION}`);
  console.log(`Reminder targets: ${REMINDER_TARGETS.length} numbers`);
});
