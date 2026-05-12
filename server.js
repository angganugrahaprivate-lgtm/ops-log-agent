require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');
const XLSX = require('xlsx');

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
const PORT             = process.env.PORT || 3000;

console.log('=== ENV CHECK ===');
console.log('ANTHROPIC_API_KEY:', ANTHROPIC_KEY ? 'OK' : 'MISSING');
console.log('GOOGLE_SHEETS_ID:', SHEET_ID ? 'OK' : 'MISSING');
console.log('FONNTE_TOKEN:', FONNTE_TOKEN ? 'OK' : 'MISSING');
console.log('TELEGRAM_BOT_TOKEN:', TELEGRAM_TOKEN ? 'OK' : 'MISSING');
console.log('=================');

// ─── ANTHROPIC CLIENT ─────────────────────────────────────
const anthropic = new Anthropic({ apiKey: ANTHROPIC_KEY });

// ─── COLUMN MAPPING ───────────────────────────────────────
// Index 0-based sesuai posisi di sheet
const COL = {
  tanggal: 0,        // A
  cutoff: 1,         // B
  shippingNum: 2,    // C
  koli: 3,           // D
  namaBarang: 4,     // E
  partner: 5,        // F
  customer: 6,       // G
  alamat: 7,         // H
  kecamatan: 8,      // I
  kota: 9,           // J
  telepon: 10,       // K
  kodePos: 11,       // L
  provinsi: 12,      // M
  ekspedisi: 13,     // N
  resi: 14,          // O
  tglPengiriman: 15, // P
  tglTiba: 16,       // Q
  aging: 17,         // R
  requestDate: 18,   // S
  remark: 19,        // T
  status: 20,        // U
  sla: 21,           // V
  noOrder: 22,       // W
};

// 12 kolom umum untuk query & monitoring
const GENERAL_COLS = [0,2,6,9,13,14,15,16,18,20,21,22];
const GENERAL_NAMES = ['Tanggal','Shipping Number','Nama Customer','Kota','Ekspedisi','Resi','Tgl Pengiriman','Tgl Tiba','Request Date','Status','SLA','No Order'];

// 7 kolom untuk update (exclude Received)
const UPDATE_COLS = [22,6,13,20,14,15,16];
const UPDATE_NAMES = ['No Order','Nama Customer','Ekspedisi','Status','Resi','Tgl Pengiriman','Tgl Tiba'];

// ─── CACHE (1 JAM) ────────────────────────────────────────
const cache = {};
const CACHE_TTL = 60 * 60 * 1000; // 1 jam

function getCache(key) {
  const c = cache[key];
  if (c && Date.now() - c.time < CACHE_TTL) return c.data;
  return null;
}
function setCache(key, data) { cache[key] = { data, time: Date.now() }; }
function clearCache(key) {
  if (key) delete cache[key];
  else Object.keys(cache).forEach(k => delete cache[k]);
}

// ─── GOOGLE SHEETS ────────────────────────────────────────
let sheetsClient = null;
async function getSheets() {
  if (sheetsClient) return sheetsClient;
  try {
    const { google } = require('googleapis');
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheetsClient = google.sheets({ version: 'v4', auth });
    console.log('Google Sheets OK');
    return sheetsClient;
  } catch (e) { console.error('Sheets init error:', e.message); return null; }
}

function toCol(n) {
  let c = '';
  while (n >= 0) { c = String.fromCharCode(65 + (n % 26)) + c; n = Math.floor(n / 26) - 1; }
  return c;
}

function getToday() {
  return new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });
}

function getCutOff() {
  const wibHour = parseInt(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta', hour: 'numeric', hour12: false }));
  return wibHour < 14 ? '1' : '2';
}

async function getSheetData() {
  const cached = getCache('sheetData');
  if (cached) { console.log('Cache hit: sheetData'); return cached; }
  const s = await getSheets();
  if (!s) return null;
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!A:AK` });
    const data = res.data.values || [];
    setCache('sheetData', data);
    console.log(`Sheet loaded: ${data.length} rows`);
    return data;
  } catch (e) { console.error('getSheetData error:', e.message); return null; }
}

function selectColumns(rows, colIndices, colNames) {
  return rows.map(row => {
    const obj = {};
    colIndices.forEach((idx, i) => { obj[colNames[i]] = (row[idx] || ''); });
    return obj;
  });
}

async function updateCell(sheetTab, rowNum, colIdx, value) {
  const s = await getSheets();
  if (!s) return;
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTab}!${toCol(colIdx)}${rowNum}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[value]] },
  });
  clearCache('sheetData');
}

async function ensureTab(tabName, headers) {
  const s = await getSheets();
  if (!s) return;
  try {
    await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${tabName}!A1` });
  } catch (e) {
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
  await ensureTab(UPDATE_LOG_TAB, ['Timestamp', 'Oleh', 'No Order', 'Kolom', 'Nilai Lama', 'Nilai Baru']);
  const s = await getSheets();
  if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: `${UPDATE_LOG_TAB}!A:F`,
      valueInputOption: 'RAW', requestBody: { values: [[ts, updatedBy, noOrder, kolom, nilaiLama, nilaiBaru]] }
    });
  } catch (e) { console.error('logUpdate error:', e.message); }
}

// ─── UPDATE FUNCTIONS ─────────────────────────────────────
async function updateOrderField(noOrder, colIdx, colName, newValue, updatedBy) {
  const data = await getSheetData();
  if (!data) return false;
  const h = data[0];
  const noOrderIdx = COL.noOrder;
  for (let i = 1; i < data.length; i++) {
    if ((data[i][noOrderIdx] || '').trim() === noOrder.trim()) {
      const oldValue = data[i][colIdx] || '';
      await updateCell(SHEET_TAB, i + 1, colIdx, newValue);
      await logUpdate(updatedBy, noOrder, colName, oldValue, newValue);
      console.log(`Updated ${colName} for ${noOrder}: ${oldValue} → ${newValue}`);
      return true;
    }
  }
  return false;
}

// ─── MEMORY ───────────────────────────────────────────────
async function getMemory() {
  const cached = getCache('memory');
  if (cached) return cached;
  await ensureTab(MEMORY_TAB, ['Timestamp', 'Category', 'Content']);
  const s = await getSheets();
  if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${MEMORY_TAB}!A:C` });
    const data = (res.data.values || []).slice(1).slice(-30);
    setCache('memory', data);
    return data;
  } catch (e) { return []; }
}

async function saveMemory(category, content) {
  await ensureTab(MEMORY_TAB, ['Timestamp', 'Category', 'Content']);
  const s = await getSheets();
  if (!s) return;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${MEMORY_TAB}!A:C`, valueInputOption: 'RAW', requestBody: { values: [[ts, category, content]] } });
    clearCache('memory');
  } catch (e) { console.error('saveMemory error:', e.message); }
}

// ─── REMINDER ─────────────────────────────────────────────
async function getReminders() {
  const cached = getCache('reminders');
  if (cached) return cached;
  await ensureTab(REMINDER_TAB, ['Timestamp', 'Dibuat Oleh', 'No Order', 'Tanggal Reminder', 'Note', 'Status']);
  const s = await getSheets();
  if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:F` });
    const data = (res.data.values || []).slice(1);
    setCache('reminders', data);
    return data;
  } catch (e) { return []; }
}

async function saveReminder(createdBy, noOrder, tanggal, note) {
  await ensureTab(REMINDER_TAB, ['Timestamp', 'Dibuat Oleh', 'No Order', 'Tanggal Reminder', 'Note', 'Status']);
  const s = await getSheets();
  if (!s) return false;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:F`, valueInputOption: 'RAW', requestBody: { values: [[ts, createdBy, noOrder || '', tanggal, note, 'Pending']] } });
    clearCache('reminders');
    console.log(`Reminder saved by ${createdBy}: ${noOrder} on ${tanggal}`);
    return true;
  } catch (e) { console.error('saveReminder error:', e.message); return false; }
}

async function markReminderDone(rowIndex) {
  const s = await getSheets();
  if (!s) return;
  try {
    await updateCell(REMINDER_TAB, rowIndex + 2, 5, 'Sent'); // col F = index 5
    clearCache('reminders');
  } catch (e) { console.error('markReminderDone error:', e.message); }
}

// ─── EXCEL UPLOAD → GSHEET ────────────────────────────────
async function handleExcelUpload(fileBuffer, senderName, chatId, isWA = false) {
  try {
    const wb = XLSX.read(fileBuffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 2) return 'File kosong atau tidak ada data.';

    const headers = rows[0];
    const dataRows = rows.slice(1).filter(r => r.some(c => c));

    // Map kolom Excel ke GSheet
    const numIdx    = headers.indexOf('Number');
    const koliIdx   = headers.indexOf('Koli');
    const barangIdx = headers.indexOf('Nama Barang');
    const partnerIdx= headers.indexOf('Partner');
    const noteIdx   = headers.indexOf('Note');
    const nameIdx   = headers.findIndex(h => h && h.toString().includes('First Name'));
    const alamatIdx = headers.findIndex(h => h && h.toString().includes('Line1'));
    const kecIdx    = headers.findIndex(h => h && h.toString().includes('Line3'));
    const kotaIdx   = headers.findIndex(h => h && h.toString().includes('Line4'));
    const phoneIdx  = headers.findIndex(h => h && h.toString().includes('Phone'));
    const posIdx    = headers.findIndex(h => h && h.toString().includes('Postcode'));
    const stateIdx  = headers.findIndex(h => h && h.toString().includes('State'));

    const today  = getToday();
    const cutoff = getCutOff();

    // Build preview
    let preview = `📋 *PREVIEW DATA EXCEL*\n`;
    preview += `Tanggal: ${today} | Cut Off: ${cutoff}\n`;
    preview += `Total: ${dataRows.length} order\n\n`;
    dataRows.forEach((r, i) => {
      preview += `${i+1}. ${r[numIdx] || '—'} - ${r[nameIdx] || '—'} - ${r[partnerIdx] || '—'}\n`;
    });
    preview += `\nMasukkan semua ke GSheet? Balas "ya" untuk konfirmasi.`;

    // Simpan pending data ke memory sementara
    cache['pendingExcel'] = {
      data: dataRows, time: Date.now(), today, cutoff,
      numIdx, koliIdx, barangIdx, partnerIdx, noteIdx,
      nameIdx, alamatIdx, kecIdx, kotaIdx, phoneIdx, posIdx, stateIdx,
      senderName, chatId, isWA
    };

    return preview;
  } catch (e) {
    console.error('handleExcelUpload error:', e.message);
    return 'Gagal membaca file Excel. Pastikan format file benar.';
  }
}

async function insertExcelToSheet(senderName) {
  const pending = cache['pendingExcel'];
  if (!pending) return 'Tidak ada data Excel yang menunggu konfirmasi.';

  const s = await getSheets();
  if (!s) return 'Gagal konek ke Google Sheets.';

  try {
    const { data, today, cutoff, numIdx, koliIdx, barangIdx, partnerIdx, noteIdx, nameIdx, alamatIdx, kecIdx, kotaIdx, phoneIdx, posIdx, stateIdx } = pending;

    const rows = data.map(r => {
      const row = Array(36).fill('');
      row[COL.tanggal]      = today;
      row[COL.cutoff]       = cutoff;
      row[COL.shippingNum]  = r[numIdx] || '';
      row[COL.koli]         = r[koliIdx] || '';
      row[COL.namaBarang]   = r[barangIdx] || '';
      row[COL.partner]      = r[partnerIdx] || '';
      row[COL.customer]     = r[nameIdx] || '';
      row[COL.alamat]       = r[alamatIdx] || '';
      row[COL.kecamatan]    = r[kecIdx] || '';
      row[COL.kota]         = r[kotaIdx] || '';
      row[COL.telepon]      = r[phoneIdx] || '';
      row[COL.kodePos]      = r[posIdx] || '';
      row[COL.provinsi]     = r[stateIdx] || '';
      // N (Ekspedisi) dikosongkan
      return row;
    });

    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!A:AK`,
      valueInputOption: 'RAW', requestBody: { values: rows }
    });

    await logUpdate(senderName, 'BULK INSERT', 'Excel Upload', '-', `${rows.length} order baru`);
    clearCache('sheetData');
    delete cache['pendingExcel'];

    return `✅ ${rows.length} order berhasil dimasukkan ke GSheet!\nTanggal: ${today} | Cut Off: ${cutoff}`;
  } catch (e) {
    console.error('insertExcelToSheet error:', e.message);
    return 'Gagal insert data ke GSheet: ' + e.message;
  }
}

// ─── SMART FILTER ─────────────────────────────────────────
const NO_DATA_INTENTS = ['greeting', 'help', 'out_of_scope'];
const UPDATE_INTENTS = ['update_resi', 'update_tgl_kirim', 'update_tgl_tiba'];

function detectIntent(message) {
  const msg = (message || '').toLowerCase().trim();

  // Greeting & help
  if (/^(halo|hai|hi|hello|selamat|pagi|siang|sore|malam|hey|sup)/.test(msg)) return 'greeting';
  if (/^(oke|ok|ya|tidak|ga|gak|siap|done|sip|noted|thanks|makasih|terima kasih)$/.test(msg)) return 'greeting';
  if (/kamu bisa|fitur apa|help|bantuan|cara pakai|apa saja/.test(msg)) return 'help';

  // Refresh
  if (/refresh|sync data|reload data/.test(msg)) return 'refresh';

  // Log harian
  if (/log hari ini|update hari ini|apa.*diupdate|history update/.test(msg)) return 'log_today';

  // Excel upload confirmation
  if (/^(ya|yes|iya|yep|yup|konfirmasi|insert|masukkan)$/.test(msg) && cache['pendingExcel']) return 'confirm_excel';

  // Update commands (layer 3)
  if (/update resi|input resi|tambah resi|masukkan resi/.test(msg)) return 'update_resi';
  if (/update.*tgl.*kirim|update.*tanggal.*kirim|tanggal pengiriman|tgl pengiriman/.test(msg)) return 'update_tgl_kirim';
  if (/update.*tgl.*tiba|update.*tanggal.*tiba|sudah tiba|tgl tiba/.test(msg)) return 'update_tgl_tiba';

  // Monitoring (layer 2)
  if (/pengiriman hari ini|shipped hari ini|dikirim hari ini/.test(msg)) return 'shipped_today';
  if (/reminder|ingatkan|set reminder/.test(msg)) return 'reminder';
  if (/cek reminder|list reminder|reminder apa/.test(msg)) return 'list_reminder';
  if (/overdue|telat|terlambat|lewat sla/.test(msg)) return 'overdue';
  if (/sla|deadline|mau deadline|mendekati|urgent/.test(msg)) return 'sla_alert';
  if (/pending|waiting|belum dikirim/.test(msg)) return 'pending';
  if (/belum.*resi|tanpa resi|tidak ada resi/.test(msg)) return 'no_resi';
  if (/summary|rangkum|rekap|laporan harian/.test(msg)) return 'summary';
  if (/performa|analisa.*ekspedisi|laporan.*ekspedisi/.test(msg)) return 'analytics';
  if (/jne|j&t|jnt|sicepat|anteraja|ninja|tiki|lion|jnl|deliveree|sentral/.test(msg)) return 'ekspedisi';
  if (/\d{6,}/.test(msg)) return 'specific_order';
  if (/nama|customer|order.*untuk|cari.*nama/.test(msg)) return 'customer_search';

  return 'out_of_scope';
}

function filterData(data, intent, message) {
  if (!data || data.length < 2) return null;
  const rows = data.slice(1);
  const today = getToday();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });

  let filtered;

  if (UPDATE_INTENTS.includes(intent)) {
    // Layer 3: 7 kolom, exclude Received
    filtered = rows.filter(r => (r[COL.status] || '').toLowerCase() !== 'received');
    return { cols: UPDATE_COLS, names: UPDATE_NAMES, rows: selectColumns(filtered, UPDATE_COLS, UPDATE_NAMES) };
  }

  // Layer 2: 12 kolom, filter by intent
  switch (intent) {
    case 'pending':
      filtered = rows.filter(r => /pending|waiting/i.test(r[COL.status] || ''));
      break;
    case 'no_resi':
      filtered = rows.filter(r => !(r[COL.resi] || '').trim() && (r[COL.status] || '').toLowerCase() !== 'received');
      break;
    case 'overdue':
      filtered = rows.filter(r => {
        const sla = r[COL.sla];
        return sla && sla < today && (r[COL.status] || '').toLowerCase() !== 'received';
      });
      break;
    case 'sla_alert':
      filtered = rows.filter(r => {
        const sla = r[COL.sla];
        if (!sla || (r[COL.status] || '').toLowerCase() === 'received') return false;
        const diff = Math.round((new Date(sla) - new Date(today)) / 86400000);
        return diff >= -1 && diff <= 2;
      });
      break;
    case 'shipped_today':
      filtered = rows.filter(r => (r[COL.tglPengiriman] || '').startsWith(today));
      break;
    case 'specific_order':
      const orderNum = (message.match(/\d{6,}/) || [])[0];
      filtered = orderNum ? rows.filter(r => (r[COL.noOrder] || '').includes(orderNum) || (r[COL.shippingNum] || '').includes(orderNum)) : rows.slice(-50);
      break;
    case 'customer_search':
      const words = message.split(/\s+/).filter(w => w.length > 3);
      filtered = rows.filter(r => words.some(w => (r[COL.customer] || '').toLowerCase().includes(w.toLowerCase())));
      break;
    case 'ekspedisi':
      const ekspKeywords = ['jne','j&t','jnt','sicepat','anteraja','ninja','tiki','lion','jnl','deliveree','sentral'];
      const matchEksp = ekspKeywords.find(k => message.toLowerCase().includes(k));
      filtered = matchEksp ? rows.filter(r => (r[COL.ekspedisi] || '').toLowerCase().includes(matchEksp)) : rows.slice(-100);
      break;
    case 'reminder':
    case 'list_reminder':
      return null; // tidak butuh data sheet
    default:
      filtered = rows.slice(-200);
      break;
  }

  console.log(`Filter [${intent}]: ${filtered.length}/${rows.length} rows`);
  return { cols: GENERAL_COLS, names: GENERAL_NAMES, rows: selectColumns(filtered, GENERAL_COLS, GENERAL_NAMES) };
}

// ─── MONITORING ───────────────────────────────────────────
function computeSLAAlerts(data) {
  if (!data || data.length < 2) return null;
  const today = getToday();
  const result = { urgent: [], warning: [], attention: [], overdue: [] };
  data.slice(1).forEach(row => {
    const status = (row[COL.status] || '').toLowerCase();
    if (status === 'received') return;
    const slaStr = row[COL.sla];
    if (!slaStr) return;
    const diff = Math.round((new Date(slaStr) - new Date(today)) / 86400000);
    const info = { no_order: row[COL.noOrder], customer: row[COL.customer], kota: row[COL.kota], status: row[COL.status], sla: slaStr, eksp: row[COL.ekspedisi], resi: row[COL.resi], diff };
    if (diff < 0) result.overdue.push(info);
    else if (diff === 0) result.urgent.push(info);
    else if (diff === 1) result.warning.push(info);
    else if (diff <= 2) result.attention.push(info);
  });
  return result;
}

function computeDailySummary(data) {
  if (!data || data.length < 2) return null;
  const today = getToday();
  const rows = data.slice(1);
  const receivedToday = rows.filter(r => (r[COL.tanggal] || '').startsWith(today)).length;
  const shippedToday  = rows.filter(r => (r[COL.tglPengiriman] || '').startsWith(today)).length;
  const waiting       = rows.filter(r => /waiting|pending/i.test(r[COL.status] || '')).length;
  const noResi        = rows.filter(r => !(r[COL.resi] || '').trim() && (r[COL.status] || '').toLowerCase() !== 'received').length;
  const overdue       = rows.filter(r => { const s = r[COL.sla]; return s && s < today && (r[COL.status] || '').toLowerCase() !== 'received'; }).length;
  const kotaCount = {};
  rows.filter(r => (r[COL.status] || '').toLowerCase() !== 'received').forEach(r => {
    const k = r[COL.kota] || 'Unknown';
    kotaCount[k] = (kotaCount[k] || 0) + 1;
  });
  return { receivedToday, shippedToday, waiting, noResi, overdue, topKota: Object.entries(kotaCount).sort((a,b)=>b[1]-a[1]).slice(0,3) };
}

function computeEkspReport(data) {
  if (!data || data.length < 2) return null;
  const today = getToday();
  const report = {};
  data.slice(1).forEach(row => {
    const eksp = row[COL.ekspedisi] || 'Unknown';
    if (!report[eksp]) report[eksp] = { total: 0, ontime: 0, late: 0, pending: 0 };
    report[eksp].total++;
    const status = (row[COL.status] || '').toLowerCase();
    const sla = row[COL.sla] || '';
    if (status === 'received') {
      (sla && row[COL.tglTiba] && row[COL.tglTiba] <= sla) ? report[eksp].ontime++ : report[eksp].late++;
    } else if (sla && sla < today) { report[eksp].late++; }
    else { report[eksp].pending++; }
  });
  return report;
}

// ─── SYSTEM PROMPT ────────────────────────────────────────
const SYSTEM_PROMPT = `
Kamu adalah OPS Agent untuk Warehouse & Logistik Palembang. Kamu membantu tim operasional mengelola data pengiriman.
Bahasa: Indonesia casual. Nada: Santai tapi profesional.

## BATAS KEMAMPUAN
Kamu HANYA bisa membantu hal-hal berikut:
1. Cek & monitor data order (pending, SLA, overdue, ekspedisi, dll)
2. Update Resi → ACTION:UPDATE_RESI:[no_order]:[resi]
3. Update Tanggal Pengiriman → ACTION:UPDATE_TGL_KIRIM:[no_order]:[YYYY-MM-DD]
4. Update Tanggal Tiba → ACTION:UPDATE_TGL_TIBA:[no_order]:[YYYY-MM-DD]
5. Set reminder → ACTION:SAVE_REMINDER:[no_order]:[YYYY-MM-DD]:[note]
6. Simpan memory → ACTION:SAVE_MEMORY:[category]:[content]
7. Cek log update harian
8. Pengiriman hari ini
9. Summary & laporan
10. Upload Excel (insert order baru)
11. Sapaan & pertanyaan ringan tentang fitur

JIKA pertanyaan di luar list di atas → balas TEPAT: "Hmmm gatau sih, diluar konteks itu keknya 😅"

## FORMAT UPDATE
Selalu konfirmasi dulu sebelum update. Setelah user bilang "ya":
ACTION:UPDATE_RESI:[no_order]:[resi]
ACTION:UPDATE_TGL_KIRIM:[no_order]:[tanggal]
ACTION:UPDATE_TGL_TIBA:[no_order]:[tanggal]

## FORMAT REMINDER
ACTION:SAVE_REMINDER:[no_order]:[YYYY-MM-DD]:[note]

## ATURAN PENTING
- ACTION hanya di akhir pesan, tidak ditampilkan ke user
- JANGAN gunakan LINK WA Customer tanpa perintah eksplisit
- Data sudah difilter relevan, gunakan semua yang tersedia
- Jika data kosong/tidak ditemukan, bilang dengan jelas
`.trim();

// ─── CLAUDE CALL ──────────────────────────────────────────
const chatHistory = {};

async function callClaude(senderId, senderName, userMessage, imageBase64 = null, imageMime = 'image/jpeg') {
  const intent = detectIntent(userMessage || '');
  const today  = getToday();

  // Out of scope → langsung tolak tanpa load data
  if (intent === 'out_of_scope') {
    return 'Hmmm gatau sih, diluar konteks itu keknya 😅';
  }

  // Confirm excel insert
  if (intent === 'confirm_excel') {
    return await insertExcelToSheet(senderName);
  }

  // Refresh cache
  if (intent === 'refresh') {
    clearCache();
    return '🔄 Data berhasil di-refresh! Cache dikosongkan, data terbaru akan diambil dari sheet.';
  }

  // Log today
  if (intent === 'log_today') {
    return await getLogToday();
  }

  // Load data sesuai kebutuhan
  let rawData = null;
  let filteredResult = null;

  if (!NO_DATA_INTENTS.includes(intent)) {
    rawData = await getSheetData();
    filteredResult = filterData(rawData, intent, userMessage || '');
  }

  const [memories, reminders] = await Promise.all([getMemory(), getReminders()]);
  const todayReminders = reminders.filter(r => r[3] === today && r[5] === 'Pending');
  const upcomingReminders = reminders.filter(r => r[3] >= today && r[5] === 'Pending');

  // SLA alerts (hanya untuk intent terkait)
  const slaAlerts = ['sla_alert', 'overdue', 'summary'].includes(intent) ? computeSLAAlerts(rawData) : null;
  const dailySummary = intent === 'summary' ? computeDailySummary(rawData) : null;
  const ekspReport = intent === 'analytics' ? computeEkspReport(rawData) : null;

  // Build dynamic context (SYSTEM_PROMPT di-cache terpisah)
  let context = `Tanggal hari ini (WIB): ${today}\nUser: ${senderName}\n\n`;

  if (memories.length > 0) {
    context += `=== MEMORY ===\n`;
    memories.forEach(m => { context += `[${m[1]}] ${m[2]}\n`; });
    context += '\n';
  }

  if (todayReminders.length > 0) {
    context += `=== REMINDER HARI INI ===\n`;
    todayReminders.forEach(r => { context += `Oleh: ${r[1]} | Order: ${r[2] || '-'} | ${r[4]}\n`; });
    context += '\n';
  }

  if (upcomingReminders.length > 0) {
    context += `=== UPCOMING REMINDERS (semua user) ===\n`;
    upcomingReminders.forEach(r => { context += `[${r[3]}] Oleh: ${r[1]} | Order: ${r[2] || '-'} | ${r[4]}\n`; });
    context += '\n';
  }

  if (slaAlerts) {
    context += `=== SLA ALERTS ===\n`;
    context += `Overdue: ${slaAlerts.overdue.length} | Urgent H-0: ${slaAlerts.urgent.length} | Warning H-1: ${slaAlerts.warning.length} | Attention H-2: ${slaAlerts.attention.length}\n`;
    if (slaAlerts.overdue.length) context += `Overdue: ${JSON.stringify(slaAlerts.overdue)}\n`;
    if (slaAlerts.urgent.length) context += `Urgent: ${JSON.stringify(slaAlerts.urgent)}\n`;
    if (slaAlerts.warning.length) context += `Warning: ${JSON.stringify(slaAlerts.warning)}\n`;
    if (slaAlerts.attention.length) context += `Attention: ${JSON.stringify(slaAlerts.attention)}\n`;
    context += '\n';
  }

  if (dailySummary) context += `=== DAILY SUMMARY ===\n${JSON.stringify(dailySummary)}\n\n`;
  if (ekspReport) context += `=== EKSPEDISI REPORT ===\n${JSON.stringify(ekspReport)}\n\n`;

  if (filteredResult && filteredResult.rows.length > 0) {
    context += `=== DATA ORDER [${intent}, ${filteredResult.rows.length} rows] ===\n`;
    context += JSON.stringify(filteredResult.rows) + '\n';
  } else if (filteredResult && filteredResult.rows.length === 0) {
    context += `=== DATA ORDER ===\nTidak ada data yang sesuai filter.\n`;
  }

  // Build messages
  if (!chatHistory[senderId]) chatHistory[senderId] = [];
  const messages = [...chatHistory[senderId]];

  let userContent;
  if (imageBase64) {
    userContent = [
      { type: 'image', source: { type: 'base64', media_type: imageMime, data: imageBase64 } },
      { type: 'text', text: userMessage || 'Tolong baca nomor resi dari foto ini.' }
    ];
  } else {
    userContent = userMessage;
  }
  messages.push({ role: 'user', content: userContent });

  const response = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1000,
    // Prompt Caching: SYSTEM_PROMPT (static) di-cache, context (dynamic) tidak
    system: [
      { type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } },
      { type: 'text', text: context }
    ],
    messages,
  });
  // Log token usage untuk monitoring
  const usage = response.usage;
  console.log(`[Tokens] in:${usage.input_tokens} out:${usage.output_tokens} cache_write:${usage.cache_creation_input_tokens||0} cache_read:${usage.cache_read_input_tokens||0}`);

  const reply = response.content[0].text;

  chatHistory[senderId].push({ role: 'user', content: userMessage || '[foto]' });
  chatHistory[senderId].push({ role: 'assistant', content: reply });
  if (chatHistory[senderId].length > 12) chatHistory[senderId] = chatHistory[senderId].slice(-12);

  await parseActions(reply, senderId, senderName);
  return reply.replace(/ACTION:[^\n]+/g, '').trim();
}

async function parseActions(text, senderId, senderName) {
  for (const line of text.split('\n')) {
    const t = line.trim();
    if (t.startsWith('ACTION:UPDATE_RESI:')) {
      const p = t.split(':'); if (p[2] && p[3]) await updateOrderField(p[2], COL.resi, 'Resi', p[3], senderName);
    }
    if (t.startsWith('ACTION:UPDATE_TGL_KIRIM:')) {
      const p = t.split(':'); if (p[2] && p[3]) await updateOrderField(p[2], COL.tglPengiriman, 'Tgl Pengiriman', p[3], senderName);
    }
    if (t.startsWith('ACTION:UPDATE_TGL_TIBA:')) {
      const p = t.split(':'); if (p[2] && p[3]) await updateOrderField(p[2], COL.tglTiba, 'Tgl Tiba', p[3], senderName);
    }
    if (t.startsWith('ACTION:SAVE_MEMORY:')) {
      const rest = t.replace('ACTION:SAVE_MEMORY:', ''), idx = rest.indexOf(':');
      if (idx > -1) await saveMemory(rest.substring(0, idx), rest.substring(idx + 1));
    }
    if (t.startsWith('ACTION:SAVE_REMINDER:')) {
      const rest = t.replace('ACTION:SAVE_REMINDER:', ''), parts = rest.split(':');
      if (parts.length >= 3) await saveReminder(senderName, parts[0], parts[1], parts.slice(2).join(':'));
    }
  }
}

// ─── LOG TODAY ────────────────────────────────────────────
async function getLogToday() {
  await ensureTab(UPDATE_LOG_TAB, ['Timestamp', 'Oleh', 'No Order', 'Kolom', 'Nilai Lama', 'Nilai Baru']);
  const s = await getSheets();
  if (!s) return 'Gagal ambil log.';
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${UPDATE_LOG_TAB}!A:F` });
    const rows = (res.data.values || []).slice(1);
    const today = getToday();
    const todayLogs = rows.filter(r => (r[0] || '').includes(today.split('-').reverse().join('/')));
    if (!todayLogs.length) return '📋 Belum ada update hari ini.';
    let msg = `📋 *LOG UPDATE HARI INI*\n\n`;
    todayLogs.forEach(r => { msg += `🕐 ${r[0]} | ${r[1]}\n📦 ${r[2]} | ${r[3]}: ${r[4] || '—'} → ${r[5]}\n\n`; });
    return msg;
  } catch (e) { return 'Gagal ambil log: ' + e.message; }
}

// ─── MESSAGING ────────────────────────────────────────────
async function sendTelegram(chatId, message) {
  try {
    await axios.post(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, { chat_id: chatId, text: message });
  } catch (e) { console.error('sendTelegram error:', e.response?.data || e.message); }
}

async function sendWA(target, message) {
  try {
    await axios.post('https://api.fonnte.com/send', { target, message }, { headers: { Authorization: FONNTE_TOKEN } });
  } catch (e) { console.error('sendWA error:', e.message); }
}

async function getTelegramFile(fileId) {
  const fileRes = await axios.get(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
  const fileUrl = `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${fileRes.data.result.file_path}`;
  const res = await axios.get(fileUrl, { responseType: 'arraybuffer' });
  return Buffer.from(res.data);
}

// ─── WEBHOOKS ─────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status: 'ok', model: 'claude-haiku-4-5', time_wib: new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }) }));
app.get('/webhook/telegram', (_, res) => res.sendStatus(200));
app.get('/webhook/wa', (_, res) => res.sendStatus(200));
app.get('/test-telegram', async (req, res) => {
  const chatId = req.query.chat_id || TELEGRAM_CHAT_ID;
  if (!chatId) return res.json({ error: 'Tambahkan ?chat_id=xxx' });
  await sendTelegram(chatId, `OPS Agent (Claude Haiku) aktif ✅\n${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })}`);
  res.json({ sent: true });
});

// Telegram
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  const msg = req.body.message || req.body.edited_message;
  if (!msg) return;

  const chatId    = msg.chat.id;
  const firstName = msg.from?.first_name || msg.chat?.first_name || 'User';
  const senderId  = `tg_${chatId}`;
  let text = msg.text || msg.caption || '';
  let imageBase64 = null;

  // Handle dokumen Excel
  if (msg.document) {
    const fname = msg.document.file_name || '';
    if (fname.match(/\.(xlsx|xls)$/i)) {
      try {
        const fileBuffer = await getTelegramFile(msg.document.file_id);
        const preview = await handleExcelUpload(fileBuffer, firstName, chatId, false);
        await sendTelegram(chatId, preview);
        return;
      } catch (e) {
        await sendTelegram(chatId, 'Gagal baca file Excel: ' + e.message);
        return;
      }
    }
  }

  // Handle foto
  if (msg.photo && msg.photo.length > 0) {
    try {
      const fileBuffer = await getTelegramFile(msg.photo[msg.photo.length - 1].file_id);
      imageBase64 = fileBuffer.toString('base64');
      if (!text) text = 'Tolong baca nomor resi dari foto ini.';
    } catch (e) { console.error('Photo error:', e.message); }
  }

  if (!text && !imageBase64) return;
  console.log(`TG [${firstName}]: ${text.substring(0, 80)}`);

  try {
    const reply = await callClaude(senderId, firstName, text, imageBase64);
    await sendTelegram(chatId, reply);
  } catch (e) {
    console.error('TG error:', e.message);
    await sendTelegram(chatId, 'Maaf, terjadi error: ' + e.message);
  }
});

// WA
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const { sender, message, name } = req.body;
  if (!sender || !message) return;
  const senderName = name || sender;
  console.log(`WA [${senderName}]: ${message.substring(0, 80)}`);
  try {
    const reply = await callClaude(`wa_${sender}`, senderName, message);
    await sendWA(sender, reply);
  } catch (e) {
    console.error('WA error:', e.message);
    await sendWA(sender, 'Maaf, terjadi error: ' + e.message);
  }
});

// ─── SCHEDULER 08:00 WIB ──────────────────────────────────
cron.schedule('0 8 * * *', async () => {
  console.log('=== DAILY SCHEDULER 08:00 WIB ===');
  try {
    clearCache('sheetData');
    const data  = await getSheetData();
    const today = getToday();
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });

    if (data) {
      // H-1 Pending
      const h1 = data.slice(1).filter(r => (r[COL.requestDate] || '').trim() === tomorrowStr && (r[COL.status] || '').toLowerCase() !== 'received');
      if (h1.length > 0) {
        let msg = `🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
        h1.forEach(r => {
          msg += `📦 ${r[COL.noOrder]} — ${r[COL.customer]} (${r[COL.ekspedisi] || '—'})\n`;
          msg += r[COL.resi] ? `  ✅ Resi: ${r[COL.resi]}\n` : `  ❌ Resi belum diinput!\n`;
        });
        msg += `\nTotal: ${h1.length} order`;
        if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
        if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
      }

      // SLA Alert
      const sla = computeSLAAlerts(data);
      if (sla && (sla.urgent.length > 0 || sla.overdue.length > 0)) {
        let msg = `⚠️ ALERT SLA - ${today}\n\n`;
        if (sla.overdue.length) { msg += `🚨 OVERDUE (${sla.overdue.length}):\n`; sla.overdue.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`; }); msg += '\n'; }
        if (sla.urgent.length)  { msg += `🔴 URGENT H-0 (${sla.urgent.length}):\n`;  sla.urgent.forEach(o => { msg += `• ${o.no_order} - ${o.customer}\n`; }); msg += '\n'; }
        if (sla.warning.length) { msg += `🟡 WARNING H-1 (${sla.warning.length}):\n`; sla.warning.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`; }); }
        if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
        if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
      }
    }

    // Reminders
    clearCache('reminders');
    const reminders = await getReminders();
    const todayReminders = reminders.filter(r => r[3] === today && r[5] === 'Pending');
    for (let i = 0; i < todayReminders.length; i++) {
      const r = todayReminders[i];
      let orderInfo = '';
      if (data && r[2]) {
        const row = data.slice(1).find(d => (d[COL.noOrder] || '').includes(r[2]));
        if (row) orderInfo = `\n📦 ${row[COL.noOrder]} - ${row[COL.customer]}\n   Status: ${row[COL.status] || '—'} | SLA: ${row[COL.sla] || '—'}`;
      }
      const msg = `🔔 REMINDER HARI INI - ${today}\n📌 ${r[4]}${orderInfo}\n\nDibuat oleh: ${r[1]}`;
      if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
      if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
      await markReminderDone(reminders.indexOf(todayReminders[i]));
    }

    console.log('=== SCHEDULER DONE ===');
  } catch (e) { console.error('Scheduler error:', e.message); }
}, { timezone: 'Asia/Jakarta' });

// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS Agent (Claude Haiku 4.5) running on port ${PORT}`);
  console.log(`Features: Smart Filter | 3-Layer | Cache 1H | Excel Upload | Update Log | Shared Reminder`);
});
