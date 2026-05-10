require('dotenv').config();
const express = require('express');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const axios = require('axios');
const cron = require('node-cron');

const app = express();
app.use(express.json());

// ─── CONFIG ───────────────────────────────────────────────
const GEMINI_KEY       = process.env.GEMINI_API_KEY;
const SHEET_ID         = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB        = process.env.SHEET_TAB_NAME || 'Sheet1';
const MEMORY_TAB       = 'Memory';
const REMINDER_TAB     = 'Reminders';
const FONNTE_TOKEN     = process.env.FONNTE_TOKEN;
const TELEGRAM_TOKEN   = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_NOTIF_CHAT_ID;
const YOUR_WA_NUMBER   = process.env.YOUR_WA_NUMBER;
const PORT             = process.env.PORT || 3000;

console.log('=== ENV CHECK ===');
console.log('GEMINI_API_KEY:', GEMINI_KEY ? 'OK' : 'MISSING');
console.log('GOOGLE_SHEETS_ID:', SHEET_ID ? 'OK' : 'MISSING');
console.log('FONNTE_TOKEN:', FONNTE_TOKEN ? 'OK' : 'MISSING');
console.log('TELEGRAM_BOT_TOKEN:', TELEGRAM_TOKEN ? 'OK' : 'MISSING');
console.log('TZ:', process.env.TZ || 'not set');
console.log('=================');

const genAI = new GoogleGenerativeAI(GEMINI_KEY);

// ─── CACHE (5 menit) ──────────────────────────────────────
const cache = {};
const CACHE_TTL = 5 * 60 * 1000;

function getCache(key) {
  const c = cache[key];
  if (c && Date.now() - c.time < CACHE_TTL) return c.data;
  return null;
}
function setCache(key, data) {
  cache[key] = { data, time: Date.now() };
}
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

async function getSheetData() {
  const cached = getCache('sheetData');
  if (cached) { console.log('Sheet data from cache'); return cached; }
  const s = await getSheets();
  if (!s) return null;
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!A:AK` });
    const data = res.data.values || [];
    setCache('sheetData', data);
    console.log(`Sheet data loaded: ${data.length} rows`);
    return data;
  } catch (e) { console.error('getSheetData error:', e.message); return null; }
}

async function updateCell(sheetTab, row, colIdx, value) {
  const s = await getSheets();
  if (!s) return;
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTab}!${toCol(colIdx)}${row}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[value]] },
  });
  clearCache('sheetData'); // invalidate cache setelah update
}

async function updateResiInSheet(noOrder, resi, ekspedisi) {
  const s = await getSheets();
  if (!s) return false;
  try {
    const data = await getSheetData();
    if (!data) return false;
    const h = data[0];
    const noOrderIdx = h.indexOf('No Order');
    const resiIdx    = h.indexOf('Driver/Booking/Resi');
    const ekspIdx    = h.indexOf('Ekspedisi');
    const statusIdx  = h.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if ((data[i][noOrderIdx] || '').trim() === noOrder.trim()) {
        await updateCell(SHEET_TAB, i + 1, resiIdx, resi);
        if (ekspedisi) await updateCell(SHEET_TAB, i + 1, ekspIdx, ekspedisi);
        if ((data[i][statusIdx] || '') === 'Pending') await updateCell(SHEET_TAB, i + 1, statusIdx, 'Diproses');
        console.log(`Resi updated: ${noOrder} → ${resi}`);
        return true;
      }
    }
    console.log(`Order not found: ${noOrder}`);
    return false;
  } catch (e) { console.error('updateResi error:', e.message); return false; }
}

async function updateStatusInSheet(noOrder, status) {
  const s = await getSheets();
  if (!s) return false;
  try {
    const data = await getSheetData();
    if (!data) return false;
    const h = data[0];
    const noOrderIdx = h.indexOf('No Order');
    const statusIdx  = h.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if ((data[i][noOrderIdx] || '').trim() === noOrder.trim()) {
        await updateCell(SHEET_TAB, i + 1, statusIdx, status);
        console.log(`Status updated: ${noOrder} → ${status}`);
        return true;
      }
    }
    return false;
  } catch (e) { console.error('updateStatus error:', e.message); return false; }
}

// ─── MEMORY & REMINDER ────────────────────────────────────
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
    console.log(`Memory: [${category}] ${content}`);
  } catch (e) { console.error('saveMemory error:', e.message); }
}

async function getReminders() {
  const cached = getCache('reminders');
  if (cached) return cached;
  await ensureTab(REMINDER_TAB, ['Timestamp', 'No Order', 'Tanggal Reminder', 'Note', 'Status']);
  const s = await getSheets();
  if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:E` });
    const data = (res.data.values || []).slice(1);
    setCache('reminders', data);
    return data;
  } catch (e) { return []; }
}

async function saveReminder(noOrder, tanggal, note) {
  await ensureTab(REMINDER_TAB, ['Timestamp', 'No Order', 'Tanggal Reminder', 'Note', 'Status']);
  const s = await getSheets();
  if (!s) return false;
  try {
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    await s.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: `${REMINDER_TAB}!A:E`, valueInputOption: 'RAW', requestBody: { values: [[ts, noOrder || '', tanggal, note, 'Pending']] } });
    clearCache('reminders');
    console.log(`Reminder saved: ${noOrder} on ${tanggal}`);
    return true;
  } catch (e) { console.error('saveReminder error:', e.message); return false; }
}

async function markReminderDone(rowIndex) {
  const s = await getSheets();
  if (!s) return;
  try {
    await updateCell(REMINDER_TAB, rowIndex + 2, 4, 'Sent');
    clearCache('reminders');
  } catch (e) { console.error('markReminderDone error:', e.message); }
}

// ─── SMART FILTER ─────────────────────────────────────────
function detectIntent(message) {
  const msg = message.toLowerCase();
  if (/reminder|ingatkan|set reminder/.test(msg)) return 'reminder';
  if (/\d{6,}/.test(msg)) return 'specific_order';
  if (/overdue|telat|terlambat|lewat sla|melewati/.test(msg)) return 'overdue';
  if (/sla|deadline|mau deadline|mendekati|urgent|warning/.test(msg)) return 'sla_alert';
  if (/pending|waiting|belum dikirim|belum kirim/.test(msg)) return 'pending';
  if (/belum.*resi|tanpa resi|tidak ada resi/.test(msg)) return 'no_resi';
  if (/summary|rangkum|rekap|laporan harian|hari ini/.test(msg)) return 'summary';
  if (/performa|analisa|report.*ekspedisi|laporan.*ekspedisi/.test(msg)) return 'analytics';
  if (/resi|update resi|input resi/.test(msg)) return 'resi';
  if (/jne|j&t|jnt|sicepat|anteraja|ninja|tiki|lion|jnl|deliveree|sentral/.test(msg)) return 'ekspedisi';
  if (/nama|customer/.test(msg)) return 'customer_search';
  return 'general';
}

function filterByIntent(data, intent, message) {
  if (!data || data.length < 2) return data;
  const h = data[0];
  const rows = data.slice(1);
  const today = new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });
  const idx = {
    noOrder:  h.indexOf('No Order'),
    customer: h.indexOf('Nama Customer'),
    status:   h.indexOf('Status'),
    resi:     h.indexOf('Driver/Booking/Resi'),
    eksp:     h.indexOf('Ekspedisi'),
    tglWajib: h.indexOf('Tgl Wajib Kirim (SLA)'),
  };
  let filtered;
  switch (intent) {
    case 'pending':
      filtered = rows.filter(r => /pending|waiting|belum/i.test(r[idx.status] || ''));
      break;
    case 'no_resi':
      filtered = rows.filter(r => !(r[idx.resi] || '').trim() && (r[idx.status] || '') !== 'Delivered');
      break;
    case 'overdue':
      filtered = rows.filter(r => {
        const sla = r[idx.tglWajib];
        return sla && sla < today && (r[idx.status] || '') !== 'Delivered';
      });
      break;
    case 'sla_alert':
      filtered = rows.filter(r => {
        const sla = r[idx.tglWajib];
        if (!sla || (r[idx.status] || '') === 'Delivered') return false;
        const diff = Math.round((new Date(sla) - new Date(today)) / 86400000);
        return diff >= -1 && diff <= 3;
      });
      break;
    case 'specific_order':
      const orderNum = (message.match(/\d{6,}/) || [])[0];
      filtered = orderNum ? rows.filter(r => (r[idx.noOrder] || '').includes(orderNum)) : rows.slice(-50);
      break;
    case 'customer_search':
      const words = message.split(/\s+/).filter(w => w.length > 3);
      filtered = rows.filter(r => words.some(w => (r[idx.customer] || '').toLowerCase().includes(w.toLowerCase())));
      break;
    case 'ekspedisi':
      const ekspKeywords = ['jne','j&t','jnt','sicepat','anteraja','ninja','tiki','lion','jnl','deliveree','sentral'];
      const matchEksp = ekspKeywords.find(k => message.toLowerCase().includes(k));
      filtered = matchEksp ? rows.filter(r => (r[idx.eksp] || '').toLowerCase().includes(matchEksp)) : rows.slice(-100);
      break;
    default:
      filtered = rows.slice(-200);
      break;
  }
  console.log(`Smart filter [${intent}]: ${filtered.length}/${rows.length} rows`);
  return [h, ...filtered];
}

// ─── MONITORING ───────────────────────────────────────────
function getToday() {
  return new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });
}

function computeSLAAlerts(data) {
  if (!data || data.length < 2) return null;
  const h = data[0];
  const today = getToday();
  const idx = {
    noOrder: h.indexOf('No Order'), customer: h.indexOf('Nama Customer'),
    kota: h.indexOf('Kota / Kabupaten'), status: h.indexOf('Status'),
    eksp: h.indexOf('Ekspedisi'), produk: h.indexOf('Nama Barang'),
    resi: h.indexOf('Driver/Booking/Resi'), tglWajib: h.indexOf('Tgl Wajib Kirim (SLA)'),
  };
  const result = { urgent: [], warning: [], attention: [], overdue: [] };
  data.slice(1).forEach(row => {
    const status = row[idx.status] || '';
    if (status === 'Delivered') return;
    const slaStr = row[idx.tglWajib];
    if (!slaStr) return;
    const diff = Math.round((new Date(slaStr) - new Date(today)) / 86400000);
    const info = { no_order: row[idx.noOrder], customer: row[idx.customer], kota: row[idx.kota], status, sla: slaStr, eksp: row[idx.eksp], produk: row[idx.produk], resi: row[idx.resi], diff };
    if (diff < 0) result.overdue.push(info);
    else if (diff === 0) result.urgent.push(info);
    else if (diff === 1) result.warning.push(info);
    else if (diff <= 2) result.attention.push(info);
  });
  return result;
}

function computeDailySummary(data) {
  if (!data || data.length < 2) return null;
  const h = data[0];
  const today = getToday();
  const idx = {
    tanggal: h.indexOf('Tanggal'), status: h.indexOf('Status'),
    resi: h.indexOf('Driver/Booking/Resi'), kota: h.indexOf('Kota / Kabupaten'),
    tglWajib: h.indexOf('Tgl Wajib Kirim (SLA)'),
  };
  const rows = data.slice(1);
  const receivedToday = rows.filter(r => (r[idx.tanggal] || '').startsWith(today));
  const shippedToday  = rows.filter(r => r[idx.status] === 'Delivered' && (r[idx.tanggal] || '').startsWith(today));
  const waiting       = rows.filter(r => /waiting|pending/i.test(r[idx.status] || ''));
  const noResi        = rows.filter(r => !(r[idx.resi] || '').trim() && r[idx.status] !== 'Delivered');
  const overdue       = rows.filter(r => { const sla = r[idx.tglWajib]; return sla && sla < today && r[idx.status] !== 'Delivered'; });
  const kotaCount = {};
  rows.filter(r => r[idx.status] !== 'Delivered').forEach(r => {
    const kota = r[idx.kota] || 'Unknown';
    kotaCount[kota] = (kotaCount[kota] || 0) + 1;
  });
  return { receivedToday: receivedToday.length, shippedToday: shippedToday.length, waiting: waiting.length, noResi: noResi.length, overdue: overdue.length, topKota: Object.entries(kotaCount).sort((a,b) => b[1]-a[1]).slice(0,3) };
}

function computeEkspedisiReport(data) {
  if (!data || data.length < 2) return null;
  const h = data[0];
  const today = getToday();
  const idx = { eksp: h.indexOf('Ekspedisi'), status: h.indexOf('Status'), tglWajib: h.indexOf('Tgl Wajib Kirim (SLA)'), tanggalTiba: h.indexOf('Tanggal Tiba'), aging: h.indexOf('Aging') };
  const report = {};
  data.slice(1).forEach(row => {
    const eksp = row[idx.eksp] || 'Unknown';
    if (!report[eksp]) report[eksp] = { total: 0, ontime: 0, late: 0, pending: 0 };
    report[eksp].total++;
    const status = row[idx.status] || '';
    const sla = row[idx.tglWajib] || '';
    if (status === 'Delivered') {
      (sla && (row[idx.tanggalTiba] || '') <= sla) ? report[eksp].ontime++ : report[eksp].late++;
    } else if (sla && sla < today) { report[eksp].late++; }
    else { report[eksp].pending++; }
  });
  return report;
}

// ─── SYSTEM PROMPT ────────────────────────────────────────
const SYSTEM_PROMPT = `
Kamu adalah Agent Operasional OPS LOG Palembang yang cerdas dan adaptif.
Kamu BISA membaca DAN menulis ke Google Sheet.
Bahasa: Indonesia. Nada: Profesional tapi santai.

## KEMAMPUAN
1. Analisis data order dari Google Sheet
2. Update resi → ACTION:UPDATE_RESI:[no_order]:[resi]:[ekspedisi]
3. Update status → ACTION:UPDATE_STATUS:[no_order]:[status]
4. Simpan memory → ACTION:SAVE_MEMORY:[category]:[content]
5. Set reminder → ACTION:SAVE_REMINDER:[no_order]:[YYYY-MM-DD]:[note]
6. Baca foto resi (OCR)
7. Monitor SLA, overdue, summary, performa ekspedisi

## FORMAT MONITORING
Alert SLA: 🔴 URGENT (H-0) | 🟡 WARNING (H-1) | 🟠 ATTENTION (H-2) | 🚨 OVERDUE
Summary: 📊 SUMMARY | ✅ Received | 📦 Shipped | ⏳ Waiting | 🚨 Needs Attention
Ekspedisi: 📈 PERFORMA [eksp] | Total | On-time | Late | Pending
Reminder: 🔔 REMINDER [tgl] | 📌 [note] | 📦 [order info]

## ATURAN
- ACTION hanya di akhir pesan, jangan tampilkan
- JANGAN gunakan LINK WA Customer tanpa perintah eksplisit
- Konfirmasi sebelum update data
- Gunakan data monitoring yang sudah dicompute di konteks
`.trim();

// ─── GEMINI ───────────────────────────────────────────────
const chatHistory = {};

async function callGemini(senderId, userMessage, imageBase64 = null, imageMime = 'image/jpeg') {
  const intent = detectIntent(userMessage || '');
  const today  = getToday();

  // Load semua data paralel (dari cache kalau ada)
  const [rawData, memories, reminders] = await Promise.all([getSheetData(), getMemory(), getReminders()]);
  const filteredData = filterByIntent(rawData, intent, userMessage || '');

  // Pre-compute hanya yang dibutuhkan
  const slaAlerts    = computeSLAAlerts(rawData);
  const dailySummary = (intent === 'summary') ? computeDailySummary(rawData) : null;
  const ekspReport   = (intent === 'analytics') ? computeEkspedisiReport(rawData) : null;
  const todayReminders = reminders.filter(r => r[2] === today && r[4] === 'Pending');
  const upcomingReminders = reminders.filter(r => r[2] >= today && r[4] === 'Pending');

  let context = SYSTEM_PROMPT + `\n\nTanggal hari ini (WIB): ${today}\n\n`;

  if (memories.length > 0) {
    context += `=== MEMORY ===\n`;
    memories.forEach(m => { context += `[${m[1]}] ${m[2]}\n`; });
    context += '\n';
  }
  if (todayReminders.length > 0) {
    context += `=== REMINDER HARI INI ===\n`;
    todayReminders.forEach(r => { context += `Order: ${r[1] || '-'} | ${r[3]}\n`; });
    context += '\n';
  }
  if (upcomingReminders.length > 0) {
    context += `=== UPCOMING REMINDERS ===\n`;
    upcomingReminders.forEach(r => { context += `[${r[2]}] Order: ${r[1] || '-'} | ${r[3]}\n`; });
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
  if (ekspReport)   context += `=== EKSPEDISI REPORT ===\n${JSON.stringify(ekspReport)}\n\n`;
  if (filteredData && filteredData.length > 1) {
    context += `=== DATA SHEET [${intent}, ${filteredData.length - 1} rows] ===\n${JSON.stringify(filteredData)}\n`;
  }

  if (!chatHistory[senderId]) chatHistory[senderId] = [];
  const history = chatHistory[senderId].map(msg => ({
    role: msg.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: typeof msg.content === 'string' ? msg.content : JSON.stringify(msg.content) }],
  }));

  const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash-lite', systemInstruction: context });
  const chat  = model.startChat({ history });

  let parts = [];
  if (imageBase64) parts.push({ inlineData: { data: imageBase64, mimeType: imageMime } });
  parts.push({ text: userMessage || 'Tolong baca nomor resi dari foto ini.' });

  const result = await chat.sendMessage(parts);
  const reply  = result.response.text();

  chatHistory[senderId].push({ role: 'user', content: userMessage || '[foto]' });
  chatHistory[senderId].push({ role: 'assistant', content: reply });
  if (chatHistory[senderId].length > 20) chatHistory[senderId] = chatHistory[senderId].slice(-20);

  await parseActions(reply);
  return reply.replace(/ACTION:[^\n]+/g, '').trim();
}

async function parseActions(text) {
  for (const line of text.split('\n')) {
    const t = line.trim();
    if (t.startsWith('ACTION:UPDATE_RESI:')) {
      const p = t.split(':');
      if (p[2] && p[3]) await updateResiInSheet(p[2], p[3], p.slice(4).join(':').trim());
    }
    if (t.startsWith('ACTION:UPDATE_STATUS:')) {
      const p = t.split(':');
      if (p[2] && p[3]) await updateStatusInSheet(p[2], p[3]);
    }
    if (t.startsWith('ACTION:SAVE_MEMORY:')) {
      const rest = t.replace('ACTION:SAVE_MEMORY:', '');
      const idx  = rest.indexOf(':');
      if (idx > -1) await saveMemory(rest.substring(0, idx), rest.substring(idx + 1));
    }
    if (t.startsWith('ACTION:SAVE_REMINDER:')) {
      const rest  = t.replace('ACTION:SAVE_REMINDER:', '');
      const parts = rest.split(':');
      if (parts.length >= 3) await saveReminder(parts[0], parts[1], parts.slice(2).join(':'));
    }
  }
}

// ─── MESSAGING ────────────────────────────────────────────
async function sendTelegram(chatId, message) {
  try {
    await axios.post(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, { chat_id: chatId, text: message });
    console.log('TG sent to', chatId);
  } catch (e) { console.error('sendTelegram error:', e.response?.data || e.message); }
}

async function sendWA(target, message) {
  try {
    await axios.post('https://api.fonnte.com/send', { target, message }, { headers: { Authorization: FONNTE_TOKEN } });
    console.log('WA sent to', target);
  } catch (e) { console.error('sendWA error:', e.message); }
}

async function getTelegramPhotoBase64(fileId) {
  try {
    const fileRes = await axios.get(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
    const fileUrl = `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${fileRes.data.result.file_path}`;
    const imgRes  = await axios.get(fileUrl, { responseType: 'arraybuffer' });
    return Buffer.from(imgRes.data).toString('base64');
  } catch (e) { console.error('getPhoto error:', e.message); return null; }
}

// ─── WEBHOOKS ─────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status: 'ok', model: 'gemini-2.5-flash-lite', timezone: process.env.TZ || 'UTC', time_wib: new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }) }));
app.get('/webhook/telegram', (_, res) => res.sendStatus(200));
app.get('/webhook/wa', (_, res) => res.sendStatus(200));
app.get('/test-telegram', async (req, res) => {
  const chatId = req.query.chat_id || TELEGRAM_CHAT_ID;
  if (!chatId) return res.json({ error: 'Tambahkan ?chat_id=xxx' });
  await sendTelegram(chatId, `OPS Agent aktif ✅\nWaktu WIB: ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })}`);
  res.json({ sent: true });
});

app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  const msg = req.body.message || req.body.edited_message;
  if (!msg) return;
  const chatId = msg.chat.id;
  let text = msg.text || msg.caption || '';
  let imageBase64 = null;
  if (msg.photo && msg.photo.length > 0) {
    imageBase64 = await getTelegramPhotoBase64(msg.photo[msg.photo.length - 1].file_id);
    if (!text) text = 'Tolong baca nomor resi dari foto ini.';
  }
  if (!text && !imageBase64) return;
  console.log(`TG [${chatId}]: ${text.substring(0, 80)}`);
  try {
    const reply = await callGemini(`tg_${chatId}`, text, imageBase64);
    await sendTelegram(chatId, reply);
  } catch (e) {
    console.error('TG error:', e.message);
    await sendTelegram(chatId, 'Maaf, terjadi error: ' + e.message);
  }
});

app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const { sender, message } = req.body;
  if (!sender || !message) return;
  console.log(`WA [${sender}]: ${message.substring(0, 80)}`);
  try {
    const reply = await callGemini(`wa_${sender}`, message);
    await sendWA(sender, reply);
  } catch (e) {
    console.error('WA error:', e.message);
    await sendWA(sender, 'Maaf, terjadi error: ' + e.message);
  }
});

// ─── SCHEDULER — 08:00 WIB (dengan TZ=Asia/Jakarta) ──────
cron.schedule('0 8 * * *', async () => {
  console.log('=== DAILY SCHEDULER 08:00 WIB ===');
  try {
    const data  = await getSheetData();
    const today = getToday();
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });

    if (data) {
      const h = data[0];
      const idx = { noOrder: h.indexOf('No Order'), customer: h.indexOf('Nama Customer'), reqDate: h.indexOf('Request Date'), status: h.indexOf('Status'), resi: h.indexOf('Driver/Booking/Resi'), eksp: h.indexOf('Ekspedisi') };

      // H-1 Pending
      const h1 = data.slice(1).filter(r => (r[idx.reqDate] || '').trim() === tomorrowStr && (r[idx.status] || '') !== 'Delivered');
      if (h1.length > 0) {
        let msg = `🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
        h1.forEach(r => {
          msg += `📦 ${r[idx.noOrder]} — ${r[idx.customer]} (${r[idx.eksp] || '—'})\n`;
          msg += r[idx.resi] ? `  ✅ Resi: ${r[idx.resi]}\n` : `  ❌ Resi belum diinput!\n`;
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
        if (sla.urgent.length)  { msg += `🔴 URGENT H-0 (${sla.urgent.length}):\n`;  sla.urgent.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | ${o.eksp || '—'}\n`; }); msg += '\n'; }
        if (sla.warning.length) { msg += `🟡 WARNING H-1 (${sla.warning.length}):\n`; sla.warning.forEach(o => { msg += `• ${o.no_order} - ${o.customer} | SLA: ${o.sla}\n`; }); }
        if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
        if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
      }
    }

    // Reminder check
    clearCache('reminders');
    const reminders = await getReminders();
    const todayReminders = reminders.filter(r => r[2] === today && r[4] === 'Pending');
    for (let i = 0; i < todayReminders.length; i++) {
      const r = todayReminders[i];
      let orderInfo = '';
      if (data && r[1]) {
        const h = data[0];
        const idx = { noOrder: h.indexOf('No Order'), customer: h.indexOf('Nama Customer'), status: h.indexOf('Status'), produk: h.indexOf('Nama Barang'), tglWajib: h.indexOf('Tgl Wajib Kirim (SLA)') };
        const row = data.slice(1).find(row => (row[idx.noOrder] || '').includes(r[1]));
        if (row) orderInfo = `\n📦 ${row[idx.noOrder]} - ${row[idx.customer]}\n   Produk: ${row[idx.produk] || '—'}\n   Status: ${row[idx.status] || '—'}\n   SLA: ${row[idx.tglWajib] || '—'}`;
      }
      const msg = `🔔 REMINDER HARI INI - ${today}\n📌 ${r[3]}${orderInfo}`;
      if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
      await markReminderDone(reminders.indexOf(todayReminders[i]));
    }

    console.log('=== SCHEDULER DONE ===');
  } catch (e) { console.error('Scheduler error:', e.message); }
}, { timezone: 'Asia/Jakarta' });

// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS LOG Agent running on port ${PORT}`);
  console.log(`Timezone: ${process.env.TZ || 'UTC (set TZ=Asia/Jakarta di Railway)'}`);
  console.log(`Model: gemini-2.5-flash-lite`);
});
