require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');

const app = express();
app.use(express.json());

// ─── CONFIG ───────────────────────────────────────────────
const ANTHROPIC_KEY    = process.env.ANTHROPIC_API_KEY;
const SHEET_ID         = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB        = process.env.SHEET_TAB_NAME || 'Sheet1';
const MEMORY_TAB       = 'Memory';
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
  } catch (e) {
    console.error('Sheets init error:', e.message);
    return null;
  }
}

function toCol(n) {
  let c = '';
  while (n >= 0) { c = String.fromCharCode(65 + (n % 26)) + c; n = Math.floor(n / 26) - 1; }
  return c;
}

async function getSheetData() {
  const s = await getSheets();
  if (!s) return null;
  try {
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_TAB}!A:AK`,
    });
    return res.data.values || [];
  } catch (e) {
    console.error('getSheetData error:', e.message);
    return null;
  }
}

async function updateResiInSheet(noOrder, resi, ekspedisi) {
  const s = await getSheets();
  if (!s) return false;
  try {
    const data = await getSheetData();
    if (!data) return false;
    const headers  = data[0];
    const noOrderIdx = headers.indexOf('No Order');
    const resiIdx    = headers.indexOf('Driver/Booking/Resi');
    const ekspIdx    = headers.indexOf('Ekspedisi');
    const statusIdx  = headers.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if ((data[i][noOrderIdx] || '').trim() === noOrder.trim()) {
        const row = i + 1;
        await s.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!${toCol(resiIdx)}${row}`, valueInputOption: 'RAW', requestBody: { values: [[resi]] } });
        if (ekspedisi) await s.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!${toCol(ekspIdx)}${row}`, valueInputOption: 'RAW', requestBody: { values: [[ekspedisi]] } });
        if ((data[i][statusIdx] || '') === 'Pending') await s.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!${toCol(statusIdx)}${row}`, valueInputOption: 'RAW', requestBody: { values: [['Diproses']] } });
        console.log(`Resi updated: ${noOrder} → ${resi}`);
        return true;
      }
    }
    console.log(`Order not found: ${noOrder}`);
    return false;
  } catch (e) {
    console.error('updateResi error:', e.message);
    return false;
  }
}

async function updateStatusInSheet(noOrder, status) {
  const s = await getSheets();
  if (!s) return false;
  try {
    const data = await getSheetData();
    if (!data) return false;
    const headers = data[0];
    const noOrderIdx = headers.indexOf('No Order');
    const statusIdx  = headers.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if ((data[i][noOrderIdx] || '').trim() === noOrder.trim()) {
        await s.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `${SHEET_TAB}!${toCol(statusIdx)}${i + 1}`, valueInputOption: 'RAW', requestBody: { values: [[status]] } });
        console.log(`Status updated: ${noOrder} → ${status}`);
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error('updateStatus error:', e.message);
    return false;
  }
}

// ─── MEMORY SYSTEM ────────────────────────────────────────
async function getMemory() {
  const s = await getSheets();
  if (!s) return [];
  try {
    const res = await s.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${MEMORY_TAB}!A:C`,
    });
    const rows = res.data.values || [];
    return rows.slice(1).slice(-30); // ambil 30 memory terakhir
  } catch (e) {
    // Tab Memory belum ada, buat otomatis
    try {
      const { google } = require('googleapis');
      const auth = new google.auth.GoogleAuth({
        credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      const sheetsApi = google.sheets({ version: 'v4', auth });
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: MEMORY_TAB } } }] },
      });
      await sheetsApi.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${MEMORY_TAB}!A1:C1`,
        valueInputOption: 'RAW',
        requestBody: { values: [['Timestamp', 'Category', 'Content']] },
      });
      console.log('Memory tab created');
    } catch (e2) {
      console.error('Create memory tab error:', e2.message);
    }
    return [];
  }
}

async function saveMemory(category, content) {
  const s = await getSheets();
  if (!s) return;
  try {
    const timestamp = new Date().toISOString().replace('T', ' ').split('.')[0];
    await s.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${MEMORY_TAB}!A:C`,
      valueInputOption: 'RAW',
      requestBody: { values: [[timestamp, category, content]] },
    });
    console.log(`Memory saved: [${category}] ${content}`);
  } catch (e) {
    console.error('saveMemory error:', e.message);
  }
}

// ─── CLAUDE ───────────────────────────────────────────────
const anthropic = new Anthropic({ apiKey: ANTHROPIC_KEY });
const chatHistory = {};

const SYSTEM_PROMPT = `
Kamu adalah Agent Operasional OPS LOG Palembang 2026 yang cerdas dan adaptif.
Kamu BISA membaca DAN menulis ke Google Sheet.
Bahasa: Indonesia. Nada: Profesional tapi santai.

## KEMAMPUAN
1. Baca & analisis data order dari Google Sheet
2. Update resi order di Google Sheet
3. Update status order
4. Baca foto/gambar resi dan ekstrak nomor resi otomatis
5. Ingat preferensi dan koreksi user dari waktu ke waktu

## CARA UPDATE RESI
Ketika user minta update resi (via teks atau foto):
- Konfirmasi dulu: "Saya akan simpan:\n• No Order: [x]\n• Ekspedisi: [x]\n• Resi: [x]\nSudah benar?"
- Setelah user bilang "ya" atau konfirmasi, tambahkan di akhir balasan:
  ACTION:UPDATE_RESI:[no_order]:[resi]:[ekspedisi]

## CARA UPDATE STATUS  
Tambahkan di akhir balasan:
ACTION:UPDATE_STATUS:[no_order]:[status]

## CARA SIMPAN MEMORY
Jika user memberikan preferensi, koreksi, atau info bisnis penting, simpan dengan:
ACTION:SAVE_MEMORY:[category]:[content]

Category yang valid: preferensi, bisnis, koreksi, info

Contoh:
- User: "Jawab lebih singkat ya" → ACTION:SAVE_MEMORY:preferensi:User minta jawaban singkat dan to the point
- User: "SLA Ninja 2 hari" → ACTION:SAVE_MEMORY:bisnis:SLA Ninja Express adalah 2 hari
- User: "Kalau order VIP, prioritaskan" → ACTION:SAVE_MEMORY:bisnis:Order VIP harus diprioritaskan

## ATURAN
- ACTION hanya di akhir pesan, jangan tampilkan ke user
- Selalu konfirmasi sebelum update data
- Gunakan memory yang tersedia untuk personalisasi jawaban
- Jika foto dikirim, baca nomor resi dari foto tersebut
`.trim();

async function callClaude(senderId, userMessage, imageBase64 = null, imageMime = 'image/jpeg') {
  const [sheetData, memories] = await Promise.all([getSheetData(), getMemory()]);
  const today = new Date().toISOString().split('T')[0];

  let dataContext = `Tanggal hari ini: ${today}\n\n`;

  if (memories.length > 0) {
    dataContext += `=== MEMORY (preferensi & info yang sudah kamu pelajari) ===\n`;
    memories.forEach(m => { dataContext += `[${m[1]}] ${m[2]}\n`; });
    dataContext += `\n`;
  }

  if (sheetData) {
    const headers = sheetData[0] || [];
    const rows = sheetData.slice(1).slice(-100);
    dataContext += `=== DATA GOOGLE SHEET (100 baris terakhir) ===\nHeader: ${JSON.stringify(headers)}\nRows: ${JSON.stringify(rows)}`;
  } else {
    dataContext += `(Koneksi Google Sheet tidak tersedia)`;
  }

  if (!chatHistory[senderId]) chatHistory[senderId] = [];

  // Build message content
  let messageContent;
  if (imageBase64) {
    messageContent = [
      { type: 'image', source: { type: 'base64', media_type: imageMime, data: imageBase64 } },
      { type: 'text', text: userMessage || 'Tolong baca nomor resi dari foto ini.' }
    ];
  } else {
    messageContent = userMessage;
  }

  chatHistory[senderId].push({ role: 'user', content: messageContent });
  const messages = chatHistory[senderId].slice(-10);

  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 1000,
    system: SYSTEM_PROMPT + '\n\n' + dataContext,
    messages,
  });

  const reply = response.content.map(c => c.text || '').join('');
  chatHistory[senderId].push({ role: 'assistant', content: reply });

  // Parse & execute ACTIONs
  await parseActions(reply);

  // Tampilkan balasan bersih (tanpa baris ACTION)
  const cleanReply = reply.replace(/ACTION:[^\n]+/g, '').trim();
  return cleanReply;
}

async function parseActions(text) {
  const lines = text.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();

    if (trimmed.startsWith('ACTION:UPDATE_RESI:')) {
      const parts = trimmed.split(':');
      const noOrder   = parts[2];
      const resi      = parts[3];
      const ekspedisi = parts.slice(4).join(':').trim();
      if (noOrder && resi) {
        const ok = await updateResiInSheet(noOrder, resi, ekspedisi);
        console.log(`ACTION UPDATE_RESI ${noOrder}: ${ok ? 'OK' : 'FAILED'}`);
      }
    }

    if (trimmed.startsWith('ACTION:UPDATE_STATUS:')) {
      const parts  = trimmed.split(':');
      const noOrder = parts[2];
      const status  = parts[3];
      if (noOrder && status) {
        const ok = await updateStatusInSheet(noOrder, status);
        console.log(`ACTION UPDATE_STATUS ${noOrder}: ${ok ? 'OK' : 'FAILED'}`);
      }
    }

    if (trimmed.startsWith('ACTION:SAVE_MEMORY:')) {
      const rest     = trimmed.replace('ACTION:SAVE_MEMORY:', '');
      const colonIdx = rest.indexOf(':');
      if (colonIdx > -1) {
        const category = rest.substring(0, colonIdx);
        const content  = rest.substring(colonIdx + 1);
        if (category && content) await saveMemory(category, content);
      }
    }
  }
}

// ─── MESSAGING ────────────────────────────────────────────
async function sendTelegram(chatId, message) {
  try {
    await axios.post(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
      chat_id: chatId,
      text: message,
    });
    console.log('Telegram sent to', chatId);
  } catch (e) {
    console.error('sendTelegram error:', e.response?.data || e.message);
  }
}

async function sendWA(target, message) {
  try {
    await axios.post('https://api.fonnte.com/send', { target, message }, {
      headers: { Authorization: FONNTE_TOKEN },
    });
    console.log('WA sent to', target);
  } catch (e) {
    console.error('sendWA error:', e.message);
  }
}

// Download foto dari Telegram dan convert ke base64
async function getTelegramPhotoBase64(fileId) {
  try {
    const fileRes = await axios.get(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/getFile?file_id=${fileId}`);
    const filePath = fileRes.data.result.file_path;
    const fileUrl  = `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${filePath}`;
    const imgRes   = await axios.get(fileUrl, { responseType: 'arraybuffer' });
    const base64   = Buffer.from(imgRes.data).toString('base64');
    return base64;
  } catch (e) {
    console.error('getPhoto error:', e.message);
    return null;
  }
}

// ─── WEBHOOKS ─────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status: 'ok', env: { anthropic: !!ANTHROPIC_KEY, sheets: !!SHEET_ID, fonnte: !!FONNTE_TOKEN, telegram: !!TELEGRAM_TOKEN } }));
app.get('/webhook/telegram', (_, res) => res.sendStatus(200));
app.get('/webhook/wa', (_, res) => res.sendStatus(200));

app.get('/test-telegram', async (req, res) => {
  const chatId = req.query.chat_id || TELEGRAM_CHAT_ID;
  if (!chatId) return res.json({ error: 'Tambahkan ?chat_id=xxx' });
  await sendTelegram(chatId, 'Test dari OPS Agent! Bot aktif ✅');
  res.json({ sent: true, to: chatId });
});

// Telegram webhook — support teks DAN foto
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  console.log('TG webhook:', JSON.stringify(req.body).substring(0, 200));
  const msg = req.body.message || req.body.edited_message;
  if (!msg) return;

  const chatId = msg.chat.id;
  let text = msg.text || msg.caption || '';
  let imageBase64 = null;

  // Kalau ada foto, download dan convert
  if (msg.photo && msg.photo.length > 0) {
    const largestPhoto = msg.photo[msg.photo.length - 1];
    imageBase64 = await getTelegramPhotoBase64(largestPhoto.file_id);
    if (!text) text = 'Tolong baca nomor resi dari foto ini.';
    console.log(`TG [${chatId}] foto diterima`);
  }

  if (!text && !imageBase64) return;
  console.log(`TG [${chatId}]: ${text}`);

  try {
    const reply = await callClaude(`tg_${chatId}`, text, imageBase64);
    await sendTelegram(chatId, reply);
  } catch (e) {
    console.error('TG handler error:', e.message);
    await sendTelegram(chatId, 'Maaf, terjadi error: ' + e.message);
  }
});

// WA webhook
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const { sender, message } = req.body;
  if (!sender || !message) return;
  console.log(`WA [${sender}]: ${message}`);
  try {
    const reply = await callClaude(`wa_${sender}`, message);
    await sendWA(sender, reply);
  } catch (e) {
    console.error('WA handler error:', e.message);
    await sendWA(sender, 'Maaf, terjadi error: ' + e.message);
  }
});

// ─── H-1 SCHEDULER — jam 08:00 WIB (01:00 UTC) ───────────
cron.schedule('0 1 * * *', async () => {
  console.log('Running H-1 check...');
  try {
    const data = await getSheetData();
    if (!data || !data.length) return;
    const headers     = data[0];
    const tomorrow    = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toISOString().split('T')[0];
    const idxMap = {
      noOrder:  headers.indexOf('No Order'),
      customer: headers.indexOf('Nama Customer'),
      reqDate:  headers.indexOf('Request Date'),
      status:   headers.indexOf('Status'),
      resi:     headers.indexOf('Driver/Booking/Resi'),
      eksp:     headers.indexOf('Ekspedisi'),
    };
    const h1 = data.slice(1).filter(row =>
      (row[idxMap.reqDate] || '').trim() === tomorrowStr &&
      (row[idxMap.status] || '') !== 'Delivered'
    );
    if (!h1.length) { console.log('No H-1 orders'); return; }
    let msg = `🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
    h1.forEach(row => {
      msg += `📦 ${row[idxMap.noOrder]} — ${row[idxMap.customer]} (${row[idxMap.eksp] || '—'})\n`;
      msg += row[idxMap.resi] ? `  ✅ Resi: ${row[idxMap.resi]}\n` : `  ❌ Resi belum diinput!\n`;
    });
    msg += `\nTotal: ${h1.length} order`;
    if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
    if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
    console.log(`H-1 notif sent for ${h1.length} orders`);
  } catch (e) {
    console.error('H-1 error:', e.message);
  }
});

// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS LOG Agent running on port ${PORT}`);
  console.log(`  GET  /health`);
  console.log(`  GET  /test-telegram?chat_id=xxx`);
  console.log(`  POST /webhook/telegram`);
  console.log(`  POST /webhook/wa`);
});
