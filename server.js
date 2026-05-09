require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');

const app = express();
app.use(express.json());

const ANTHROPIC_KEY    = process.env.ANTHROPIC_API_KEY;
const SHEET_ID         = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB        = process.env.SHEET_TAB_NAME || 'Sheet1';
const FONNTE_TOKEN     = process.env.FONNTE_TOKEN;
const TELEGRAM_TOKEN   = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_NOTIF_CHAT_ID;
const YOUR_WA_NUMBER   = process.env.YOUR_WA_NUMBER;
const PORT             = process.env.PORT || 3000;

console.log('=== ENV CHECK ===');
console.log('ANTHROPIC_API_KEY:', ANTHROPIC_KEY ? 'OK' : 'MISSING');
console.log('GOOGLE_SHEETS_ID:', SHEET_ID ? 'OK' : 'MISSING');
console.log('SHEET_TAB_NAME:', SHEET_TAB);
console.log('FONNTE_TOKEN:', FONNTE_TOKEN ? 'OK' : 'MISSING');
console.log('TELEGRAM_BOT_TOKEN:', TELEGRAM_TOKEN ? 'OK' : 'MISSING');
console.log('TELEGRAM_NOTIF_CHAT_ID:', TELEGRAM_CHAT_ID ? 'OK' : 'MISSING');
console.log('YOUR_WA_NUMBER:', YOUR_WA_NUMBER ? 'OK' : 'MISSING');
console.log('=================');

let sheets = null;
async function getSheets() {
  if (sheets) return sheets;
  try {
    const { google } = require('googleapis');
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheets = google.sheets({ version: 'v4', auth });
    console.log('Google Sheets initialized OK');
    return sheets;
  } catch (e) {
    console.error('Google Sheets init error:', e.message);
    return null;
  }
}

const anthropic = new Anthropic({ apiKey: ANTHROPIC_KEY });

const SYSTEM_PROMPT = `
Kamu adalah Agent Operasional OPS LOG Palembang 2026. Tugasmu membantu tim operasional mengelola data pengiriman.
Bahasa: Indonesia. Nada: Profesional dan singkat.
Jika data sheet tidak tersedia, tetap bantu user semampunya dan informasikan bahwa koneksi sheet sedang bermasalah.
`.trim();

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

const chatHistory = {};

async function callClaude(senderId, userMessage) {
  const sheetData = await getSheetData();
  const today = new Date().toISOString().split('T')[0];
  let dataContext = `Tanggal hari ini: ${today}\n`;
  if (sheetData) {
    dataContext += `Data Google Sheet:\n${JSON.stringify(sheetData, null, 2)}`;
  } else {
    dataContext += `(Koneksi ke Google Sheet sedang tidak tersedia)`;
  }

  if (!chatHistory[senderId]) chatHistory[senderId] = [];
  chatHistory[senderId].push({ role: 'user', content: userMessage });
  const messages = chatHistory[senderId].slice(-10);

  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 1000,
    system: SYSTEM_PROMPT + '\n\n' + dataContext,
    messages,
  });

  const reply = response.content.map(c => c.text || '').join('');
  chatHistory[senderId].push({ role: 'assistant', content: reply });
  return reply;
}

async function sendTelegram(chatId, message) {
  try {
    const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
    const res = await axios.post(url, { chat_id: chatId, text: message });
    console.log('Telegram sent OK to', chatId);
    return res.data;
  } catch (e) {
    console.error('sendTelegram error:', e.response?.data || e.message);
  }
}

async function sendWA(target, message) {
  try {
    await axios.post('https://api.fonnte.com/send', { target, message }, {
      headers: { Authorization: FONNTE_TOKEN },
    });
    console.log('WA sent OK to', target);
  } catch (e) {
    console.error('sendWA error:', e.message);
  }
}

// Health check
app.get('/health', (_, res) => {
  res.json({
    status: 'ok',
    env: {
      anthropic: !!ANTHROPIC_KEY,
      sheets: !!SHEET_ID,
      fonnte: !!FONNTE_TOKEN,
      telegram: !!TELEGRAM_TOKEN,
    }
  });
});

// Test kirim pesan Telegram
app.get('/test-telegram', async (req, res) => {
  const chatId = req.query.chat_id || TELEGRAM_CHAT_ID;
  if (!chatId) return res.json({ error: 'Tambahkan ?chat_id=xxxxx di URL' });
  await sendTelegram(chatId, 'Test dari OPS Agent! Bot aktif ✅');
  res.json({ sent: true, to: chatId });
});
app.get('/webhook/telegram', (_, res) => res.sendStatus(200));
app.get('/webhook/wa', (_, res) => res.sendStatus(200));

// Telegram webhook
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  console.log('Telegram webhook received:', JSON.stringify(req.body));
  const msg = req.body.message || req.body.edited_message;
  if (!msg) { console.log('No message in body'); return; }
  if (!msg.text) { console.log('No text in message'); return; }
  const chatId = msg.chat.id;
  const text = msg.text;
  console.log(`TG [${chatId}]: ${text}`);
  try {
    const reply = await callClaude(`tg_${chatId}`, text);
    await sendTelegram(chatId, reply);
  } catch (e) {
    console.error('Telegram handler error:', e.message);
    await sendTelegram(chatId, 'Maaf, terjadi error: ' + e.message);
  }
});

// WA webhook
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  console.log('WA webhook received:', JSON.stringify(req.body));
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

// H-1 Scheduler jam 08:00 WIB (01:00 UTC)
cron.schedule('0 1 * * *', async () => {
  console.log('Running H-1 check...');
  try {
    const data = await getSheetData();
    if (!data || !data.length) return;
    const headers = data[0];
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toISOString().split('T')[0];
    const noOrderIdx  = headers.indexOf('No Order');
    const customerIdx = headers.indexOf('Nama Customer');
    const reqDateIdx  = headers.indexOf('Request Date');
    const statusIdx   = headers.indexOf('Status');
    const resiIdx     = headers.indexOf('Driver/Booking/Resi');
    const ekspIdx     = headers.indexOf('Ekspedisi');
    const h1Orders = data.slice(1).filter(row =>
      (row[reqDateIdx] || '').trim() === tomorrowStr && (row[statusIdx] || '') !== 'Delivered'
    );
    if (!h1Orders.length) return;
    let msg = `🔔 H-1 PENDING REMINDER\nRequest Date besok: ${tomorrowStr}\n\n`;
    h1Orders.forEach(row => {
      msg += `📦 ${row[noOrderIdx]} — ${row[customerIdx]} (${row[ekspIdx] || '—'})\n`;
      msg += row[resiIdx] ? `  ✅ Resi: ${row[resiIdx]}\n` : `  ❌ Resi belum diinput!\n`;
    });
    msg += `\nTotal: ${h1Orders.length} order`;
    if (TELEGRAM_CHAT_ID) await sendTelegram(TELEGRAM_CHAT_ID, msg);
    if (YOUR_WA_NUMBER) await sendWA(YOUR_WA_NUMBER, msg);
  } catch (e) {
    console.error('H-1 scheduler error:', e.message);
  }
});

app.listen(PORT, () => {
  console.log(`OPS LOG Agent running on port ${PORT}`);
  console.log(`  POST /webhook/wa`);
  console.log(`  POST /webhook/telegram`);
  console.log(`  GET  /health`);
  console.log(`  GET  /test-telegram?chat_id=xxx`);
});
