// ============================================================
// OPS LOG PALEMBANG — Agent Backend
// Integrasi: Claude API + Google Sheets + Fonnte (WA) + Telegram
// ============================================================

require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const Anthropic = require('@anthropic-ai/sdk');
const axios = require('axios');
const cron = require('node-cron');

const app = express();
app.use(express.json());

// ─── CONFIG ───────────────────────────────────────────────
const ANTHROPIC_KEY    = process.env.ANTHROPIC_API_KEY;
const SHEET_ID         = process.env.GOOGLE_SHEETS_ID;
const SHEET_TAB        = process.env.SHEET_TAB_NAME || 'Sheet1';
const FONNTE_TOKEN     = process.env.FONNTE_TOKEN;
const TELEGRAM_TOKEN   = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_NOTIF_CHAT_ID; // chat ID untuk notifikasi H-1
const YOUR_WA_NUMBER   = process.env.YOUR_WA_NUMBER;          // no WA kamu untuk notifikasi H-1
const PORT             = process.env.PORT || 3000;

// ─── CLIENTS ──────────────────────────────────────────────
const anthropic = new Anthropic({ apiKey: ANTHROPIC_KEY });

const googleAuth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth: googleAuth });

// ─── SYSTEM PROMPT ────────────────────────────────────────
const SYSTEM_PROMPT = `
Kamu adalah Agent Operasional OPS LOG Palembang 2026. Tugasmu membantu tim operasional mengelola data pengiriman yang tersimpan di Google Sheets.

## IDENTITAS
Nama: OPS Agent Palembang
Bahasa: Indonesia (selalu)
Nada: Profesional, singkat, dan to the point

## KEMAMPUAN UTAMA

### 1. Input & Update Resi
Ketika user memberikan informasi resi, ekstrak:
- No Order (contoh: ORD-PLM-001)
- Nomor Resi
- Ekspedisi (JNE, J&T, SiCepat, AnterAja, dll)

Konfirmasi data sebelum menyimpan:
"Saya akan menyimpan:
• No Order: [no_order]
• Ekspedisi: [ekspedisi]
• Resi: [nomor_resi]
Sudah benar? Balas 'ya' untuk konfirmasi."

Setelah dikonfirmasi, balas: "ACTION:UPDATE_RESI:[no_order]:[resi]:[ekspedisi]"

### 2. Monitor SLA & Aging
Definisi:
- AGING = order yang Tanggal Tiba sudah melewati SLA Max
- SLA WARNING = order yang mendekati batas SLA (H-1 atau H-0 dari Tgl Wajib Kirim)
- H-1 PENDING = order berstatus Pending/Belum Dikirim dengan Request Date = besok

Format laporan:
🔴 AGING (X order): [list no order + customer]
🟡 SLA WARNING (X order): [list no order + customer]
🟠 H-1 PENDING (X order): [list no order + customer]
✅ ON TRACK (X order)

### 3. Cek & Filter Data
Jawab pertanyaan seperti:
- "Order mana yang belum ada resi?"
- "Tampilkan semua order ekspedisi JNE"
- "Status order ORD-PLM-XXX?"
- "Berapa order yang sudah delivered hari ini?"

### 4. Update Status
Untuk update status, gunakan format:
"ACTION:UPDATE_STATUS:[no_order]:[status_baru]"

Status yang valid: Pending, Diproses, Dalam Perjalanan, Delivered, Aging, SLA Warning

## KOLOM DATA
Tanggal, Cutoff, Shipping Number, Koli, Nama Barang, Partner, Nama Customer, Alamat,
Kecamatan, Kota/Kabupaten, No. Telepon, Kode Pos, Provinsi, Ekspedisi,
Driver/Booking/Resi, Tanggal Pengiriman, Tanggal Tiba, Aging, Request Date,
Remark, Status, SLA, No Order, SLA Min, SLA Max, Tgl WAJIB Kirim (Order Create),
Tgl Wajib Kirim (SLA), Shipping Status, tanggal_created, Tgl WAJIB Kirim (By Req),
Request Date Pokedex, Request Compare, Aging SLA, Group SLA

## ATURAN PENTING
- Selalu konfirmasi sebelum melakukan perubahan data
- Format tanggal selalu YYYY-MM-DD
- Jika ada H-1 Pending yang resi-nya belum ada, ingatkan user
- Gunakan emoji secukupnya: ✅ 🔴 🟡 🟠 📦
- Jawaban singkat dan terstruktur
`.trim();

// ─── GOOGLE SHEETS HELPERS ────────────────────────────────
async function getSheetData() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A:AK`,
  });
  return res.data.values || [];
}

function colIndex(headers, name) {
  return headers.indexOf(name);
}

function numToCol(n) {
  let col = '';
  while (n >= 0) {
    col = String.fromCharCode(65 + (n % 26)) + col;
    n = Math.floor(n / 26) - 1;
  }
  return col;
}

async function updateCell(row, col, value) {
  const range = `${SHEET_TAB}!${numToCol(col)}${row}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values: [[value]] },
  });
}

async function updateResiInSheet(noOrder, resi, ekspedisi) {
  const data = await getSheetData();
  if (!data.length) return false;
  const headers = data[0];
  const noOrderIdx = colIndex(headers, 'No Order');
  const resiIdx    = colIndex(headers, 'Driver/Booking/Resi');
  const ekspIdx    = colIndex(headers, 'Ekspedisi');
  const statusIdx  = colIndex(headers, 'Status');

  for (let i = 1; i < data.length; i++) {
    if (data[i][noOrderIdx] === noOrder) {
      const sheetRow = i + 1;
      await updateCell(sheetRow, resiIdx, resi);
      if (ekspedisi) await updateCell(sheetRow, ekspIdx, ekspedisi);
      if (data[i][statusIdx] === 'Pending') {
        await updateCell(sheetRow, statusIdx, 'Diproses');
      }
      return true;
    }
  }
  return false;
}

async function updateStatusInSheet(noOrder, status) {
  const data = await getSheetData();
  if (!data.length) return false;
  const headers = data[0];
  const noOrderIdx = colIndex(headers, 'No Order');
  const statusIdx  = colIndex(headers, 'Status');

  for (let i = 1; i < data.length; i++) {
    if (data[i][noOrderIdx] === noOrder) {
      await updateCell(i + 1, statusIdx, status);
      return true;
    }
  }
  return false;
}

// ─── CLAUDE HELPER ────────────────────────────────────────
const conversationHistory = {}; // { senderId: [messages] }

async function callClaude(senderId, userMessage) {
  const sheetData = await getSheetData();
  const today = new Date().toISOString().split('T')[0];
  const dataContext = `\nTanggal hari ini: ${today}\n\nData Google Sheet:\n${JSON.stringify(sheetData, null, 2)}`;

  if (!conversationHistory[senderId]) conversationHistory[senderId] = [];
  conversationHistory[senderId].push({ role: 'user', content: userMessage });

  // Keep last 10 messages to avoid token overflow
  const messages = conversationHistory[senderId].slice(-10);

  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1000,
    system: SYSTEM_PROMPT + dataContext,
    messages,
  });

  const replyText = response.content.map(c => c.text || '').join('');
  conversationHistory[senderId].push({ role: 'assistant', content: replyText });

  // Parse ACTION commands from Claude's response
  await parseAndExecuteActions(replyText);

  // Clean display text (remove ACTION: lines from reply to user)
  const displayText = replyText.replace(/ACTION:[^\n]+/g, '').trim();
  return displayText;
}

async function parseAndExecuteActions(text) {
  const lines = text.split('\n');
  for (const line of lines) {
    if (line.startsWith('ACTION:UPDATE_RESI:')) {
      const parts = line.split(':');
      // ACTION:UPDATE_RESI:[no_order]:[resi]:[ekspedisi]
      const noOrder   = parts[2];
      const resi      = parts[3];
      const ekspedisi = parts[4] || '';
      if (noOrder && resi) {
        const ok = await updateResiInSheet(noOrder, resi, ekspedisi);
        console.log(`Update resi ${noOrder}: ${ok ? 'OK' : 'not found'}`);
      }
    }

    if (line.startsWith('ACTION:UPDATE_STATUS:')) {
      const parts  = line.split(':');
      const noOrder = parts[2];
      const status  = parts[3];
      if (noOrder && status) {
        const ok = await updateStatusInSheet(noOrder, status);
        console.log(`Update status ${noOrder} → ${status}: ${ok ? 'OK' : 'not found'}`);
      }
    }
  }
}

// ─── MESSAGING HELPERS ────────────────────────────────────
async function sendWA(target, message) {
  await axios.post('https://api.fonnte.com/send', {
    target,
    message,
  }, {
    headers: { Authorization: FONNTE_TOKEN },
  });
}

async function sendTelegram(chatId, message) {
  await axios.post(
    `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`,
    { chat_id: chatId, text: message, parse_mode: 'HTML' }
  );
}

// ─── WEBHOOKS ─────────────────────────────────────────────

// WhatsApp via Fonnte
app.post('/webhook/wa', async (req, res) => {
  res.sendStatus(200);
  const { sender, message } = req.body;
  if (!sender || !message) return;
  console.log(`WA [${sender}]: ${message}`);
  try {
    const reply = await callClaude(`wa_${sender}`, message);
    await sendWA(sender, reply);
  } catch (e) {
    console.error('WA error:', e.message);
    await sendWA(sender, 'Maaf, terjadi error. Coba lagi sebentar.');
  }
});

// Telegram Bot
app.post('/webhook/telegram', async (req, res) => {
  res.sendStatus(200);
  const msg = req.body.message || req.body.edited_message;
  if (!msg || !msg.text) return;
  const chatId = msg.chat.id;
  const text   = msg.text;
  console.log(`TG [${chatId}]: ${text}`);
  try {
    const reply = await callClaude(`tg_${chatId}`, text);
    await sendTelegram(chatId, reply);
  } catch (e) {
    console.error('Telegram error:', e.message);
    await sendTelegram(chatId, 'Maaf, terjadi error. Coba lagi.');
  }
});

// Health check
app.get('/health', (_, res) => res.json({ status: 'ok' }));

// ─── H-1 NOTIFICATION SCHEDULER ──────────────────────────
// Jalan setiap hari pukul 08:00 WIB (UTC+7 = 01:00 UTC)
cron.schedule('0 1 * * *', async () => {
  console.log('Running H-1 notification check...');
  try {
    const data = await getSheetData();
    if (!data.length) return;

    const headers = data[0];
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = tomorrow.toISOString().split('T')[0];

    const noOrderIdx  = colIndex(headers, 'No Order');
    const customerIdx = colIndex(headers, 'Nama Customer');
    const reqDateIdx  = colIndex(headers, 'Request Date');
    const statusIdx   = colIndex(headers, 'Status');
    const resiIdx     = colIndex(headers, 'Driver/Booking/Resi');
    const ekspIdx     = colIndex(headers, 'Ekspedisi');

    const h1Orders = data.slice(1).filter(row => {
      const reqDate = (row[reqDateIdx] || '').trim();
      const status  = (row[statusIdx] || '').trim();
      return reqDate === tomorrowStr && status !== 'Delivered';
    });

    if (!h1Orders.length) {
      console.log('No H-1 pending orders today.');
      return;
    }

    let msg = `🔔 <b>H-1 PENDING REMINDER</b>\n`;
    msg += `Request Date besok: <b>${tomorrowStr}</b>\n\n`;

    h1Orders.forEach(row => {
      const no       = row[noOrderIdx] || '—';
      const customer = row[customerIdx] || '—';
      const eksp     = row[ekspIdx] || '—';
      const resi     = row[resiIdx] || '';
      msg += `📦 <b>${no}</b> — ${customer} (${eksp})\n`;
      msg += resi
        ? `   ✅ Resi: ${resi}\n`
        : `   ❌ Resi belum diinput!\n`;
    });

    msg += `\nTotal: ${h1Orders.length} order`;

    // Kirim ke Telegram
    if (TELEGRAM_CHAT_ID) {
      await sendTelegram(TELEGRAM_CHAT_ID, msg);
      console.log('H-1 notification sent to Telegram.');
    }

    // Kirim ke WA
    if (YOUR_WA_NUMBER) {
      const waMsg = msg.replace(/<[^>]+>/g, ''); // strip HTML tags for WA
      await sendWA(YOUR_WA_NUMBER, waMsg);
      console.log('H-1 notification sent to WhatsApp.');
    }

  } catch (e) {
    console.error('H-1 scheduler error:', e.message);
  }
});

// ─── START ────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`OPS LOG Agent running on port ${PORT}`);
  console.log(`Webhooks:`);
  console.log(`  POST /webhook/wa       → WhatsApp (Fonnte)`);
  console.log(`  POST /webhook/telegram → Telegram Bot`);
});
