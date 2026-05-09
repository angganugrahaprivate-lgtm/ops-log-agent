# OPS LOG Agent — Panduan Setup

## File yang Kamu Dapat
- `server.js` — backend utama agent
- `.env.example` → duplikat jadi `.env`, isi dengan credentials kamu
- `package.json` — daftar dependensi

---

## Langkah 1 — Install & Jalankan

```bash
npm install
cp .env.example .env
# edit .env dengan credentials kamu
node server.js
```

---

## Langkah 2 — Setup Google Sheets API

1. Buka **console.cloud.google.com**
2. Buat project baru (atau pakai yang ada)
3. Enable **Google Sheets API**
4. Buat **Service Account**:
   - IAM & Admin → Service Accounts → Create
   - Download JSON credentials
5. Copy isi JSON tersebut ke `GOOGLE_CREDENTIALS` di `.env` (satu baris)
6. **Bagikan Google Sheet kamu** ke email service account (editor access)

---

## Langkah 3 — Setup Telegram Bot

1. Chat **@BotFather** di Telegram
2. Kirim `/newbot`, ikuti instruksi
3. Copy token ke `TELEGRAM_BOT_TOKEN` di `.env`
4. Daftarkan webhook ke Telegram:

```bash
curl -X POST "https://api.telegram.org/bot[TOKEN]/setWebhook" \
     -d "url=https://[domain-kamu]/webhook/telegram"
```

5. Cari **Chat ID** kamu:
   - Kirim pesan ke bot kamu
   - Buka: `https://api.telegram.org/bot[TOKEN]/getUpdates`
   - Copy nilai `chat.id`
   - Masukkan ke `TELEGRAM_NOTIF_CHAT_ID`

---

## Langkah 4 — Setup Fonnte (WhatsApp)

1. Login ke **fonnte.com**
2. Copy token ke `FONNTE_TOKEN` di `.env`
3. Di dashboard Fonnte, set **Webhook URL**:
   ```
   https://[domain-kamu]/webhook/wa
   ```

---

## Langkah 5 — Deploy (Pilih Salah Satu)

### Opsi A: Railway (Mudah, Gratis)
1. Push kode ke GitHub
2. Buka **railway.app** → New Project → Deploy from GitHub
3. Tambahkan environment variables dari `.env`
4. Railway otomatis berikan domain publik

### Opsi B: VPS (DigitalOcean, Contabo, dll)
```bash
# Install Node.js
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt-get install -y nodejs

# Clone/upload kode
git clone [repo-kamu]
cd ops-agent
npm install
cp .env.example .env
nano .env  # isi credentials

# Jalankan dengan PM2 (agar tetap jalan)
npm install -g pm2
pm2 start server.js --name "ops-agent"
pm2 startup
pm2 save
```

### Opsi C: Ngrok (Testing lokal)
```bash
ngrok http 3000
# Gunakan URL ngrok sebagai domain untuk webhook
```

---

## Cara Pakai

### Via WhatsApp
Kirim pesan ke nomor Fonnte kamu:
- *"Order mana yang aging?"*
- *"Input resi JNE123456 untuk ORD-PLM-001"*
- *"Tampilkan H-1 pending hari ini"*
- *"Update status ORD-PLM-003 jadi Delivered"*

### Via Telegram
Chat langsung ke bot Telegram kamu — sama seperti WA.

### Notifikasi Otomatis H-1
Setiap hari jam **08:00 WIB**, bot otomatis kirim ringkasan H-1 Pending ke WA dan Telegram kamu.

---

## Tentang Poe (Opsional)

Poe bisa digunakan sebagai **interface chat tambahan** (web & mobile):
1. Buka **poe.com** → Create Bot
2. Pilih model: Claude Sonnet
3. Paste system prompt dari file `system_prompt.txt`
4. Share bot ke tim

> **Catatan:** Poe tidak bisa konek ke Google Sheets atau kirim notifikasi WA.
> Gunakan untuk akses chat santai, bukan untuk update data.

---

## Troubleshooting

| Masalah | Solusi |
|---|---|
| Bot tidak balas WA | Cek Fonnte webhook URL sudah benar |
| Bot tidak balas Telegram | Cek webhook sudah didaftarkan ke Telegram |
| Error Google Sheets | Pastikan service account sudah di-share ke sheet |
| Notifikasi H-1 tidak jalan | Cron berjalan UTC, pastikan waktu server sesuai |
