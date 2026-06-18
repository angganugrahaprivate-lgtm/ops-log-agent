const axios = require('axios');

// Mengambil API Key aman dari environment variable Railway Anda
const BITESHIP_KEY = process.env.BITESHIP_API_KEY;

async function testTrackBiteship(noResi) {
  if (!BITESHIP_KEY) {
    console.error('❌ ERROR: Variable BITESHIP_API_KEY tidak ditemukan di Railway!');
    return;
  }

  try {
    console.log(`[API] Menghubungi gateway Biteship untuk resi Sentral Cargo: ${noResi}...`);
    
    // Request ke API Biteship dengan nomor resi dan kurir Sentral Cargo
    const response = await axios.post('https://biteship.com', {
      waybill_id: noResi,
      courier_code: 'sentral'
    }, {
      headers: {
        'Authorization': `Bearer ${BITESHIP_KEY}`,
        'Content-Type': 'application/json'
      }
    });

    const data = response.data;
    
    console.log('\n=== HASIL INTEGRASI API RESMI ===');
    console.log('Status Paket :', data.status?.toUpperCase());
    console.log('Kurir        :', data.courier?.company);
    console.log('Detail Terakhir:', data.history?.[0]?.note || 'Tidak ada catatan');
    console.log('=================================\n');

  } catch (error) {
    console.error('\n❌ GAGAL MENGAMBIL DATA API');
    console.error('Pesan Error:', error.response?.data?.message || error.message);
    console.log('===========================\n');
  }
}

// Jalankan tes menggunakan nomor resi simulasi / asli Anda
testTrackBiteship('000123456789');
