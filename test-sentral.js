const axios = require('axios');

// Fungsi pembantu parseDate
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

// Fungsi Scraper Utama Sentral Cargo
async function trackSentralCargo(awbNumber) {
  if (!awbNumber) return null;
  const cleanAwb = awbNumber.toString().trim();
  
  try {
    console.log(`[Scraper] Menghubungi web Sentral Cargo untuk resi: ${cleanAwb}...`);
    
    const response = await axios.post(
      'https://sentralcargo.co.id', 
      new URLSearchParams({ 'awb[]': cleanAwb }),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        },
        timeout: 10000
      }
    );

    const html = response.data;
    
    if (html.includes('Data tidak ditemukan') || html.includes('tidak terdaftar')) {
      return { 
        success: false,
        status: 'RESI_TIDAK_VALID', 
        detail: 'Nomor resi tidak ditemukan atau belum terdaftar di sistem Sentral Cargo.' 
      };
    }

    let statusText = 'ON PROCESS';
    
    if (html.toLowerCase().includes('delivered') || html.toLowerCase().includes('diterima oleh')) {
      statusText = 'DELIVERED';
    } else if (html.toLowerCase().includes('pod') || html.toLowerCase().includes('dalam pengantaran')) {
      statusText = 'WITH COURIER';
    } else if (html.toLowerCase().includes('received at')) {
      statusText = 'RECEIVED AT HUB';
    }

    let remarkText = 'Diproses otomatis oleh sistem';
    const receiverMatch = html.match(/Diterima Oleh\s*:\s*([^<]+)/i) || html.match(/Receiver\s*:\s*([^<]+)/i);
    if (receiverMatch && receiverMatch) {
      remarkText = `Diterima oleh: ${receiverMatch[1].trim()}`;
    }

    let tanggalTibaText = '';
    if (statusText === 'DELIVERED') {
      const dateMatch = html.match(/(\d{4}-\d{2}-\d{2})|(\d{2}\/\d{2}\/\d{4})/);
      if (dateMatch) {
        tanggalTibaText = parseDate(dateMatch[0]);
      } else {
        tanggalTibaText = new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' });
      }
    }

    return {
      success: true,
      status: statusText,
      remark: remarkText,
      tglTiba: tanggalTibaText
    };

  } catch (error) {
    return {
      success: false,
      status: 'ERROR_CONNECTION',
      detail: error.message
    };
  }
}

// Jalankan Tes Pengujian
async function jalankanUjiCoba() {
  console.log('==================================================');
  console.log('       UJI COBA INTEGRASI TRACKING SENTRAL CARGO  ');
  console.log('==================================================\n');
  
  // !!! GANTI teks di bawah ini dengan NOMOR RESI ASLI milik Anda !!!
  const nomorResiTarget = '993597928'; 
  
  const hasilEkstraksi = await trackSentralCargo(nomorResiTarget);
  
  console.log('Hasil keluaran data objek untuk Google Sheets Anda:');
  console.log(JSON.stringify(hasilEkstraksi, null, 2));
  console.log('\n==================================================');
}

jalankanUjiCoba();
