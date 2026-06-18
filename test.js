const axios = require('axios');

axios.post('https://sentralcargo.co.id', 
  new URLSearchParams({ 'awb[]': '993597928' }), 
  { headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'User-Agent': 'Mozilla/5.0' } }
)
.then(res => {
  const html = JSON.stringify(res.data).toLowerCase();
  console.log('\n=== HASIL TRACKING SENTRAL CARGO ===');
  if (html.includes('tidak ditemukan')) {
    console.log('Status Terbaca: RESI TIDAK VALID');
  } else if (html.includes('delivered') || html.includes('diterima oleh')) {
    console.log('Status Terbaca: DELIVERED');
  } else {
    console.log('Status Terbaca: ON PROCESS');
  }
  console.log('====================================\n');
})
.catch(err => console.log('Error Koneksi:', err.message));
