fetch('https://script.google.com/macros/s/AKfycbzbNWs4I_NvuCfb_DL18mEGdyK4e7vHTb057UzDVl4zNg25BZmm5ecEMxI6NWb1u537/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
