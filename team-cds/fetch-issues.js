fetch('https://script.google.com/macros/s/AKfycbwZZI1jG69t5-S5GnAJfoU4yIOQzoADgyLiKUCMfBBIDYeJ-la3JWeBySNVkAs1db_A/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
