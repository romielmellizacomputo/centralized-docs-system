fetch('https://script.google.com/macros/s/AKfycbwXt9aUsRGDMxxxmFPW_2UngHFJKbutzBWZaNiXn4TQD-aUmKATpGe2DcPuh3A7xKV9/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
