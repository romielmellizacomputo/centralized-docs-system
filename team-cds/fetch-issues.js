fetch('https://script.google.com/macros/s/AKfycbzrUpNHtnXg5yZEWOysg-AH5O6RVYXKrZUhzECxxEu3-vBomlF_vak84gAHlFtbw_-1/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
