fetch('https://script.google.com/macros/s/AKfycbxt4g4mEksmnkKbuRL9bYrAKFo4CHfiavSxpUm5QtwngLWIW4pD48IDvVaEIg7NVjRerA/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
