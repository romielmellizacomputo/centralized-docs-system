fetch('https://script.google.com/macros/s/AKfycbzlfKEQqosn_MrRhsPHimk1mYumbJBDq7SVZOo9owRdhfJQW9xrXVDcbmFFJh4_vc2W/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
