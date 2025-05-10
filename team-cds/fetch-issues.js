fetch('https://script.google.com/macros/s/AKfycbzAmAMI-NGzhTQd4polExAvF-FkcjucVOP2tWF4Vs-RbvxYbcF1NOCQAwFqyHP0tHZp/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
