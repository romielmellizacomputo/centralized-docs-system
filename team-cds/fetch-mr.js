fetch('https://script.google.com/macros/s/AKfycbxt4g4mEksmnkKbuRL9bYrAKFo4CHfiavSxpUm5QtwngLWIW4pD48IDvVaEIg7NVjRerA/exec')
  .then(res => res.text())
  .then(text => console.log(text))
  .catch(err => console.error('❌ Error calling GAS:', err));
