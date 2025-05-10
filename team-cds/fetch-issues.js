fetch('https://script.google.com/macros/s/AKfycbxkm2tm2z0O3z73i1pjEtnmzh25gsttHLbowUIUnKpEoQjccj98RcUJ_pigD7aNMOqE/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
