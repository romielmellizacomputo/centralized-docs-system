fetch('https://script.google.com/macros/s/AKfycbwu_puWRN-drXODXYhkQf0hE0d4NHuA3Y7eFA4E-mLASu8bHN1jIGLgCntOrp5KC6XN/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
