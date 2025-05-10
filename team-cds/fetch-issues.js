fetch('https://script.google.com/macros/s/AKfycbwMbVNwBJq94o3DSB2YWduvJgx44lciPARed1OJkHkiPddNPBFogNgL8O8FxH0xG1wf/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
