fetch('https://script.google.com/macros/s/AKfycbydPMFK7InLVFVNYMDuYstBBx7unqnEFogU0whS44DUwD6gdkehko9618kjfAIh5o32/exec')
  .then(res => res.text())
  .then(result => console.log(result))  // You'll see either '✅ Sync successful' or error message
  .catch(err => console.error('❌ Error:', err));
