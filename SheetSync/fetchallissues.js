fetch('https://script.google.com/macros/s/your-script-id/exec', {
  method: 'POST'
})
.then(res => res.text())
.then(console.log)
.catch(console.error);
