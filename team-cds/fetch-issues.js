const fetch = require('node-fetch');

const url = process.env.FETCH_GITLAB_ISSUES_SERVICE;

fetch(url)
  .then(res => res.text())
  .then(result => console.log(result))  // e.g., ✅ Sync successful
  .catch(err => console.error('❌ Error:', err));

