const fetch = require('node-fetch');

// Fetch the GitLab issues data from your Google Apps Script endpoint
const url = process.env.FETCH_ALL_GITLAB_ISSUES_SERVICE;

fetch(url)
  .then(res => res.text()) // Expecting text response from the Apps Script
  .then(result => console.log(result))  // Log the result (e.g., ✅ Sync successful)
  .catch(err => console.error('❌ Error:', err)); // Handle any errors
