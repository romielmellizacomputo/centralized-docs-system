const { google } = require('googleapis');
const fetch = require('node-fetch'); // Ensure 'node-fetch' is installed if using Node.js <18

const skipSheets = ['ToC', 'Issues', 'Roster'];
const ISSUES_SHEET = 'Issues';
const DROPDOWN_RANGE = 'K3:K';
const webAppUrl = 'https://script.google.com/macros/s/AKfycbzR3hWvfItvEOKjadlrVRx5vNTz4QH04WZbz2ufL8fAdbiZVsJbkzueKfmMCfGsAO62/exec';

async function refreshDropdown() {
  const sheetData = JSON.parse(process.env.SHEET_DATA);
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheetsAPI = google.sheets({ version: 'v4', auth });

  const spreadsheetId = sheetData.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

  try {
    const metadata = await sheetsAPI.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets
      .map(s => s.properties.title)
      .filter(name => !skipSheets.includes(name));

    const rule = {
      requests: [{
        setDataValidation: {
          range: {
            sheetId: metadata.data.sheets.find(s => s.properties.title === ISSUES_SHEET).properties.sheetId,
            startRowIndex: 2, // Row 3
            startColumnIndex: 10, // Column K
            endColumnIndex: 11
          },
          rule: {
            condition: {
              type: 'ONE_OF_LIST',
              values: sheetNames.map(name => ({ userEnteredValue: name }))
            },
            showCustomUi: true,
            strict: true
          }
        }
      }]
    };

    await sheetsAPI.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: rule
    });

    console.log('Dropdown updated successfully.');

    // Post to the web app
    await fetch(webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetUrl: sheetData.spreadsheetUrl })
    });

    console.log('Posted to web app successfully.');
  } catch (err) {
    console.error('Error refreshing dropdown:', err.message);
    process.exit(1);
  }
}

refreshDropdown();
