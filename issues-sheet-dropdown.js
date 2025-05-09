// issues-sheet-dropdown.js

const { google } = require('googleapis');
const fetch = require('node-fetch'); // Use global fetch if on Node 18+

// Constants
const skipSheets = ['ToC', 'Issues', 'Roster'];
const ISSUES_SHEET = 'Issues';
const DROPDOWN_RANGE = 'K3:K';
const webAppUrl = 'https://script.google.com/macros/s/AKfycbzR3hWvfItvEOKjadlrVRx5vNTz4QH04WZbz2ufL8fAdbiZVsJbkzueKfmMCfGsAO62/exec';

async function refreshDropdown() {
  // Load environment variables
  const sheetData = JSON.parse(process.env.SHEET_DATA);
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheetsAPI = google.sheets({ version: 'v4', auth });

  // Extract Spreadsheet ID from URL
  const spreadsheetIdMatch = sheetData.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!spreadsheetIdMatch) {
    console.error('Invalid spreadsheet URL');
    process.exit(1);
  }

  const spreadsheetId = spreadsheetIdMatch[1];

  try {
    // Get metadata including sheet names and IDs
    const metadata = await sheetsAPI.spreadsheets.get({ spreadsheetId });
    const sheets = metadata.data.sheets;

    const sheetNames = sheets
      .map(s => s.properties.title)
      .filter(name => !skipSheets.includes(name));

    const issuesSheet = sheets.find(s => s.properties.title === ISSUES_SHEET);
    if (!issuesSheet) {
      throw new Error(`Sheet "${ISSUES_SHEET}" not found`);
    }

    const dropdownRule = {
      requests: [{
        setDataValidation: {
          range: {
            sheetId: issuesSheet.properties.sheetId,
            startRowIndex: 2, // Row 3
            startColumnIndex: 10, // Column K
            endColumnIndex: 11
          },
          rule: {
            condition: {
              type: 'ONE_OF_LIST',
              values: sheetNames.map(name => ({ userEnteredValue: name }))
            },
            strict: true,
            showCustomUi: true
          }
        }
      }]
    };

    // Apply validation rule
    await sheetsAPI.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: dropdownRule
    });

    console.log('Dropdown updated successfully.');

    // Notify web app
    const postRes = await fetch(webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheetUrl: sheetData.spreadsheetUrl })
    });

    if (!postRes.ok) {
      throw new Error(`Web app POST failed: ${postRes.statusText}`);
    }

    console.log('Posted to web app successfully.');

  } catch (err) {
    console.error('Error refreshing dropdown:', err.message);
    process.exit(1);
  }
}

refreshDropdown();
