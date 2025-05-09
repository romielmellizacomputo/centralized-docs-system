const { google } = require('googleapis');
const fetch = require('node-fetch'); // Use global fetch if on Node 18+

const skipSheets = ['ToC', 'Issues', 'Roster'];
const ISSUES_SHEET = 'Issues';
const DROPDOWN_RANGE = 'K3:K';

async function refreshDropdown() {
  const sheetData = JSON.parse(process.env.SHEET_DATA);
  const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheetsAPI = google.sheets({ version: 'v4', auth });

  const spreadsheetIdMatch = sheetData.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!spreadsheetIdMatch) {
    console.error('Invalid spreadsheet URL');
    process.exit(1);
  }

  const spreadsheetId = spreadsheetIdMatch[1];
  const baseSheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=`;

  try {
    const metadata = await sheetsAPI.spreadsheets.get({ spreadsheetId });
    const sheets = metadata.data.sheets;

    const sheetNames = sheets
      .map(s => s.properties.title)
      .filter(name => !skipSheets.includes(name));

    const sheetNameToGid = {};
    sheets.forEach(s => {
      const name = s.properties.title;
      const gid = s.properties.sheetId;
      if (!skipSheets.includes(name)) {
        sheetNameToGid[name] = gid;
      }
    });

    const issuesSheet = sheets.find(s => s.properties.title === ISSUES_SHEET);
    if (!issuesSheet) throw new Error(`Sheet "${ISSUES_SHEET}" not found`);

    // 1. Apply dropdown validation
    const dropdownRule = {
      requests: [{
        setDataValidation: {
          range: {
            sheetId: issuesSheet.properties.sheetId,
            startRowIndex: 2,
            startColumnIndex: 10,
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

    await sheetsAPI.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: dropdownRule
    });

    console.log('Dropdown updated successfully.');

    // 2. Fetch existing K3:K values
    const getRes = await sheetsAPI.spreadsheets.values.get({
      spreadsheetId,
      range: `${ISSUES_SHEET}!K3:K`
    });

    const values = getRes.data.values || [];

    // 3. Replace with =HYPERLINK formulas if matching a valid sheet name
    const updatedValues = values.map(row => {
      const val = row[0]?.trim();
      if (sheetNameToGid[val]) {
        const link = `${baseSheetUrl}${sheetNameToGid[val]}`;
        return [`=HYPERLINK("${link}", "${val}")`];
      } else {
        return [val || ''];
      }
    });

    await sheetsAPI.spreadsheets.values.update({
      spreadsheetId,
      range: `${ISSUES_SHEET}!K3`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: updatedValues }
    });

    console.log('Hyperlinks added successfully.');

  } catch (err) {
    console.error('Error refreshing dropdown:', err.message);
    process.exit(1);
  }
}

refreshDropdown();
