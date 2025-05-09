const { google } = require('googleapis');

// Retrieve sheet data from environment variable
const sheetData = JSON.parse(process.env.SHEET_DATA);

async function main() {
  const spreadsheetUrl = sheetData.spreadsheetUrl;
  const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets.map(s => s.properties.title);
    const skip = ['ToC', 'Roster', 'Issues'];

    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let num = 1;
      const values = rows.map(([e, f], index) => {
        const eValue = (e || '').trim();
        const fValue = (f || '').trim();

        // If F is empty, remove the number in E and unmerge cells
        if (!fValue) {
          return ['', '']; // Clear both E and F
        }

        // Handle merging E cells and clearing F if needed
        if (eValue && !fValue) {
          return ['', '']; // Remove number in E if F is empty
        }

        return [eValue ? num++ : '', fValue];
      });

      // Update the values in the sheet (columns E and F)
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${12 + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: values.map(row => [row[0]]) }
      });

      // Update F column values
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!F12:F${12 + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: values.map(row => [row[1]]) }
      });

      // Unmerge cells if F is empty
      const unmergeRequests = [];
      for (let i = 0; i < rows.length; i++) {
        const [e, f] = rows[i];
        if (!f) {
          unmergeRequests.push({
            unmerge: {
              range: {
                sheetId: metadata.data.sheets.find(s => s.properties.title === name).properties.sheetId,
                startRowIndex: 12 + i,
                endRowIndex: 13 + i,
                startColumnIndex: 4,
                endColumnIndex: 5
              }
            }
          });
        }
      }

      // Handle merging E and F cells if necessary
      const mergeRequests = [];
      for (let i = 0; i < rows.length - 1; i++) {
        const [e1, f1] = rows[i];
        const [e2, f2] = rows[i + 1];
        if (f1 && f2 && f1 === f2 && e1 && e2) {
          mergeRequests.push({
            mergeCells: {
              range: {
                sheetId: metadata.data.sheets.find(s => s.properties.title === name).properties.sheetId,
                startRowIndex: 12 + i,
                endRowIndex: 13 + i + 1,
                startColumnIndex: 4,
                endColumnIndex: 5
              },
              mergeType: 'MERGE_COLUMNS'
            }
          });
        }
      }

      // Execute the unmerge and merge requests
      if (unmergeRequests.length > 0 || mergeRequests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [...unmergeRequests, ...mergeRequests]
          }
        });
      }

      console.log(`Updated sheet: ${name}`);
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1);
  }
}

main();
