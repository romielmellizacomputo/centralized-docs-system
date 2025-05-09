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
    const skip = ['ToC', 'Roster', 'Issues']; // Sheets to skip

    const requests = [];

    for (const name of sheetNames) {
      if (skip.includes(name)) continue; // Skip certain sheets

      const range = `'${name}'!E12:F`; // Range to fetch data from column E to F
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let num = 1;

      for (let i = 0; i < rows.length; i++) {
        const eCell = rows[i][0];
        const fCell = rows[i][1];

        // Ensure the cells are not undefined or null before using trim
        const eCellTrimmed = (eCell && eCell.trim()) || '';
        const fCellTrimmed = (fCell && fCell.trim()) || '';

        // Auto number logic: Only increment if F cell has data
        const updatedValue = (fCellTrimmed || '') ? num++ : '';

        // Prepare request to update the E column with auto-numbering or empty
        requests.push({
          updateCells: {
            rows: [{
              values: [{
                userEnteredValue: { stringValue: updatedValue.toString() }
              }]
            }],
            fields: "userEnteredValue",
            start: {
              sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
              rowIndex: 11 + i,
              columnIndex: 4
            }
          }
        });

        // If F column is empty, unmerge E column cells
        if (!fCellTrimmed) {
          requests.push({
            unmergeCells: {
              range: {
                sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
                startRowIndex: 11 + i,
                endRowIndex: 12 + i,
                startColumnIndex: 4,
                endColumnIndex: 5,
              }
            }
          });
        } else if (eCellTrimmed && fCellTrimmed) {
          // If F column has data, ensure E column is merged appropriately
          if (i + 1 < rows.length && rows[i + 1][1] && rows[i + 1][1].trim()) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
                  startRowIndex: 11 + i,
                  endRowIndex: 12 + i + 1,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL' // Merge the E cells when F has data
              }
            });
          }
        }
      }

      // Execute batch update requests if there are any changes
      if (requests.length) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
        });

        console.log(`Updated sheet: ${name}`);
      }
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1); // Exit with an error code if something goes wrong
  }
}

main();
