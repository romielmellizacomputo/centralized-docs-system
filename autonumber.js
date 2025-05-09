const { google } = require('googleapis');

// Retrieve sheet data from environment variable
const sheetData = JSON.parse(process.env.SHEET_DATA);

async function main() {
  const spreadsheetUrl = sheetData.spreadsheetUrl;
  const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets.map((s) => s.properties.title);
    const skip = ['ToC', 'Roster', 'Issues'];

    const requests = [];

    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let num = 1;
      const values = rows.map(([_, f]) => [(f || '').trim() ? num++ : '']);

      // Request for updating values in the E column
      requests.push({
        updateCells: {
          range: {
            sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
            startRowIndex: 11,
            endRowIndex: 11 + values.length,
            startColumnIndex: 4,
            endColumnIndex: 5,
          },
          rows: values.map((value) => ({
            values: [
              {
                userEnteredValue: {
                  stringValue: value[0].toString(),
                },
              },
            ],
          })),
          fields: 'userEnteredValue',
        },
      });

      // Logic to unmerge cells in the E and F columns if necessary
      for (let i = 0; i < rows.length; i++) {
        const eCell = rows[i][0];
        const fCell = rows[i][1];

        if (!fCell || !fCell.trim()) {
          // If F column is empty, unmerge E column cells
          requests.push({
            mergeCells: {
              range: {
                sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
                startRowIndex: 11 + i,
                endRowIndex: 12 + i,
                startColumnIndex: 4,
                endColumnIndex: 5,
              },
              mergeType: 'MERGE_UNDEFINED', // Unmerge cells
            },
          });
        } else if (eCell.trim() && fCell.trim()) {
          // If F column has data, ensure E column is merged appropriately
          if (i + 1 < rows.length && rows[i + 1][1].trim()) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId: metadata.data.sheets.find((s) => s.properties.title === name).properties.sheetId,
                  startRowIndex: 11 + i,
                  endRowIndex: 12 + i + 1,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL', // Merge the E cells when F has data
              },
            });
          }
        }
      }

      console.log(`Updated sheet: ${name}`);
    }

    if (requests.length) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1);
  }
}

main();
