const { google } = require('googleapis');

const sheetData = JSON.parse(process.env.SHEET_DATA);

async function main() {
  const spreadsheetUrl = sheetData.spreadsheetUrl;
  const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

  const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  const skip = ['ToC', 'Roster', 'Issues'];

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets.map(s => s.properties.title);

    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const sheetMeta = metadata.data.sheets.find(s => s.properties.title === name);
      const sheetId = sheetMeta.properties.sheetId;

      const rangeE = `'${name}'!E12:E`;
      const resE = await sheets.spreadsheets.values.get({ spreadsheetId, range: rangeE });
      const rowsE = resE.data.values || [];
      const startRow = 12;

      const mergedRanges = sheetMeta.merges || [];
      const requests = [];

      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;

        // Only consider merged cells in column E (index 4)
        if (startColumnIndex === 4 && endColumnIndex === 5 && startRowIndex >= startRow) {
          let isEmpty = true;

          for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex++) {
            const valueE = (rowsE[rowIndex - startRow] && rowsE[rowIndex - startRow][0]) || '';
            if (valueE.trim()) {
              isEmpty = false;
              break;
            }
          }

          if (isEmpty) {
            requests.push({
              unmergeCells: {
                range: { ...merge, sheetId }
              }
            });
          }
        }
      }

      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
        });
        console.log(`✅ Unmerged empty cells in sheet: ${name}`);
      } else {
        console.log(`ℹ️ No empty merged cells to unmerge in sheet: ${name}`);
      }
    }
  } catch (err) {
    console.error('❌ ERROR:', err);
    process.exit(1);
  }
}

main();
