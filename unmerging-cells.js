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
      const rangeF = `'${name}'!F12:F`;
      const resE = await sheets.spreadsheets.values.get({ spreadsheetId, range: rangeE });
      const resF = await sheets.spreadsheets.values.get({ spreadsheetId, range: rangeF });
      const rowsE = resE.data.values || [];
      const rowsF = resF.data.values || [];
      const startRow = 12;

      const mergedRanges = sheetMeta.merges || [];
      const requests = [];

      // Iterate through all merged ranges
      const toUnmerge = [];

      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;

        // Only check merged ranges in columns E (index 4) and F (index 5)
        if (startColumnIndex === 4 && endColumnIndex === 5) {
          const mergeStart = startRowIndex;
          const mergeEnd = endRowIndex;

          let isEmpty = true;

          // Check if all cells in the merged range (both E and F) are empty
          for (let rowIndex = mergeStart; rowIndex < mergeEnd; rowIndex++) {
            const valueE = (rowsE[rowIndex - startRow] && rowsE[rowIndex - startRow][0]) || ''; // Get value from column E
            const valueF = (rowsF[rowIndex - startRow] && rowsF[rowIndex - startRow][0]) || ''; // Get value from column F

            if (valueE.trim() || valueF.trim()) {
              isEmpty = false;
              break;
            }
          }

          // If both columns are empty, mark for unmerging
          if (isEmpty) {
            toUnmerge.push(merge);
          }
        }
      }

      // Unmerge empty merged cells
      for (const merge of toUnmerge) {
        requests.push({
          unmergeCells: { range: { ...merge, sheetId } }
        });

        // Optional: Add borders to unmerged cells
        requests.push({
          updateBorders: {
            range: { ...merge, sheetId },
            top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
            bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
            left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
            right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
          }
        });
      }

      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
        });
      }

      console.log(`✅ Updated: ${name}`);
    }
  } catch (err) {
    console.error('❌ ERROR:', err);
    process.exit(1);
  }
}

main();
