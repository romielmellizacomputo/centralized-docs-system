const { google } = require('googleapis');

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

      // We will iterate over all merged ranges and unmerge if they are empty in either column E or F
      const toUnmerge = [];

      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;

        // Check merged ranges that affect E12:E (Column 4) or F12:F (Column 5)
        if ((startColumnIndex === 4 || startColumnIndex === 5) && endColumnIndex === startColumnIndex + 1) {
          const mergeStart = startRowIndex;
          const mergeEnd = endRowIndex;
          let isEmpty = true;

          // Check the merged cells in column E and column F for emptiness
          for (let r = mergeStart; r < mergeEnd; r++) {
            const rowE = rowsE[r - 12] ? rowsE[r - 12][0] : '';
            const rowF = rowsF[r - 12] ? rowsF[r - 12][0] : '';

            // If any cell is not empty, set isEmpty to false and break
            if (rowE.trim() || rowF.trim()) {
              isEmpty = false;
              break;
            }
          }

          if (isEmpty) {
            toUnmerge.push(merge);
          }
        }
      }

      // Unmerge all empty merged ranges in E and F
      for (const merge of toUnmerge) {
        requests.push({
          unmergeCells: { range: { ...merge, sheetId } }
        });

        // Add black border to newly unmerged empty merged ranges
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
