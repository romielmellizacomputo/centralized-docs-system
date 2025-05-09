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

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];
      const startRow = 12;

      const mergedRanges = sheetMeta.merges || [];
      const requests = [];

      const mergedMap = new Map();

      // Identify merged ranges in column F (index 5)
      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;
        if (startRowIndex >= 11 && startColumnIndex === 5 && endColumnIndex === 6) {
          mergedMap.set(startRowIndex, endRowIndex);

          // Merge column E (index 4) to match
          requests.push({
            mergeCells: {
              range: {
                sheetId,
                startRowIndex,
                endRowIndex,
                startColumnIndex: 4,
                endColumnIndex: 5,
              },
              mergeType: 'MERGE_ALL',
            }
          });

          // Center alignment for merged E cells
          requests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex,
                endRowIndex,
                startColumnIndex: 4,
                endColumnIndex: 5,
              },
              cell: {
                userEnteredFormat: {
                  horizontalAlignment: 'CENTER',
                  verticalAlignment: 'MIDDLE',
                }
              },
              fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)',
            }
          });
        }
      }

      // Prepare values for E12:E
      const values = Array(rows.length).fill(['']);
      let row = 0;
      let number = 1;

      while (row < rows.length) {
        const absRow = row + startRow;
        const fCell = (rows[row] && rows[row][1] || '').trim();
        const isMerged = mergedMap.has(absRow);
        const mergeEnd = isMerged ? mergedMap.get(absRow) : absRow + 1;
        const mergeLength = mergeEnd - absRow;

        if (fCell) {
          // Write number only in the top cell of the merged block
          values[row] = [number.toString()];
          number++;
        }

        row += mergeLength;
      }

      // Write numbers to E12:E
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${startRow + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });

      // Apply formatting and merging
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
