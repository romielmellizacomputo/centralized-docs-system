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

      const range = `'${name}'!F12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];
      const startRow = 12;

      const mergedRanges = sheetMeta.merges || [];
      const requests = [];

      const numberData = [];
      let row = 0;
      let number = 1;

      while (row < rows.length) {
        const absRow = row + startRow;
        const fCell = (rows[row] && rows[row][0] || '').trim();

        // Check if F cell is part of a merged range
        const merge = mergedRanges.find(m =>
          m.startColumnIndex === 5 && m.endColumnIndex === 6 &&
          absRow >= m.startRowIndex && absRow < m.endRowIndex
        );

        const mergeStart = merge ? merge.startRowIndex : absRow;
        const mergeEnd = merge ? merge.endRowIndex : absRow + 1;
        const mergeLength = mergeEnd - mergeStart;

        if (fCell) {
          // Add number at top of E cell
          numberData.push({
            rowIndex: mergeStart,
            value: number.toString(),
            mergeStart,
            mergeEnd
          });

          // Merge E column to match F
          requests.push({
            mergeCells: {
              range: {
                sheetId,
                startRowIndex: mergeStart,
                endRowIndex: mergeEnd,
                startColumnIndex: 4, // column E
                endColumnIndex: 5,
              },
              mergeType: 'MERGE_ALL',
            }
          });

          // Center alignment
          requests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: mergeStart,
                endRowIndex: mergeEnd,
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

          number++;
        }

        row += mergeLength;
      }

      // Prepare values array for E12:E
      const values = Array(rows.length).fill(['']);
      for (const item of numberData) {
        const rowIdx = item.rowIndex - startRow;
        if (rowIdx >= 0 && rowIdx < values.length) {
          values[rowIdx] = [item.value];
        }
      }

      // Write to E12:E
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${startRow + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });

      // Apply merging and formatting
      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests },
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
