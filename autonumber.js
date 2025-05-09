const { google } = require('googleapis');

const sheetData = JSON.parse(process.env.SHEET_DATA);

async function main() {
  const spreadsheetUrl = sheetData.spreadsheetUrl;
  const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
  const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets.map(s => s.properties.title);
    const skip = ['ToC', 'Roster', 'Issues'];
    const startRow = 11; // 0-indexed row 12

    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const sheetMeta = metadata.data.sheets.find(s => s.properties.title === name);
      const merges = sheetMeta.merges || [];
      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let values = Array(rows.length).fill(['']);
      let requests = [];
      let number = 1;

      const mergedMap = new Map();

      for (const merge of merges) {
        if (
          (merge.startColumnIndex === 5 || merge.startColumnIndex === 4) &&
          merge.endColumnIndex <= 6 &&
          merge.startRowIndex >= startRow
        ) {
          mergedMap.set(merge.startRowIndex, merge.endRowIndex);
        }
      }

      let row = 0;
      while (row < rows.length) {
        const absRow = row + startRow;
        const fCell = (rows[row][1] || '').trim();
        const mergeEnd = mergedMap.get(absRow);

        const isMerged = mergeEnd && mergeEnd > absRow;
        const mergeLength = isMerged ? mergeEnd - absRow : 1;

        if (isMerged && !fCell) {
          // Empty merged range: Add black border
          requests.push({
            updateBorders: {
              range: {
                sheetId: sheetMeta.properties.sheetId,
                startRowIndex: absRow,
                endRowIndex: mergeEnd,
                startColumnIndex: 4,
                endColumnIndex: 6,
              },
              top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } }
            }
          });
        }

        if (fCell) {
          values[row] = [number.toString()];

          if (isMerged) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId: sheetMeta.properties.sheetId,
                  startRowIndex: absRow,
                  endRowIndex: mergeEnd,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL'
              }
            });
          }

          number++;
        }

        row += mergeLength;
      }

      // Unmerge all in E and F columns from row 12 onward
      requests.unshift({
        unmergeCells: {
          range: {
            sheetId: sheetMeta.properties.sheetId,
            startRowIndex: startRow,
            endRowIndex: startRow + rows.length,
            startColumnIndex: 4,
            endColumnIndex: 6
          }
        }
      });

      // Update E column values
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${12 + rows.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values }
      });

      // Apply formatting and merging updates
      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
        });
      }

      console.log(`Updated and formatted sheet: ${name}`);
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1);
  }
}

main();
