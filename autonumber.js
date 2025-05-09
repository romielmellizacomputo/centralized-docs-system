const { google } = require('googleapis');

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

      const sheetMeta = metadata.data.sheets.find(s => s.properties.title === name);
      const merges = sheetMeta.merges || [];
      const startRow = 11; // zero-based index for row 12
      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let values = Array(rows.length).fill(['']);
      let requests = [];
      let number = 1;

      const mergedMap = new Map();

      for (const merge of merges) {
        if (
          merge.startColumnIndex === 5 && // Column F (zero-based)
          merge.endColumnIndex === 6 &&
          merge.startRowIndex >= startRow
        ) {
          mergedMap.set(merge.startRowIndex, merge.endRowIndex);
        }
      }

      let row = 0;
      while (row < rows.length) {
        const cellF = (rows[row][1] || '').trim();
        const absRow = row + startRow;
        const mergeEnd = mergedMap.get(absRow);

        if (cellF) {
          if (mergeEnd) {
            const mergeLength = mergeEnd - absRow + 1; // Include end row
            values[row] = [number.toString()];
            requests.push({
              mergeCells: {
                range: {
                  sheetId: sheetMeta.properties.sheetId,
                  startRowIndex: absRow,
                  endRowIndex: mergeEnd + 1,
                  startColumnIndex: 4,
                  endColumnIndex: 5
                },
                mergeType: 'MERGE_ALL'
              }
            });
            row += mergeLength;
          } else {
            values[row] = [number.toString()];
            row += 1;
          }
          number++;
        } else {
          if (mergeEnd) {
            // Unmerge if F is empty but was previously merged
            requests.push({
              unmergeCells: {
                range: {
                  sheetId: sheetMeta.properties.sheetId,
                  startRowIndex: absRow,
                  endRowIndex: mergeEnd + 1,
                  startColumnIndex: 5,
                  endColumnIndex: 6
                }
              }
            });
            // Apply borders to the unmerged cells
            for (let i = absRow; i <= mergeEnd; i++) {
              requests.push({
                updateBorders: {
                  range: {
                    sheetId: sheetMeta.properties.sheetId,
                    startRowIndex: i,
                    endRowIndex: i + 1,
                    startColumnIndex: 5,
                    endColumnIndex: 6
                  },
                  top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
                  bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
                  left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
                  right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
                }
              });
            }
            row = mergeEnd - absRow + 1; // Move to the end of the merged cells
          } else {
            row += 1;
          }
        }
      }

      // Clear old merges first
      requests.unshift({
        unmergeCells: {
          range: {
            sheetId: sheetMeta.properties.sheetId,
            startRowIndex: startRow,
            endRowIndex: startRow + rows.length,
            startColumnIndex: 4,
            endColumnIndex: 5
          }
        }
      });

      // Update values in column E
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${12 + rows.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values }
      });

      // Apply merge requests
      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
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
