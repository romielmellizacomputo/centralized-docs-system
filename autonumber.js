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

      // Track F column merged ranges
      const fMergedMap = new Map();
      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;
        if (startColumnIndex === 5 && endColumnIndex === 6 && startRowIndex >= 11) {
          fMergedMap.set(startRowIndex, endRowIndex);
        }
      }

      // Track E column merged ranges for correction
      const eMergedRanges = mergedRanges.filter(m =>
        m.startColumnIndex === 4 && m.endColumnIndex === 5 && m.startRowIndex >= 11
      );

      const values = Array(rows.length).fill(['']);
      let number = 1;
      let row = 0;

      while (row < rows.length) {
        const absRow = row + startRow;
        const fCell = (rows[row] && rows[row][1] || '').trim();
        const fMergeEnd = fMergedMap.get(absRow) || absRow + 1;
        const fMergeLen = fMergeEnd - absRow;

        // Check if E is wrongly merged but F isn't
        const eMerge = eMergedRanges.find(m => absRow >= m.startRowIndex && absRow < m.endRowIndex);
        const eMergeStart = eMerge?.startRowIndex;
        const eMergeEnd = eMerge?.endRowIndex;

        if (fCell) {
          values[row] = [number.toString()];

          if (fMergeLen > 1) {
            // F is merged: ensure E is merged similarly
            requests.push({
              mergeCells: {
                range: {
                  sheetId,
                  startRowIndex: absRow,
                  endRowIndex: fMergeEnd,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL'
              }
            });
          }

          number++;
        }

        if (!fMergedMap.has(absRow) && eMergeStart !== undefined) {
          // F is NOT merged but E IS => unmerge
          requests.push({
            unmergeCells: {
              range: {
                sheetId,
                startRowIndex: eMergeStart,
                endRowIndex: eMergeEnd,
                startColumnIndex: 4,
                endColumnIndex: 5,
              }
            }
          });

          requests.push({
            updateBorders: {
              range: {
                sheetId,
                startRowIndex: eMergeStart,
                endRowIndex: eMergeEnd,
                startColumnIndex: 4,
                endColumnIndex: 5,
              },
              top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
            }
          });
        }

        row += fMergeLen;
      }

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${startRow + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });

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
