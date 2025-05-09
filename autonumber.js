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
      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];
      const startRow = 12;

      const mergedRanges = sheetMeta.merges || [];
      const requests = [];

      const mergedMap = new Map();
      const toUnmerge = [];

      // 1. Identify merged ranges in F column and determine which to unmerge or preserve
      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;
        if (startColumnIndex === 5 && endColumnIndex === 6 && startRowIndex >= 11) {
          const isEmpty = [...Array(endRowIndex - startRowIndex)].every((_, i) => {
            const row = rows[startRowIndex - 12 + i];
            return !(row && row[1] && row[1].trim());
          });

          if (isEmpty) {
            toUnmerge.push({ ...merge, sheetId });
          } else {
            mergedMap.set(startRowIndex, endRowIndex);
          }
        }
      }

      // 2. Unmerge the empty ranges in column F and apply borders to E
      for (const merge of toUnmerge) {
        requests.push({ unmergeCells: { range: merge } });
        for (let r = merge.startRowIndex; r < merge.endRowIndex; r++) {
          requests.push({
            updateBorders: {
              range: {
                sheetId,
                startRowIndex: r,
                endRowIndex: r + 1,
                startColumnIndex: 4,
                endColumnIndex: 5
              },
              top:    { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              left:   { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              right:  { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } }
            }
          });
        }
      }

      // 3. Write numbering to column E and prepare new merge requests if F is merged
      const values = [];
      let row = 0;
      let number = 1;

      while (row < rows.length) {
        const absRow = startRow + row;
        const fCell = (rows[row] && rows[row][1] || '').trim();
        const isMerged = mergedMap.has(absRow);
        const mergeEnd = isMerged ? mergedMap.get(absRow) : absRow + 1;
        const mergeLength = mergeEnd - absRow;

        if (fCell) {
          // Write number in top of merged block
          values[absRow - startRow] = [number.toString()];

          // Merge E-cell block to match F-cell merge if needed
          if (mergeLength > 1) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId,
                  startRowIndex: absRow,
                  endRowIndex: mergeEnd,
                  startColumnIndex: 4,
                  endColumnIndex: 5
                },
                mergeType: 'MERGE_ALL'
              }
            });
          }

          number++;
        }

        row += mergeLength;
      }

      // 4. Fill in empty values in between (with empty strings)
      for (let i = 0; i < rows.length; i++) {
        if (!values[i]) values[i] = [''];
      }

      // 5. Update sheet values
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${startRow + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });

      // 6. Apply all requests (unmerge, border, merge)
      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests }
        });
      }

      console.log(`✅ Updated: ${name}`);
    }
  } catch (err) {
    console.error('❌ ERROR:', err.message || err);
    process.exit(1);
  }
}

main();
