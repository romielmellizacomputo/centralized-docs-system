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

      // Track merged ranges in E-F to handle unmerge
      const toUnmerge = [];
      const mergedMap = new Map();

      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;
        if (startRowIndex >= 11 && startColumnIndex === 4 && endColumnIndex >= 6) {
          const mergeStart = startRowIndex;
          const mergeEnd = endRowIndex;
          let isEmpty = true;

          for (let r = mergeStart - 12; r < mergeEnd - 12; r++) {
            if (rows[r] && rows[r][1] && rows[r][1].trim()) {
              isEmpty = false;
              break;
            }
          }

          if (isEmpty) {
            toUnmerge.push(merge);
          } else {
            mergedMap.set(mergeStart, mergeEnd);
          }
        }
      }

      // Unmerge only empty merged cells
      for (const merge of toUnmerge) {
        requests.push({
          unmergeCells: { range: { ...merge, sheetId } }
        });

        // Add black border to newly unmerged empty merged ranges
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


      // Prepare numbering and merges
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
          values[row] = [number.toString()];
          if (isMerged) {
            // Already merged, skip re-merging
          } else if (mergeLength > 1) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId,
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

      // Write numbers to E12:E
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
