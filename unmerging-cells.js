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

      const toUnmergeE = [];
      const toUnmergeF = [];

      // Track merged ranges in E12:E and F12:F
      for (const merge of mergedRanges) {
        const { startRowIndex, endRowIndex, startColumnIndex, endColumnIndex } = merge;

        if (startRowIndex >= 11 && startColumnIndex === 4 && endColumnIndex >= 5) {
          const mergeStart = startRowIndex;
          const mergeEnd = endRowIndex;
          let isEmpty = true;

          for (let r = mergeStart - 12; r < mergeEnd - 12; r++) {
            if (rowsE[r] && rowsE[r][0] && rowsE[r][0].trim()) {
              isEmpty = false;
              break;
            }
          }

          if (isEmpty) {
            toUnmergeE.push(merge);
          }
        }

        if (startRowIndex >= 11 && startColumnIndex === 5 && endColumnIndex >= 6) {
          const mergeStart = startRowIndex;
          const mergeEnd = endRowIndex;
          let isEmpty = true;

          for (let r = mergeStart - 12; r < mergeEnd - 12; r++) {
            if (rowsF[r] && rowsF[r][0] && rowsF[r][0].trim()) {
              isEmpty = false;
              break;
            }
          }

          if (isEmpty) {
            toUnmergeF.push(merge);
          }
        }
      }

      // Unmerge empty merged cells in E12:E
      for (const merge of toUnmergeE) {
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

      // Unmerge empty merged cells in F12:F
      for (const merge of toUnmergeF) {
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
