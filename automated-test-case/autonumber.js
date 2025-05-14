import { google } from 'googleapis';

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
      const merges = sheetMeta.merges || [];

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];
      const startRow = 12;

      const requests = [];
      const values = Array(rows.length).fill(['']);
      let number = 1;
      let row = 0;

      while (row < rows.length) {
        const absRow = row + startRow;

        const fMerge = merges.find(m =>
          m.startRowIndex === absRow - 1 &&
          m.startColumnIndex === 5 &&
          m.endColumnIndex === 6
        );

        let mergeStart = absRow;
        let mergeEnd = absRow + 1;

        if (fMerge) {
          mergeStart = fMerge.startRowIndex + 1;
          mergeEnd = fMerge.endRowIndex + 1;
        }

        const isMergedInF = mergeEnd > mergeStart;
        const mergeLength = mergeEnd - mergeStart;

        const fValue = (rows[row] && rows[row][1])?.trim();
        const eValue = (rows[row] && rows[row][0])?.trim();

        const eMerge = merges.find(m =>
          m.startRowIndex === mergeStart - 1 &&
          m.endRowIndex === mergeEnd - 1 &&
          m.startColumnIndex === 4 &&
          m.endColumnIndex === 5
        );

        if (fValue) {
          // Fill number
          values[row] = [number.toString()];

          if (isMergedInF && !eMerge) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId,
                  startRowIndex: mergeStart - 1,
                  endRowIndex: mergeEnd - 1,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL',
              },
            });
          }

          if (!isMergedInF && eMerge) {
            requests.push({
              unmergeCells: {
                range: {
                  sheetId,
                  startRowIndex: eMerge.startRowIndex,
                  endRowIndex: eMerge.endRowIndex,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                }
              }
            });
          }

          number++;
        }

        row += mergeLength;
      }

      await updateValuesWithRetry(sheets, spreadsheetId, `'${name}'!E12:E${startRow + values.length - 1}`, values);

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

async function updateValuesWithRetry(sheets, spreadsheetId, range, values) {
  let attempts = 0;

  while (true) { 
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });
      return; 
    } catch (error) {
      if (error.response && error.response.status === 429) {
        // Quota exceeded error, apply exponential backoff
        attempts++;
        const waitTime = Math.pow(2, attempts) * 1000; // Exponential backoff
        console.log(`Quota exceeded. Retrying in ${waitTime / 1000} seconds...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      } else {
        throw error; 
      }
    }
  }
}

main();
