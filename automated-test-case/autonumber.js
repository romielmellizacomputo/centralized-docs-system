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
    const allData = {};

    // Fetch all data from each sheet
    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const range = `'${name}'!A:Z`; // Adjust range as needed
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      allData[name] = res.data.values || [];
    }

    const requests = [];

    // Process data in memory
    for (const [name, rows] of Object.entries(allData)) {
      const merges = metadata.data.sheets.find(s => s.properties.title === name).merges || [];
      const startRow = 12; // Assuming numbering starts from row 12
      let number = 1;

      for (let row = 0; row < rows.length; row++) {
        const absRow = row + startRow;

        // Check for merge in column F
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

        const fValue = (rows[row] && rows[row][1])?.trim();

        if (fValue) {
          // Fill number
          rows[row][0] = number.toString(); // Assuming column E is the first column in the fetched range

          // Prepare merge/unmerge requests if needed
          if (isMergedInF) {
            requests.push({
              mergeCells: {
                range: {
                  sheetId: metadata.data.sheets.find(s => s.properties.title === name).properties.sheetId,
                  startRowIndex: mergeStart - 1,
                  endRowIndex: mergeEnd - 1,
                  startColumnIndex: 4,
                  endColumnIndex: 5,
                },
                mergeType: 'MERGE_ALL',
              },
            });
          }

          number++;
        }
      }
    }

    // Update all values in one go
    for (const [name, rows] of Object.entries(allData)) {
      const range = `'${name}'!E12:E${12 + rows.length - 1}`; // Adjust range as needed
      await updateValuesWithRetry(sheets, spreadsheetId, range, rows.map(row => [row[0]]));
    }

    // Apply merge/unmerge requests in batches
    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });
    }

    console.log('✅ All sheets updated successfully.');
  } catch (err) {
    console.error('❌ ERROR:', err);
    process.exit(1);
  }
}

async function updateValuesWithRetry(sheets, spreadsheetId, range, values) {
  let attempts = 0;

  while (true) { // Infinite loop for retries
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });
      return; // Success, exit the function
    } catch (error) {
      if (error.response && error.response.status === 429) {
        // Quota exceeded error, apply exponential backoff
        attempts++;
        const waitTime = Math.pow(2, attempts) * 1000; // Exponential backoff
        console.log(`Quota exceeded. Retrying in ${waitTime / 1000} seconds...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      } else {
        throw error; // Rethrow other errors
      }
    }
  }
}

main();
