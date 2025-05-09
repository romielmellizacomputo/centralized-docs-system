const { google } = require('googleapis');

// Retrieve sheet data from environment variable
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

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let num = 1;
      let lastStepNumber = 0;
      const values = rows.map(([_, f], idx) => {
        const cellF = f || '';
        const currentStep = cellF.trim() ? num++ : '';
        if (currentStep) {
          lastStepNumber = currentStep;
        }
        return [currentStep];
      });

      // Update step numbers in column E
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${12 + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values }
      });

      // Check and merge/unmerge cells
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cellF = row[1] || ''; // Column F
        const cellE = values[i][0]; // Column E

        // Fetch current merged ranges in E and F
        const eRange = `'${name}'!E${12 + i}:E${12 + i}`;
        const fRange = `'${name}'!F${12 + i}:F${12 + i}`;

        // Merging logic for Column E
        if (cellF.trim() && cellE) {
          const mergeCountF = await getMergedCount(spreadsheetId, name, eRange, sheets);
          if (mergeCountF > 1) {
            await mergeCells(spreadsheetId, name, eRange, sheets);
          }
        } else {
          await unmergeCells(spreadsheetId, name, eRange, sheets);
        }

        // Merging logic for Column F
        if (cellF.trim() && lastStepNumber === cellE) {
          const mergeCountE = await getMergedCount(spreadsheetId, name, fRange, sheets);
          if (mergeCountE > 1) {
            await mergeCells(spreadsheetId, name, fRange, sheets);
          }
        } else {
          await unmergeCells(spreadsheetId, name, fRange, sheets);
        }
      }

      console.log(`Updated sheet: ${name}`);
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1);
  }
}

// Helper function to check for merged ranges in a given range
async function getMergedCount(spreadsheetId, sheetName, range, sheets) {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [range], // Ensure ranges is an array of ranges
    includeGridData: true
  });

  const gridData = res.data.sheets[0].data[0];
  const mergedRanges = gridData.merges || [];
  return mergedRanges.filter(merge => merge.startRowIndex <= range[1] && merge.endRowIndex >= range[0]).length;
}

// Helper function to merge cells
async function mergeCells(spreadsheetId, sheetName, range, sheets) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          mergeCells: {
            range: {
              sheetId: sheetName,
              startRowIndex: range[0],
              endRowIndex: range[1]
            },
            mergeType: 'MERGE_ALL'
          }
        }
      ]
    }
  });
}

// Helper function to unmerge cells
async function unmergeCells(spreadsheetId, sheetName, range, sheets) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          unmergeCells: {
            range: {
              sheetId: sheetName,
              startRowIndex: range[0],
              endRowIndex: range[1]
            }
          }
        }
      ]
    }
  });
}

main();
