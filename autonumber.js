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
      const sheetMetadata = metadata.data.sheets.find(sheet => sheet.properties.title === name);
      const sheetId = sheetMetadata.properties.sheetId;

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cellF = row[1] || ''; // Column F
        const cellE = values[i][0]; // Column E

        // Fetch current merged ranges in E and F
        const eRange = { startRowIndex: 12 + i, endRowIndex: 12 + i + 1 };
        const fRange = { startRowIndex: 12 + i, endRowIndex: 12 + i + 1 };

        // Merging logic for Column E
        if (cellF.trim() && cellE) {
          const mergeCountF = await getMergedCount(spreadsheetId, sheetId, fRange, sheets);
          if (mergeCountF > 1) {
            await mergeCells(spreadsheetId, sheetId, eRange, sheets);
          }
        } else {
          await unmergeCells(spreadsheetId, sheetId, eRange, sheets);
        }

        // Merging logic for Column F
        if (cellF.trim() && lastStepNumber === cellE) {
          const mergeCountE = await getMergedCount(spreadsheetId, sheetId, eRange, sheets);
          if (mergeCountE > 1) {
            await mergeCells(spreadsheetId, sheetId, fRange, sheets);
          }
        } else {
          await unmergeCells(spreadsheetId, sheetId, fRange, sheets);
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
async function getMergedCount(spreadsheetId, sheetId, range, sheets) {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: [range],
    includeGridData: true
  });

  const gridData = res.data.sheets.find(sheet => sheet.properties.sheetId === sheetId).data[0];
  const mergedRanges = gridData.merges || [];
  return mergedRanges.filter(merge => merge.startRowIndex <= range.endRowIndex && merge.endRowIndex >= range.startRowIndex).length;
}

// Helper function to merge cells
async function mergeCells(spreadsheetId, sheetId, range, sheets) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          mergeCells: {
            range: {
              sheetId: sheetId,
              startRowIndex: range.startRowIndex,
              endRowIndex: range.endRowIndex
            },
            mergeType: 'MERGE_ALL'
          }
        }
      ]
    }
  });
}

// Helper function to unmerge cells
async function unmergeCells(spreadsheetId, sheetId, range, sheets) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          unmergeCells: {
            range: {
              sheetId: sheetId,
              startRowIndex: range.startRowIndex,
              endRowIndex: range.endRowIndex
            }
          }
        }
      ]
    }
  });
}

main();
