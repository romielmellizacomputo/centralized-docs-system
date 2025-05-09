import { google } from 'googleapis';

// Your Google Sheets constants
const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const UTILS_SHEET_NAME = 'UTILS';
const G_ISSUES = 'G-Issues';
const G_MILESTONES = 'G-Milestones';
const DASHBOARD = 'Dashboard';
const ALL_ISSUES_SHEET = 'ALL ISSUES';

// This function loads the credentials from the environment variable
async function authenticate() {
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return auth;
}

async function getSelectedMilestones(sheets, sheetId) {
  // Fetch selected milestones from G4:G of G-Milestones sheet
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${G_MILESTONES}!G4:G`, // Range for selected milestones
  });

  return data.values.flat().filter(Boolean); // Flatten and filter out any empty values
}

async function getAllIssues(sheets) {
  try {
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: UTILS_SHEET_ID,
      range: `${ALL_ISSUES_SHEET}!A2:N`, // Adjust as per your sheet
    });

    if (!data.values || data.values.length === 0) {
      throw new Error(`No data found in range ${ALL_ISSUES_SHEET}!A2:N`);
    }

    return data.values;
  } catch (err) {
    console.error(`❌ Error fetching ALL ISSUES: ${err.message}`);
    throw err; // Rethrow the error to halt further execution
  }
}

async function clearSheet(sheets, sheetId) {
  // Clear the existing sheet data
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_ISSUES}!A2:Z`, // Adjust as per your sheet structure
  });
}

async function insertData(sheets, sheetId, data) {
  // Insert the processed data
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_ISSUES}!A2`, // Adjust as per your sheet structure
    valueInputOption: 'RAW',
    requestBody: {
      values: data,
    },
  });
}

async function updateTimestamp(sheets, sheetId) {
  // Update the timestamp for when the sheet was processed
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD}!A1`, // Adjust as per your sheet structure
    valueInputOption: 'RAW',
    requestBody: {
      values: [[timestamp]],
    },
  });
}

async function main() {
  try {
    // Authenticate using the service account credentials
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    // Check available sheets and ranges
    const metadata = await sheets.spreadsheets.get({
      spreadsheetId: UTILS_SHEET_ID,
    });
    console.log(metadata); // Log metadata to check sheet names

    // Fetch list of spreadsheet IDs from UTILS sheet
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: UTILS_SHEET_ID,
      range: `${UTILS_SHEET_NAME}!B2:B`,
    });

    const sheetIds = data.values.flat().filter(Boolean);

    for (const sheetId of sheetIds) {
      try {
        // Fetch selected milestones and issues data
        const [milestones, issuesData] = await Promise.all([
          getSelectedMilestones(sheets, sheetId),
          getAllIssues(sheets),
        ]);

        console.log('Milestones:', milestones); // Debugging log
        console.log('Issues Data:', issuesData); // Debugging log

        // Filter issues based on the selected milestones
        const filtered = issuesData.filter(row =>
          milestones.includes(row[8]) // Assuming column I (index 8) has the milestone name
        );

        console.log('Filtered Issues:', filtered); // Debugging log

        // Process data for insertion into the sheet
        const processedData = filtered.map(row => [
          row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12],
        ]);

        console.log('Processed Data:', processedData); // Debugging log

        // Clear the sheet, insert the processed data, and update the timestamp
        await clearSheet(sheets, sheetId);
        await insertData(sheets, sheetId, processedData);
        await updateTimestamp(sheets, sheetId);
        console.log(`✔ Processed sheet: ${sheetId}`);
      } catch (err) {
        console.error(`❌ Failed to process ${sheetId}: ${err.message}`);
      }
    }
  } catch (err) {
    console.error(`❌ Main processing failed: ${err.message}`);
  }
}

main();
