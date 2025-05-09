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
  // Parse the credentials stored in the GitHub secret as an environment variable
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials, // Use the credentials directly from the secret
    scopes: ['https://www.googleapis.com/auth/spreadsheets'], // Use the appropriate scope for your needs
  });

  return auth;
}

async function getSelectedMilestones(sheets, sheetId) {
  // Fetch the selected milestones data from G-Milestones sheet
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${G_MILESTONES}!A2:A`, // Adjust the range as needed
  });

  return data.values.flat().filter(Boolean); // Returns an array of non-empty values (milestones)
}

async function getAllIssues(sheets) {
  // Fetch all issues data from the ALL ISSUES sheet in the UTILS spreadsheet
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${ALL_ISSUES_SHEET}!A2:N`, // Adjust as per your sheet structure
  });

  return data.values;
}

async function clearSheet(sheets, sheetId) {
  // Clears the data from the G-Issues sheet
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_ISSUES}!A2:Z`, // Adjust the range to clear data as necessary
  });
}

async function insertData(sheets, sheetId, data) {
  // Inserts the processed data into the G-Issues sheet
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_ISSUES}!A2`, // Adjust the range to insert data at the correct location
    valueInputOption: 'RAW',
    requestBody: {
      values: data,
    },
  });
}

async function updateTimestamp(sheets, sheetId) {
  // Updates the timestamp for when the sheet was processed
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD}!A1`, // Adjust to where you want the timestamp
    valueInputOption: 'RAW',
    requestBody: {
      values: [[timestamp]],
    },
  });
}

async function main() {
  // Authenticate using the service account credentials
  const auth = await authenticate();
  const sheets = google.sheets({ version: 'v4', auth });

  // Fetch list of sheet IDs from the UTILS sheet (column B)
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${UTILS_SHEET_NAME}!B2:B`, // Fetch all sheet IDs from column B
  });

  const sheetIds = data.values.flat().filter(Boolean); // Filter out any empty values

  for (const sheetId of sheetIds) {
    try {
      // Fetch milestones and issues data in parallel
      const [milestones, issuesData] = await Promise.all([
        getSelectedMilestones(sheets, sheetId),
        getAllIssues(sheets),
      ]);

      // Filter issues based on matching milestones (Column I in ALL ISSUES)
      const filtered = issuesData.filter(row =>
        milestones.includes(row[8]) // Assuming column I (index 8) holds the milestone data
      );

      // Process data and prepare for insertion into G-Issues
      const processedData = filtered.map(row => {
        // Assuming row[3:5] contains relevant data (adjust as per your sheet)
        return [row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12]];
      });

      // Clear the G-Issues sheet, insert the processed data, and update the timestamp
      await clearSheet(sheets, sheetId);
      await insertData(sheets, sheetId, processedData);
      await updateTimestamp(sheets, sheetId);
      console.log(`✔ Processed sheet: ${sheetId}`);
    } catch (err) {
      console.error(`❌ Failed to process ${sheetId}: ${err.message}`);
    }
  }
}

main();
