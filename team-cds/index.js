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
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${G_MILESTONES}!A2:A`,
  });

  const milestones = data.values.flat().filter(Boolean);
  console.log("Milestones:", milestones); // Debugging log
  return milestones;
}

async function getAllIssues(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${ALL_ISSUES_SHEET}!A2:N`,
  });

  return data.values;
}

async function clearSheet(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_ISSUES}!A2:Z`,
  });
}

async function insertData(sheets, sheetId, data) {
  try {
    console.log("Inserting Data:", data); // Debugging log
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${G_ISSUES}!A2`,
      valueInputOption: 'RAW',
      requestBody: {
        values: data,
      },
    });
  } catch (err) {
    console.error("Error inserting data:", err); // Error handling
  }
}

async function updateTimestamp(sheets, sheetId) {
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD}!A1`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [[timestamp]],
    },
  });
}

async function main() {
  const auth = await authenticate();
  const sheets = google.sheets({ version: 'v4', auth });

  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${UTILS_SHEET_NAME}!B2:B`,
  });

  const sheetIds = data.values.flat().filter(Boolean);

  for (const sheetId of sheetIds) {
    try {
      const [milestones, issuesData] = await Promise.all([
        getSelectedMilestones(sheets, sheetId),
        getAllIssues(sheets),
      ]);

      const filtered = issuesData.filter(row =>
        milestones.includes(row[8]) // Ensure this is the correct column for milestones
      );

      console.log("Filtered Issues:", filtered); // Debugging log

      const processedData = filtered.map(row => [
        row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12],
      ]);

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
