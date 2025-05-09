import { google } from 'googleapis';
import { authenticate } from './fetch-issues.js';

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

async function main() {
  // Authenticate using the service account credentials
  const auth = await authenticate();
  const sheets = google.sheets({ version: 'v4', auth });

  // Fetch list of spreadsheet IDs from UTILS sheet
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${UTILS_SHEET_NAME}!B2:B`,
  });

  const sheetIds = data.values.flat().filter(Boolean);

  for (const sheetId of sheetIds) {
    try {
      // Fetch milestones and issues data in parallel
      const [milestones, issuesData] = await Promise.all([
        getSelectedMilestones(sheets, sheetId),
        getAllIssues(sheets),
      ]);

      const filtered = issuesData.filter(row =>
        milestones.includes(row[6]) // Assuming column I (index 6)
      );

      const processedData = filtered.map(row => {
        const hyperlink = row[4];
        const linkText = hyperlink?.text || '';
        const linkUrl = hyperlink?.hyperlink || '';
        return [linkText, linkUrl, ...row.slice(0, 4), ...row.slice(5)];
      });

      // Clear the sheet, insert the processed data, and update the timestamp
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
