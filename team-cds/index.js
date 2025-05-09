import { google } from 'googleapis';
import { authenticate } from './sheets.js';

const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const UTILS_SHEET_NAME = 'UTILS';
const G_ISSUES = 'G-Issues';
const G_MILESTONES = 'G-Milestones';
const DASHBOARD = 'Dashboard';
const ALL_ISSUES_SHEET = 'ALL ISSUES';

async function main() {
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
      const [milestones, issuesData] = await Promise.all([
        getSelectedMilestones(sheets, sheetId),
        getAllIssues(sheets)
      ]);

      const filtered = issuesData.filter(row =>
        milestones.includes(row[6]) // assuming column I (index 6)
      );

      const processedData = filtered.map(row => {
        const hyperlink = row[4];
        const linkText = hyperlink?.text || '';
        const linkUrl = hyperlink?.hyperlink || '';
        return [linkText, linkUrl, ...row.slice(0, 4), ...row.slice(5)];
      });

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
