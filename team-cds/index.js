import { google } from 'googleapis';

const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const ALL_ISSUES_SHEET = 'ALL ISSUES';
const MILESTONES_SHEET = 'G-Milestones';
const ISSUES_SHEET = 'G-Issues';
const DASHBOARD_SHEET = 'Dashboard';

async function authenticate() {
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getTeamSheetIds(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: 'UTILS!B2:B',
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function getSelectedMilestones(sheets, teamSheetId) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: teamSheetId,
    range: `${MILESTONES_SHEET}!G4:G`,
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function getAllIssues(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: `${ALL_ISSUES_SHEET}!C4:N`,
  });
  return data.values || [];
}

async function filterIssuesByMilestones(issues, milestones) {
  return issues.filter(row => milestones.includes(row[6])); // Index 6 = column I
}

async function clearIssuesSheet(sheets, teamSheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: teamSheetId,
    range: `${ISSUES_SHEET}!C4:N`,
  });
}

async function insertIssues(sheets, teamSheetId, data) {
  if (!data.length) return;
  await sheets.spreadsheets.values.update({
    spreadsheetId: teamSheetId,
    range: `${ISSUES_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: data },
  });
}

async function updateTimestamp(sheets, teamSheetId) {
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: teamSheetId,
    range: `${DASHBOARD_SHEET}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [[timestamp]] },
  });
}

async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    const teamSheetIds = await getTeamSheetIds(sheets);
    const allIssues = await getAllIssues(sheets);

    for (const teamSheetId of teamSheetIds) {
      try {
        const milestones = await getSelectedMilestones(sheets, teamSheetId);
        const filtered = await filterIssuesByMilestones(allIssues, milestones);

        // Prepare only columns C to N (indexes 0 to 11)
        const processed = filtered.map(row => row.slice(0, 12));

        await clearIssuesSheet(sheets, teamSheetId);
        await insertIssues(sheets, teamSheetId, processed);
        await updateTimestamp(sheets, teamSheetId);

        console.log(`✔ Processed: ${teamSheetId}`);
      } catch (err) {
        console.error(`❌ Error processing ${teamSheetId}: ${err.message}`);
      }
    }
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
