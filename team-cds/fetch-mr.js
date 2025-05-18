import { google } from 'googleapis';
import {
  UTILS_SHEET_ID,
  G_MILESTONES,
  G_ISSUES_SHEET,
  DASHBOARD_SHEET,
  CENTRAL_ISSUE_SHEET_ID,
  ALL_ISSUES,
  generateTimestampString
} from '../constants.js';

async function authenticate() {
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getSheetTitles(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  const titles = res.data.sheets.map(sheet => sheet.properties.title);
  console.log(`📄 Sheets in ${spreadsheetId}:`, titles);
  return titles;
}

async function getAllTeamCDSSheetIds(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: 'UTILS!B2:B',
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function getSelectedMilestones(sheets, sheetId) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${G_MILESTONES}!G4:G`,
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function getAllIssues(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: ALL_MR,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_MR}`);
  }

  return data.values;
}

async function clearGIssues(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4:N`,
  });
}

async function insertDataToGIssues(sheets, sheetId, data) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: data },
  });
}

async function updateTimestamp(sheets, sheetId) {
  const formatted = generateTimestampString();

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD_SHEET}!W6`,
    valueInputOption: 'RAW',
    requestBody: { values: [[formatted]] },
  });
}


async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    await getSheetTitles(sheets, UTILS_SHEET_ID);

    const sheetIds = await getAllTeamCDSSheetIds(sheets);
    if (!sheetIds.length) {
      console.error('❌ No Team CDS sheet IDs found in UTILS!B2:B');
      return;
    }

    for (const sheetId of sheetIds) {
      try {
        console.log(`🔄 Processing: ${sheetId}`);

        const sheetTitles = await getSheetTitles(sheets, sheetId);

        if (!sheetTitles.includes(G_MILESTONES)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${G_MILESTONES}' sheet`);
          continue;
        }

        if (!sheetTitles.includes(G_MR_SHEET)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${G_ISSUES_SHEET}' sheet`);
          continue;
        }

        const [milestones, issuesData] = await Promise.all([ 
          getSelectedMilestones(sheets, sheetId),
          getAllIssues(sheets),
        ]);

        const filtered = issuesData.filter(row => milestones.includes(row[7])); // Column I (index 6)
        const processedData = filtered.map(row => row.slice(0, 13));

        await clearGIssues(sheets, sheetId);
        await insertDataToGIssues(sheets, sheetId, processedData);
        await updateTimestamp(sheets, sheetId);

        console.log(`✅ Finished: ${sheetId}`);
      } catch (err) {
        console.error(`❌ Error processing ${sheetId}: ${err.message}`);
      }
    }
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
