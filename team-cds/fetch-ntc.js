import { google } from 'googleapis';

// Google Sheets constants
const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const G_MILESTONES = 'G-Milestones';
const NTC_SHEET = 'NTC'; // Changed target sheet to NTC
const DASHBOARD_SHEET = 'Dashboard';

const CENTRAL_ISSUE_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';  // External sheet ID
const ALL_ISSUES_RANGE = 'ALL ISSUES!C4:N'; // Range to pull issues from
const H_ISSUE_STATUS_RANGE = 'ALL ISSUES!H4:H'; // Range for the status check (Needs Test Case, etc.)

async function authenticate() {
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getSheetTitles(sheets, spreadsheetId) {
  try {
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    const titles = res.data.sheets ? res.data.sheets.map(sheet => sheet.properties.title) : [];
    console.log(`📄 Sheets in ${spreadsheetId}:`, titles);
    return titles;
  } catch (error) {
    console.error(`Error fetching sheet titles for ${spreadsheetId}:`, error);
    return []; // Return empty array in case of error
  }
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
    range: ALL_ISSUES_RANGE,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_ISSUES_RANGE}`);
  }

  return data.values;
}

async function getIssueStatuses(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: H_ISSUE_STATUS_RANGE,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${H_ISSUE_STATUS_RANGE}`);
  }

  return data.values;
}

async function clearNTC(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4:N`,
  });
}

async function insertDataToNTC(sheets, sheetId, data) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: data },
  });
}

async function updateTimestamp(sheets, sheetId) {
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD_SHEET}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: [[timestamp]] },
  });
}

async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    // Confirm correct sheet titles
    const sheetTitles = await getSheetTitles(sheets, UTILS_SHEET_ID);
    if (!Array.isArray(sheetTitles) || sheetTitles.length === 0) {
      console.error('❌ No sheet titles found in UTILS sheet.');
      return;
    }

    // Get list of Google Sheet IDs from UTILS!B2:B
    const sheetIds = await getAllTeamCDSSheetIds(sheets);
    if (!sheetIds.length) {
      console.error('❌ No Team CDS sheet IDs found in UTILS!B2:B');
      return;
    }

    for (const sheetId of sheetIds) {
      try {
        console.log(`🔄 Processing: ${sheetId}`);

        const sheetTitles = await getSheetTitles(sheets, sheetId);

        // Ensure sheetTitles is an array and contains the necessary sheets
        if (!sheetTitles || !Array.isArray(sheetTitles) || sheetTitles.length === 0) {
          console.warn(`⚠️ Skipping ${sheetId} — No sheet titles found.`);
          continue;
        }

        if (!sheetTitles.includes(G_MILESTONES)) {
          console.warn(`⚠️ Skipping ${sheetId} — Missing '${G_MILESTONES}' sheet.`);
          continue;
        }

        if (!sheetTitles.includes(NTC_SHEET)) {
          console.warn(`⚠️ Skipping ${sheetId} — Missing '${NTC_SHEET}' sheet.`);
          continue;
        }

        const [milestones, issuesData, issueStatuses] = await Promise.all([ 
          getSelectedMilestones(sheets, sheetId),
          getAllIssues(sheets),
          getIssueStatuses(sheets),
        ]);

        // Ensure the data arrays are defined
        if (!Array.isArray(milestones) || !Array.isArray(issuesData) || !Array.isArray(issueStatuses)) {
          console.error(`❌ Invalid data in one of the arrays for sheet ${sheetId}`);
          continue;
        }

        // Filter issues that match both the selected milestones and the issue statuses
        const filtered = issuesData.filter((row, index) => {
          const matchesMilestone = milestones.includes(row[6]); // Column I (index 6)
          const status = issueStatuses[index] ? issueStatuses[index][0] : '';
          const matchesStatus = ['Needs Test Case', 'Needs Test Scenario', 'Test Case Needs Update'].some(statusText =>
            status.includes(statusText)
          );
          return matchesMilestone && matchesStatus;
        });

        const processedData = filtered.map(row => row.slice(0, 11)); // C to N → index 0 to 10

        await clearNTC(sheets, sheetId);
        await insertDataToNTC(sheets, sheetId, processedData);
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
