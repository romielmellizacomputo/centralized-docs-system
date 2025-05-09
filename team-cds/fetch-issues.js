import { google } from 'googleapis';

// Google Sheets constants
const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const G_MILESTONES = 'G-Milestones';
const G_ISSUES_SHEET = 'G-Issues';
const DASHBOARD_SHEET = 'Dashboard';

const CENTRAL_ISSUE_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';  // External sheet ID
const ALL_ISSUES_RANGE = 'ALL ISSUES!C4:N'; // Range to pull issues from
const ALL_ISSUES_HYPERLINKS_RANGE = 'ALL ISSUES!E4:E'; // Range to pull hyperlinks from

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
    range: ALL_ISSUES_RANGE,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_ISSUES_RANGE}`);
  }

  const hyperlinksData = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: ALL_ISSUES_HYPERLINKS_RANGE,
  });

  // Combine issues data with hyperlinks
  return data.values.map((row, index) => {
    const hyperlink = hyperlinksData.data.values[index] ? hyperlinksData.data.values[index][0] : '';
    return [...row, hyperlink]; // Append hyperlink to the row
  });
}

async function clearGIssues(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_ISSUES_SHEET}!C4:N`,
  });
}

async function insertDataToGIssues(sheets, sheetId, data) {
  const formattedData = data.map(row => {
    const hyperlink = row.pop(); // Remove hyperlink from the end
    if (hyperlink) {
      row.push({
        userEnteredValue: { stringValue: row[0] }, // Assuming the first column is the display text
        userEnteredFormat: { textFormat: { foregroundColor: { red: 0, green: 0, blue: 1 }, underline: true } },
        hyperlink: hyperlink // Add the hyperlink
      });
    }
    return row;
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_ISSUES_SHEET}!C4`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: formattedData },
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
    await getSheetTitles(sheets, UTILS_SHEET_ID);

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

        if (!sheetTitles.includes(G_MILESTONES)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${G_MILESTONES}' sheet`);
          continue;
        }

        if (!sheetTitles.includes(G_ISSUES_SHEET)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${G_ISSUES_SHEET}' sheet`);
          continue;
        }

        const [milestones, issuesData] = await Promise.all([ 
          getSelectedMilestones(sheets, sheetId),
          getAllIssues(sheets),
        ]);

        const filtered = issuesData.filter(row => milestones.includes(row[6])); // Column I (index 6)
        const processedData = filtered.map(row => row.slice(0, 11)); // C to N → index 0 to 10

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
