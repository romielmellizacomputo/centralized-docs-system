import { google } from 'googleapis';

// Google Sheets constants
const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const G_MILESTONES = 'G-Milestones';
const G_MR_SHEET = 'G-MR';
const DASHBOARD_SHEET = 'Dashboard';

const CENTRAL_MR_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';
const ALL_MRS_RANGE = 'ALL MRs!C4:O'; // ✅ Columns C to O

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

async function getAllMRs(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_MR_SHEET_ID,
    range: ALL_MRS_RANGE,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_MRS_RANGE}`);
  }

  return data.values;
}

async function clearGMR(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4:O`,
  });
}

async function insertDataToGMR(sheets, sheetId, data) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: data },
  });
}

async function updateTimestamp(sheets, sheetId) {
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD_SHEET}!X6`,
    valueInputOption: 'RAW',
    requestBody: { values: [[timestamp]] },
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
          console.warn(`⚠️ Skipping ${sheetId} — missing '${G_MR_SHEET}' sheet`);
          continue;
        }

        const [milestones, mrData] = await Promise.all([
          getSelectedMilestones(sheets, sheetId),
          getAllMRs(sheets),
        ]);

        const filtered = mrData.filter(row => milestones.includes(row[7])); // ✅ Validate using Column J
        const processedData = filtered.map(row => row.slice(0, 13)); // ✅ C to O = index 0 to 12

        await clearGMR(sheets, sheetId);
        await insertDataToGMR(sheets, sheetId, processedData);
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
