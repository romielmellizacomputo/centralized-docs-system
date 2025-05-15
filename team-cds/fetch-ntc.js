import { google } from 'googleapis';

const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const G_MILESTONES = 'G-Milestones';
const NTC_SHEET = 'NTC'; 
const DASHBOARD_SHEET = 'Dashboard';

const CENTRAL_ISSUE_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY'; 
const ALL_ISSUES_RANGE = 'ALL ISSUES!C4:N'; 

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

  return data.values;
}

async function clearNTCSheet(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4:N`,
  });
}

async function insertDataToNTCSheet(sheets, sheetId, data) {
  if (data.length === 0) {
    console.log("No data to insert.");
    return; 
  }
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: data },
  });
}

async function updateTimestamp(sheets, sheetId) {
  const now = new Date();
  const timeZoneEAT = 'Africa/Nairobi'; // East Africa Time
  const timeZonePHT = 'Asia/Manila'; // Philippine Time

  const options = {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  };

  const formattedUTC = now.toLocaleString('en-US', options);
  const formattedEAT = new Intl.DateTimeFormat('en-US', { 
    ...options, 
    timeZone: timeZoneEAT 
  }).format(now);
  const formattedPHT = new Intl.DateTimeFormat('en-US', { 
    ...options, 
    timeZone: timeZonePHT 
  }).format(now);

  const formatted = `Sync on ${formattedUTC} (UTC) / ${formattedEAT} (EAT) / ${formattedPHT} (PHT)`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${DASHBOARD_SHEET}!X6`,
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

        if (!sheetTitles.includes(NTC_SHEET)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${NTC_SHEET}' sheet`);
          continue;
        }

        const [milestones, issuesData] = await Promise.all([ 
          getSelectedMilestones(sheets, sheetId),
          getAllIssues(sheets),
        ]);

        const filtered = issuesData.filter(row => {
          const milestoneMatches = milestones.includes(row[6]);

          const labelsRaw = row[5] || '';  
          const labels = labelsRaw.split(',').map(label => label.trim().toLowerCase());

          console.log(`Raw labels for row: ${labelsRaw}`);
          console.log(`Processed labels for row: ${labels}`);

          const labelsMatch = labels.some(label => 
            ["needs test case", "needs test scenario", "test case needs update"].includes(label)
          );

          return milestoneMatches && labelsMatch;
        });

        if (filtered.length > 0) {
          const processedData = filtered.map(row => row.slice(0, 12)); 

          await clearNTCSheet(sheets, sheetId);
          await insertDataToNTCSheet(sheets, sheetId, processedData);
          await updateTimestamp(sheets, sheetId);

          console.log(`✅ Finished: ${sheetId}`);
        } else {
          console.log(`⚠️ No matching data for ${sheetId}`);
        }
      } catch (err) {
        console.error(`❌ Error processing ${sheetId}: ${err.message}`);
      }
    }
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
