import { google } from 'googleapis';
import {
  UTILS_SHEET_ID,
  G_MILESTONES,
  G_MR_SHEET,
  DASHBOARD_SHEET,
  CENTRAL_ISSUE_SHEET_ID,
  ALL_MR,
  generateTimestampString
} from '../constants.js';

import {
  authenticate,
  getSheetTitles,
  getAllTeamCDSSheetIds,
  getSelectedMilestones,
} from './common.js';

async function getAllMR(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: ALL_MR,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_MR}`);
  }

  return data.values;
}

async function clearGMR(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4:S`,
  });
}

function padRowToU(row) {
  const fullLength = 17;
  return [...row, ...Array(fullLength - row.length).fill('')];
}

async function insertDataToGMR(sheets, sheetId, data) {
  const paddedData = data.map(row => padRowToU(row.slice(0, 17)));

  console.log(`📤 Inserting ${paddedData.length} rows to ${G_MR_SHEET}!C4`);
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${G_MR_SHEET}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: paddedData },
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

    const sheetIds = await getAllTeamCDSSheetIds(sheets, UTILS_SHEET_ID);
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

        const [milestones, issuesData] = await Promise.all([
          getSelectedMilestones(sheets, sheetId, G_MILESTONES),
          getAllMR(sheets),
        ]);

        const filtered = issuesData.filter(row => milestones.includes(row[7])); // Column J
        const processedData = filtered.map(row => row.slice(0, 17));

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
