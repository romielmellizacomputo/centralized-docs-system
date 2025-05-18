import { google } from 'googleapis';
import {
  UTILS_SHEET_ID,
  G_MILESTONES,
  NTC_SHEET,
  DASHBOARD_SHEET,
  CENTRAL_ISSUE_SHEET_ID,
  ALL_NTC,
  generateTimestampString
} from '../constants.js';

import {
  authenticate,
  getSheetTitles,
  getAllTeamCDSSheetIds,
  getSelectedMilestones,
} from './common.js';

async function getAllNTC(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: ALL_NTC,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_NTC}`);
  }

  return data.values;
}

async function clearNTC(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4:U`,
  });
}

function padRowToU(row) {
  const fullLength = 14;
  return [...row, ...Array(fullLength - row.length).fill('')];
}

async function insertDataToNTC(sheets, sheetId, data) {
  const paddedData = data.map(row => padRowToU(row.slice(0, 14)));

  console.log(`📤 Inserting ${paddedData.length} rows to ${NTC_SHEET}!C4`);
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${NTC_SHEET}!C4`,
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

        if (!sheetTitles.includes(NTC_SHEET)) {
          console.warn(`⚠️ Skipping ${sheetId} — missing '${NTC_SHEET}' sheet`);
          continue;
        }

        const [milestones, ntcData] = await Promise.all([
          getSelectedMilestones(sheets, sheetId, G_MILESTONES),
          getAllNTC(sheets),
        ]);

        const filtered = ntcData.filter(row => {
          const milestone = row[8]; // Column I
          const labelsRaw = row[7] || ''; // Column H
          const labels = labelsRaw.split(',').map(label => label.trim());

          const matchesMilestone = milestones.includes(milestone);
          const requiredLabels = [
            'Needs Test Case',
            'Needs Test Scenario',
            'Test Case Needs Update'
          ];
          const hasRelevantLabel = labels.some(label =>
            requiredLabels.includes(label)
          );

          return matchesMilestone && hasRelevantLabel;
        });

        const processedData = filtered.map(row => row.slice(0, 21));

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
