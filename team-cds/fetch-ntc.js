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

import {
  authenticate,
  getSheetTitles,
  getAllTeamCDSSheetIds,
  getSelectedMilestones,
} from './common.js';

async function getAllIssues(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    range: ALL_NTC,
  });

  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ALL_NTC}`);
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
