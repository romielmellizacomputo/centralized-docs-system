import { google } from 'googleapis';

// Constants
const UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const CENTRAL_ISSUE_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';

const G_MILESTONES = 'G-Milestones';
const G_ISSUES_SHEET = 'G-Issues';
const DASHBOARD_SHEET = 'Dashboard';
const ALL_ISSUES_RANGE = 'ALL ISSUES!C4:N1000'; // Adjust if needed

async function authenticate() {
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getAllTeamCDSSheetIds(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: UTILS_SHEET_ID,
    range: 'UTILS!B2:B',
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function getAllIssuesRichText(sheets) {
  const { data } = await sheets.spreadsheets.get({
    spreadsheetId: CENTRAL_ISSUE_SHEET_ID,
    ranges: [ALL_ISSUES_RANGE],
    includeGridData: true,
  });

  const gridData = data.sheets[0]?.data[0]?.rowData || [];

  const rows = gridData.map(row => {
    const values = row.values || [];
    return values.map(cell => {
      const richText = cell?.textFormatRuns?.length
        ? {
            userEnteredValue: cell.effectiveValue,
            richTextValue: {
              text: cell.formattedValue || '',
              link: cell.hyperlink ? { uri: cell.hyperlink } : undefined,
            },
          }
        : {
            userEnteredValue: cell.effectiveValue,
          };
    });
  });

  return rows;
}

async function getSelectedMilestones(sheets, sheetId) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${G_MILESTONES}!G4:G`,
  });
  return data.values?.flat().filter(Boolean) || [];
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

async function clearGIssues(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${G_ISSUES_SHEET}!C4:N`,
  });
}

async function insertRichDataToGIssues(sheets, sheetId, data) {
  const requests = data.map((row, rowIndex) =>
    row.map((cell, colIndex) => ({
      updateCells: {
        rows: [
          {
            values: [
              cell.richTextValue
                ? {
                    userEnteredValue: cell.userEnteredValue,
                    textFormatRuns: [
                      {
                        startIndex: 0,
                        format: {
                          link: cell.richTextValue.link,
                        },
                      },
                    ],
                  }
                : {
                    userEnteredValue: cell.userEnteredValue,
                  },
            ],
          },
        ],
        fields: '*',
        start: {
          sheetId: undefined, // fallback if using sheet name
          rowIndex: rowIndex + 3, // C4 → row 4 (0-based index)
          columnIndex: colIndex + 2, // C → col 2
        },
      },
    }))
  ).flat();

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const gIssuesSheet = spreadsheet.data.sheets.find(s => s.properties.title === G_ISSUES_SHEET);
  if (!gIssuesSheet) throw new Error(`Sheet '${G_ISSUES_SHEET}' not found`);

  const batchUpdate = {
    spreadsheetId: sheetId,
    requestBody: {
      requests: requests.map(r => {
        r.updateCells.start.sheetId = gIssuesSheet.properties.sheetId;
        return r;
      }),
    },
  };

  await sheets.spreadsheets.batchUpdate(batchUpdate);
}

async function processSheet(sheets, sheetId, allIssuesRich) {
  const spreadsheetMeta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const sheetTitles = spreadsheetMeta.data.sheets.map(s => s.properties.title);

  if (!sheetTitles.includes(G_MILESTONES) || !sheetTitles.includes(G_ISSUES_SHEET)) {
    console.warn(`⚠️ Skipping ${sheetId} — missing required sheets`);
    return;
  }

  const milestones = await getSelectedMilestones(sheets, sheetId);

  const filtered = allIssuesRich.filter(row => {
    const milestone = row[6]?.userEnteredValue?.stringValue || '';
    return milestones.includes(milestone);
  });

  const trimmed = filtered.map(row => row.slice(0, 11)); // C to N

  await clearGIssues(sheets, sheetId);
  await insertRichDataToGIssues(sheets, sheetId, trimmed);
  await updateTimestamp(sheets, sheetId);

  console.log(`✅ Processed: ${sheetId}`);
}

async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    const sheetIds = await getAllTeamCDSSheetIds(sheets);
    if (!sheetIds.length) {
      console.error('❌ No Team CDS sheet IDs found in UTILS!B2:B');
      return;
    }

    console.log(`📥 Fetching ALL ISSUES rich data...`);
    const allIssuesRich = await getAllIssuesRichText(sheets);

    console.log(`⚙️  Processing ${sheetIds.length} sheets in parallel...`);
    await Promise.all(
      sheetIds.map(sheetId => processSheet(sheets, sheetId, allIssuesRich))
    );
  } catch (err) {
    console.error(`❌ Fatal error: ${err.message}`);
  }
}

main();
