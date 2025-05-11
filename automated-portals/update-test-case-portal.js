import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const DEST_SHEET = 'Test Case Portal';
const START_COLUMN = 'C'; // Destination start column
const START_ROW = 3;

const EXCLUDED_SHEETS = [
  'Metrics Comparison',
  'Test Scenario Portal',
  'Scenario Extractor',
  'TEMPLATE',
  'Template',
  'Help',
  'Feature Change Log',
  'Logs',
  'UTILS'
];

const SHEET_NAME_MAP = {
  'Boards Test Cases': 'Boards',
  'Desktop Test Cases': 'Desktop',
  'Android Test Cases': 'Android',
  'HQZen Admin Test Cases': 'HQZen Administration',
  'Scalema Test Cases': 'Scalema',
  'HR/Policy Test Cases': 'HR/Policy'
};

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});

async function fetchSheetTitles(sheets) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return res.data.sheets
    .map(sheet => sheet.properties.title)
    .filter(title => !EXCLUDED_SHEETS.includes(title));
}

async function fetchSheetData(sheets, sheetName) {
  const range = `${sheetName}!B3:W`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
    valueRenderOption: 'FORMULA' // Important to preserve hyperlinks
  });

  const values = res.data.values || [];
  // Filter rows where columns B to D are non-empty
  return values.filter(row => row[0] && row[1] && row[2]);
}

async function insertBatchData(sheets, rows) {
  const range = `${DEST_SHEET}!${START_COLUMN}${START_ROW}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: rows
    }
  });
}

async function main() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  const sheetTitles = await fetchSheetTitles(sheets);

  let allRows = [];

  for (const sheetTitle of sheetTitles) {
    const label = SHEET_NAME_MAP[sheetTitle];
    if (!label) continue;

    const data = await fetchSheetData(sheets, sheetTitle);
    const labeledData = data.map(row => [label, ...row]); // Column B = label
    allRows = [...allRows, ...labeledData];
  }

  if (allRows.length === 0) {
    console.log('No data found to insert.');
    return;
  }

  await insertBatchData(sheets, allRows);
  console.log('Data successfully inserted into Test Case Portal.');
}

main().catch(err => {
  console.error('Failed to update Test Case Portal:', err.message);
});
