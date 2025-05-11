import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const DEST_SHEET = 'Test Case Portal';
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

function detectHyperlinks(row) {
  return row.map(cell => {
    if (typeof cell === 'string' && cell.startsWith('=HYPERLINK')) {
      const matches = cell.match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/);
      if (matches && matches[1] && matches[2]) {
        return `=HYPERLINK("${matches[1]}", "${matches[2]}")`;
      }
    }
    return cell;
  });
}

function convertSerialDate(serial) {
  // Google Sheets date serials are days since 1899-12-30
  if (!serial || isNaN(serial)) return '';
  const baseDate = new Date(Date.UTC(1899, 11, 30));
  baseDate.setUTCDate(baseDate.getUTCDate() + parseInt(serial));
  return baseDate.toISOString().split('T')[0]; // Returns 'YYYY-MM-DD'
}

async function fetchSheetData(sheets, sheetName) {
  const range = `${sheetName}!B3:Y`; // Include column Y
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE'  // Get raw values (including serial dates)
  });

  const values = res.data.values || [];

  // Filter rows with B, C, D non-empty
  return values.filter(row => row[1] && row[2] && row[3]);
}

async function clearTargetSheet(sheets) {
  const range = `${DEST_SHEET}!B3:X`;  // May be updated to Y if needed
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range
  });
}

async function insertBatchData(sheets, rows) {
  const range = `${DEST_SHEET}!B3`;
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

  await clearTargetSheet(sheets);
  console.log('Target sheet cleared from B3 to X.');

  const sheetTitles = await fetchSheetTitles(sheets);
  let allRows = [];

  for (const sheetTitle of sheetTitles) {
    const label = SHEET_NAME_MAP[sheetTitle];
    if (!label) continue;

    const data = await fetchSheetData(sheets, sheetTitle);
    const labeledData = data.map(row => {
      const processedRow = detectHyperlinks(row);

      // Convert O (column index 13) and Y (column index 23) to readable dates
      if (processedRow[13]) processedRow[13] = convertSerialDate(processedRow[13]);
      if (processedRow[23]) processedRow[23] = convertSerialDate(processedRow[23]);

      return [label, ...processedRow];
    });

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
