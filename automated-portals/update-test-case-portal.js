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
    // Check if the cell contains a hyperlink formula
    if (typeof cell === 'string' && cell.startsWith('=HYPERLINK')) {
      const matches = cell.match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/);  // Capture both URL and description
      if (matches && matches[1] && matches[2]) {
        const url = matches[1];
        const description = matches[2];
        // Return the correct HYPERLINK formula as a string
        return [`=HYPERLINK("${url}", "${description}")`];  // Ensure the formula is passed as a string
      }
    }
    return [cell];  // Return the original cell value if no hyperlink
  });
}

async function fetchSheetData(sheets, sheetName) {
  const range = `${sheetName}!B3:W`; 
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
    valueRenderOption: 'FORMULA'  // Ensure formulas are rendered as formulas
  });

  const values = res.data.values || [];

  // Filter out rows where any of the 3 columns (B, C, D) are empty
  return values.filter(row => row[1] && row[2] && row[3]);
}

async function clearTargetSheet(sheets) {
  const range = `${DEST_SHEET}!B3:X`;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range
  });
}

async function insertBatchData(sheets, rows) {
  const range = `${DEST_SHEET}!B3`;  // Start inserting at row 3
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',  // Ensures formulas are recognized
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
    if (!label) continue;  // Skip if no corresponding label is found

    const data = await fetchSheetData(sheets, sheetTitle);
    const labeledData = data.map(row => {
      // Process each row, add the label, and handle hyperlinks
      const processedRow = detectHyperlinks(row);
      return [label, ...processedRow];  // Add label to the first column
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
