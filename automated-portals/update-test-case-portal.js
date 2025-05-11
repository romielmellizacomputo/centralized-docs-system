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

// Fetch all sheet names excluding the ones in EXCLUDED_SHEETS
async function fetchSheetTitles(sheets) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return res.data.sheets
    .map(sheet => sheet.properties.title)
    .filter(title => !EXCLUDED_SHEETS.includes(title));
}

// Fetch data from each sheet, starting from C3 to X (column B should be empty)
async function fetchSheetData(sheets, sheetName) {
  const range = `${sheetName}!C3:X`;  // Fetch from C3 to X (columns C to X)
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
    valueRenderOption: 'FORMULA' // Preserve hyperlinks and formulas
  });

  const values = res.data.values || [];
  // Filter rows where columns C to X are not empty
  return values.filter(row => row.length > 0 && row.some(cell => cell !== ""));
}

// Clear the target range in 'Test Case Portal' from B3 to X
async function clearTargetSheet(sheets) {
  const range = `${DEST_SHEET}!B3:X`;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range
  });
}

// Insert data into 'Test Case Portal' sheet starting from C3 (columns C to X, with label in column B)
async function insertBatchData(sheets, rows) {
  const range = `${DEST_SHEET}!B3`; // Start from B3 (for the label and data)
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: rows
    }
  });
}

// Main function to fetch data from sheets and insert into 'Test Case Portal'
async function main() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  // Clear the target sheet before inserting new data
  await clearTargetSheet(sheets);
  console.log('Target sheet cleared from B3 to X.');

  const sheetTitles = await fetchSheetTitles(sheets);

  let allRows = [];

  // Loop through each sheet, fetch data, and add it to the target rows
  for (const sheetTitle of sheetTitles) {
    const label = SHEET_NAME_MAP[sheetTitle];
    if (!label) continue; // Skip if the sheet doesn't have a label

    const data = await fetchSheetData(sheets, sheetTitle);
    const labeledData = data.map(row => [label, ...row]); // Column B = label
    allRows = [...allRows, ...labeledData];
  }

  if (allRows.length === 0) {
    console.log('No data found to insert.');
    return;
  }

  // Insert data into the target sheet in 'Test Case Portal' starting from B3 (with label in column B)
  await insertBatchData(sheets, allRows);
  console.log('Data successfully inserted into Test Case Portal.');
}

// Run the script
main().catch(err => {
  console.error('Failed to update Test Case Portal:', err.message);
});
