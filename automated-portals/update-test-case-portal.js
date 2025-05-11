import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const TARGET_SHEET = 'Test Case Portal';
const EXCLUDED_SHEETS = [
  'Metrics Comparison', 'Test Scenario Portal', 'Scenario Extractor',
  'TEMPLATE', 'Template', 'Help', 'Feature Change Log', 'Logs', 'UTILS'
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

async function main() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  // 1. Get list of all sheet names
  const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheetTitles = sheetMeta.data.sheets.map(s => s.properties.title);
  const filteredSheets = sheetTitles.filter(
    title => !EXCLUDED_SHEETS.includes(title) && SHEET_NAME_MAP[title]
  );

  const allRowsToInsert = [];

  for (const title of filteredSheets) {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${title}'!B3:W`
    });

    const values = response.data.values || [];

    // Filter out rows where B, C, or D is missing
    const filtered = values.filter(row =>
      row[0]?.trim() && row[1]?.trim() && row[2]?.trim()
    );

    const mappedRows = filtered.map(row => [
      SHEET_NAME_MAP[title], // Column B value
      ...row                 // Columns C to W (shifted to C:X)
    ]);

    allRowsToInsert.push(...mappedRows);
  }

  // 2. Clear existing rows in the target range before inserting (optional)
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `${TARGET_SHEET}!B3:X`
  });

  // 3. Insert all data at once starting at B3
  if (allRowsToInsert.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${TARGET_SHEET}!B3`,
      valueInputOption: 'RAW',
      requestBody: {
        values: allRowsToInsert
      }
    });
    console.log(`Inserted ${allRowsToInsert.length} rows to "${TARGET_SHEET}".`);
  } else {
    console.log('No data found to insert.');
  }
}

main().catch(console.error);
