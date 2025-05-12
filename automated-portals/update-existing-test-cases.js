import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';
import pLimit from 'p-limit';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases';  // Change to the actual sheet name
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues']; // Skip sheets that you don't need to process
const MAX_CONCURRENT_REQUESTS = 3;  // Max concurrent requests
const RATE_LIMIT_DELAY = 3000; // 3 seconds delay between requests

// Set up Google Auth
const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

// Fetch URLs from the Google Sheets document
async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!D3:D`; // Assuming URLs are in column D starting from row 3
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];

  const limit = pLimit(MAX_CONCURRENT_REQUESTS);

  const tasks = values.map((row, index) => limit(async () => {
    const rowIndex = index + 3;
    const text = row[0] || null;
    if (!text) return null;

    const cellRange = `${SHEET_NAME}!D${rowIndex}`;

    try {
      const linkResponse = await sheets.spreadsheets.get({
        spreadsheetId: SHEET_ID,
        ranges: [cellRange],
        includeGridData: true,
        fields: 'sheets.data.rowData.values.hyperlink'
      });

      const hyperlink = linkResponse.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0]?.hyperlink;
      if (hyperlink) return { url: hyperlink, rowIndex };

      const urlMatch = text.match(/https?:\/\/\S+/);
      return urlMatch ? { url: urlMatch[0], rowIndex } : null;
    } catch (error) {
      console.error(`Error at D${rowIndex}:`, error.message);
      return null;
    }
  }));

  const results = await Promise.all(tasks);
  const filteredResults = results.filter(Boolean);
  const seen = new Set();
  return filteredResults.filter(({ url }) => {
    if (seen.has(url)) return false;
    seen.add(url);
    return true;
  });
}

// Log data back into the sheet (e.g., to update test coverage stats)
async function logData(auth, message) {
  const sheets = google.sheets({ version: 'v4', auth });
  const logCell = 'B1';
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!${logCell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[message]] }
  });
  console.log(message);
}

// Process a specific sheet to count test cases and valid cases
async function processSheet(sheetName, auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  try {
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A1:Z`
    });

    const header = data.values?.[0] || [];
    const testCaseColumnIndex = header.findIndex(val => val === 'Test Case');
    const statusColumnIndex = header.findIndex(val => val === 'Status');

    if (testCaseColumnIndex === -1 || statusColumnIndex === -1) {
      return { sheetName, validCases: 0, totalCases: 0 };
    }

    const rows = data.values.slice(1);
    const totalCases = rows.length;
    const validCases = rows.reduce((count, row) => {
      const status = row[statusColumnIndex]?.toLowerCase();
      return (status === 'pass' || status === 'fail') ? count + 1 : count;
    }, 0);

    return { sheetName, validCases, totalCases };
  } catch (error) {
    console.error(`Failed to process sheet "${sheetName}": ${error.message}`);
    return { sheetName, validCases: 0, totalCases: 0 };
  }
}

// Main execution function
async function main() {
  try {
    const authClient = await auth.getClient();
    const urls = await fetchUrls(authClient);

    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const { data: meta } = await sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID,
      fields: 'sheets.properties.title'
    });

    const sheetNames = meta.sheets
      .map(sheet => sheet.properties.title)
      .filter(name => !SHEETS_TO_SKIP.includes(name));

    const limit = pLimit(MAX_CONCURRENT_REQUESTS);
    const results = [];

    for (const sheetName of sheetNames) {
      const result = await limit(() => processSheet(sheetName, authClient));
      results.push(result);
      await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY));
    }

    const totalCases = results.reduce((acc, val) => acc + val.totalCases, 0);
    const validCases = results.reduce((acc, val) => acc + val.validCases, 0);
    const coverage = totalCases ? ((validCases / totalCases) * 100).toFixed(2) : '0.00';

    await logData(authClient, `Test Coverage: ${coverage}% (${validCases}/${totalCases})`);
  } catch (error) {
    console.error('Fatal error:', error.message);
  }
}

main();
