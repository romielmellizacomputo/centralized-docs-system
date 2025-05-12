import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';
import pLimit from 'p-limit';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases'; // Change to the actual sheet name
const MAX_CONCURRENT_REQUESTS = 3; // Max concurrent requests
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
    const rowIndex = index + 3; // Adjust for the starting row
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
  return results.filter(Boolean);
}

// Insert data back into the sheet
async function insertData(auth, rowIndex, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  const values = [
    data.C24,
    data.C3,
    `=HYPERLINK("${data.url}", "${data.C4}")`, // Embed the hyperlink
    data.B27,
    data.C5,
    data.C6,
    data.C7,
    '',
    data.C11,
    data.C32,
    data.C15,
    data.C13,
    data.C14,
    data.C18,
    data.C19,
    data.C20,
    data.C21
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!B${rowIndex}:R${rowIndex}`, // Insert data in the same row
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

// Process the fetched URLs and insert data into the sheet
async function processUrls(auth) {
  const urls = await fetchUrls(auth);
  const sheets = google.sheets({ version: 'v4', auth });

  const limit = pLimit(MAX_CONCURRENT_REQUESTS);
  for (const { url, rowIndex } of urls) {
    const data = await fetchDataFromCells(sheets, rowIndex); // Fetch data from specified cells
    await limit(() => insertData(auth, rowIndex, data)); // Insert data into the same row
    await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY)); // Rate limit delay
  }
}

// Fetch data from specified cells
async function fetchDataFromCells(sheets, rowIndex) {
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'B27', 'C32', 'C11'];
  const requests = cellRefs.map(ref => `${SHEET_NAME}!${ref}${rowIndex}`);
  const response = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID,
    ranges: requests,
  });

  const data = {};
  response.data.valueRanges.forEach((range, index) => {
    data[cellRefs[index]] = range.values ? range.values[0][0] : null;
  });

  return data;
}

// Main execution function
async function main() {
  try {
    const authClient = await auth.getClient();
    await processUrls(authClient);
  } catch (error) {
    console.error('Fatal error:', error.message);
  }
}

main();
