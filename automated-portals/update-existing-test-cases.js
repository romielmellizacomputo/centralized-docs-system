import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';
import pLimit from 'p-limit';

dotenv.config();

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

const sheets = google.sheets({ version: 'v4', auth });
const sheetId = process.env.CDS_PORTAL_SPREADSHEET_ID; // Set your spreadsheet ID here
const sheetTitle = "Boards Test Cases"; // The title of the sheet to work with

const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'B27', 'C32', 'C11'];

async function fetchDataFromSheet() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `${sheetTitle}!D:D`, // Fetching URLs from column D
  });

  return response.data.values || [];
}

async function fetchDataFromCells(rowIndex) {
  const requests = cellRefs.map(ref => `${sheetTitle}!${ref}${rowIndex}`);
  const response = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: sheetId,
    ranges: requests,
  });

  const data = {};
  response.data.valueRanges.forEach((range, index) => {
    data[cellRefs[index]] = range.values ? range.values[0][0] : null;
  });

  // Extract the embedded Google Sheet URL from the hyperlink text in D column
  const sheetUrl = data.C4 ? data.C4.match(/https?:\/\/[^"]+/)[0] : null;
  return { ...data, sheetUrl };
}

async function insertDataInRow(row, data) {
  const values = [
    data.C24,
    data.C3,
    `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
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
    spreadsheetId: sheetId,
    range: `${sheetTitle}!B${row}:R${row}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

async function processSheet() {
  const urls = await fetchDataFromSheet();

  const limit = pLimit(1); // Limit to 1 concurrent request
  for (let rowIndex = 1; rowIndex <= urls.length; rowIndex++) { // Start from 1 to skip header
    const url = urls[rowIndex - 1][0]; // Get URL from the D column
    if (url) {
      const data = await fetchDataFromCells(rowIndex + 1); // Fetch data from specified cells
      await limit(() => insertDataInRow(rowIndex + 1, data)); // Insert data into the same row
      await new Promise(resolve => setTimeout(resolve, 1000)); // Delay to avoid hitting API quota
    }
  }
}

processSheet().catch(console.error);
