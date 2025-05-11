import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';
dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Logs';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues'];
const MAX_URLS = 20;

async function authGoogle() {
  const auth = new GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly'],
    keyFile: 'credentials.json' // Store this securely!
  });
  return await auth.getClient();
}

async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!B3:B`;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];
  return values.flat().filter(url => url);
}

async function clearFetchedRows(auth, numRowsToClear) {
  const sheets = google.sheets({ version: 'v4', auth });
  if (numRowsToClear === 0) return;

  await sheets.spreadsheets.values.batchClear({
    spreadsheetId: SHEET_ID,
    requestBody: {
      ranges: [`${SHEET_NAME}!B3:B${3 + numRowsToClear - 1}`]
    }
  });

  console.log(`Cleared ${numRowsToClear} rows from Logs sheet.`);
}

async function processSpreadsheetUrl(url, auth) {
  try {
    const match = url.match(/[-\w]{25,}/); // Extract spreadsheet ID from URL
    if (!match) throw new Error('Invalid sheet URL');
    const sheetId = match[0];

    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
    const sheetTitles = meta.data.sheets.map(s => s.properties.title);

    const processedSheets = [];

    for (const title of sheetTitles) {
      if (SHEETS_TO_SKIP.includes(title)) continue;

      const range = `${title}!C24`;
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range
      });

      const C24 = res.data.values?.[0]?.[0] || null;
      if (!C24) continue;

      console.log(`Fetched C24 from sheet: ${title} in ${url}`);
      processedSheets.push(title);
    }

    return processedSheets.length;
  } catch (err) {
    console.error(`Error processing ${url}: ${err.message}`);
    return 0;
  }
}

async function updateTestCasesInLibrary() {
  const auth = await authGoogle();
  const urls = await fetchUrls(auth);

  if (!urls.length) {
    console.log('No URLs to process.');
    return;
  }

  console.log(`Starting processing ${Math.min(urls.length, MAX_URLS)} URLs...`);

  const uniqueUrls = new Set();
  let processedCount = 0;

  for (let i = 0; i < urls.length && processedCount < MAX_URLS; i++) {
    const url = urls[i];
    if (uniqueUrls.has(url)) continue;

    const result = await processSpreadsheetUrl(url, auth);
    if (result > 0) {
      uniqueUrls.add(url);
      processedCount++;
    }
  }

  await clearFetchedRows(auth, processedCount);
  console.log('Processing complete.');
}

updateTestCasesInLibrary().catch(console.error);
