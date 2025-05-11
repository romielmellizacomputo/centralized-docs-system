import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Logs';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues'];
const MAX_URLS = 20;

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!B3:B`;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];
  return values.map((row, index) => ({ url: row[0], rowIndex: index + 3 })).filter(entry => entry.url);
}

async function clearFetchedRows(auth, rowIndices) {
  const sheets = google.sheets({ version: 'v4', auth });
  const ranges = rowIndices.map(rowIndex => `${SHEET_NAME}!B${rowIndex}`);
  if (ranges.length === 0) return;

  await sheets.spreadsheets.values.batchClear({
    spreadsheetId: SHEET_ID,
    requestBody: { ranges }
  });

  console.log(`Cleared ${ranges.length} rows from Logs sheet.`);
}

async function getSheetTitles(auth, spreadsheetId) {
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  return meta.data.sheets.map(s => s.properties.title);
}

async function getCellValue(auth, spreadsheetId, range) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values?.[0]?.[0] || null;
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24'];
  const data = {};

  for (const ref of cellRefs) {
    const range = `${sheetTitle}!${ref}`;
    data[ref] = await getCellValue(auth, spreadsheetId, range);
  }

  if (!data['C24']) return null;

  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.title === sheetTitle);
  const sheetId = sheet.properties.sheetId;
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;

  return {
    C24: data['C24'],
    C3: data['C3'],
    C4: data['C4'],
    C5: data['C5'],
    C6: data['C6'],
    C7: data['C7'],
    C13: data['C13'],
    C14: data['C14'],
    C15: data['C15'],
    C18: data['C18'],
    C19: data['C19'],
    C20: data['C20'],
    C21: data['C21'],
    sheetUrl,
    sheetName: sheetTitle
  };
}

async function getTargetSheetTitles(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return meta.data.sheets.map(s => s.properties.title);
}

async function getColumnValues(auth, sheetTitle, column) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${sheetTitle}!${column}:${column}`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  return res.data.values?.map(row => row[0]) || [];
}

async function updateOrInsertData(auth, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  const targetSheetTitles = await getTargetSheetTitles(auth);

  for (const sheetTitle of targetSheetTitles) {
    if (SHEETS_TO_SKIP.includes(sheetTitle)) continue;

    const CColumn = await getColumnValues(auth, sheetTitle, 'C');
    const DColumn = await getColumnValues(auth, sheetTitle, 'D');

    let lastC24Index = -1;
    let existingC3Index = -1;

    for (let i = 0; i < CColumn.length; i++) {
      if (CColumn[i] === data.C24) lastC24Index = i + 1;
      if (DColumn[i] === data.C3) {
        existingC3Index = i + 1;
        break;
      }
    }

    if (existingC3Index !== -1) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetTitle}!D${existingC3Index}:T${existingC3Index}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[
            data.C3,
            `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
            data.C5,
            data.C6,
            data.C7,
            '', '', '', '',
            data.C15,
            data.C13,
            data.C14,
            data.C18,
            data.C19,
            data.C20,
            data.C21
          ]]
        }
      });
      console.log(`Updated row ${existingC3Index} in sheet '${sheetTitle}'`);
      return true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetTitle}!D${newRowIndex}:T${newRowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[
            data.C3,
            `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
            data.C5,
            data.C6,
            data.C7,
            '', '', '', '',
            data.C15,
            data.C13,
            data.C14,
            data.C18,
            data.C19,
            data.C20,
            data.C21
          ]]
        }
      });
      console.log(`Inserted row after ${lastC24Index} in sheet '${sheetTitle}'`);
      return true;
    }
  }

  console.warn(`Neither C24 ('${data.C24}') nor C3 ('${data.C3}') found for insertion in any sheet.`);
  return false;
}

async function updateTestCasesInLibrary() {
  const authClient = await auth.getClient();
  const urlsWithIndices = await fetchUrls(authClient);

  if (!urlsWithIndices.length) {
    console.log('No URLs to process.');
    return;
  }

  console.log(`Starting processing ${Math.min(urlsWithIndices.length, MAX_URLS)} URLs...`);

  const uniqueUrls = new Set();
  const processedRowIndices = [];

  for (let i = 0; i < urlsWithIndices.length && uniqueUrls.size < MAX_URLS; i++) {
    const { url, rowIndex } = urlsWithIndices[i];
    if (uniqueUrls.has(url)) continue;

    const match = url.match(/[-\w]{25,}/);
    if (!match) {
      console.error(`Invalid sheet URL: ${url}`);
      continue;
    }
    const sheetId = match[0];

    const sheetTitles = await getSheetTitles(authClient, sheetId);
    for (const title of sheetTitles) {
      if (SHEETS_TO_SKIP.includes(title)) continue;

      const data = await collectSheetData(authClient, sheetId, title);
      if (!data) continue;

      const success = await updateOrInsertData(authClient, data);
      if (success) {
        uniqueUrls.add(url);
        processedRowIndices.push(rowIndex);
        break;
      }
    }
  }

  await clearFetchedRows(authClient, processedRowIndices);
  console.log('Processing complete.');
}

updateTestCasesInLibrary().catch(console.error);
