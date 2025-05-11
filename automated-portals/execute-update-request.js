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
  const ranges = rowIndices.map(rowIndex => `${SHEET_NAME}!${rowIndex}:${rowIndex}`);
  if (ranges.length === 0) return;

  await sheets.spreadsheets.values.batchClear({
    spreadsheetId: SHEET_ID,
    requestBody: { ranges }
  });

  console.log(`Cleared ${ranges.length} entire rows from Logs sheet.`);
}

async function logData(auth, message) {
  const sheets = google.sheets({ version: 'v4', auth });
  const logCell = 'B1'; // Reference to cell B1 for logging
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!${logCell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[message]] }
  });
  console.log(message);
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24'];
  const ranges = cellRefs.map(ref => `${sheetTitle}!${ref}`);

  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges,
  });

  const data = {};
  res.data.valueRanges.forEach((range, index) => {
    const value = range.values?.[0]?.[0] || null;
    data[cellRefs[index]] = value;
  });

  if (!data['C24']) return null;

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


async function getCellValue(auth, spreadsheetId, range) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values?.[0]?.[0] || null;
}

async function processUrl(url, auth) {
  const targetSpreadsheetId = url.match(/[-\w]{25,}/)[0];
  const sheets = google.sheets({ version: 'v4', auth });
  const targetSpreadsheet = await sheets.spreadsheets.get({ spreadsheetId: targetSpreadsheetId });
  
  const allData = [];
  const processedSheets = [];

  for (const sheet of targetSpreadsheet.data.sheets) {
    const sheetTitle = sheet.properties.title;
    if (SHEETS_TO_SKIP.includes(sheetTitle)) continue;

    const data = await collectSheetData(auth, targetSpreadsheetId, sheetTitle);
    if (data) {
      allData.push(data);
      processedSheets.push(sheetTitle);
    }
  }

  if (processedSheets.length > 0) {
    await logData(auth, `Fetched sheets: ${processedSheets.join(", ")}`);
  }

  if (allData.length > 0) {
    for (const data of allData) {
      await validateAndInsertData(auth, data);
    }
  } else {
    await logData(auth, `No valid data found in fetched sheets from URL: ${url}`);
  }
}

async function validateAndInsertData(auth, data) {
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
      await clearRowData(auth, sheetTitle, existingC3Index);
      await insertDataInRow(auth, sheetTitle, existingC3Index, data);
      await logData(auth, `Updated row ${existingC3Index} in sheet '${sheetTitle}' with data: ${JSON.stringify(data)}`);
      return true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await insertDataInRow(auth, sheetTitle, newRowIndex, data);
      await logData(auth, `Inserted row after ${lastC24Index} in sheet '${sheetTitle}' with data: ${JSON.stringify(data)}`);
      return true;
    }
  }

  await logData(auth, `Neither C24 ('${data.C24}') nor C3 ('${data.C3}') found for insertion in any sheet.`);
  return false;
}

async function insertDataInRow(auth, sheetTitle, row, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!C${row}:T${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        data.C24,
        data.C3,
        `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
        data.C5,
        data.C6,
        data.C7,
        '',
        '',
        '',
        '',
        '',
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
}

async function clearRowData(auth, sheetTitle, row) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!D${row}:T${row}`
  });
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

async function updateTestCasesInLibrary() {
  const authClient = await auth.getClient();
  const urlsWithIndices = await fetchUrls(authClient);

  if (!urlsWithIndices.length) {
    await logData(authClient, 'No URLs to process.');
    return;
  }

  await logData(authClient, `Starting processing ${Math.min(urlsWithIndices.length, MAX_URLS)} URLs...`);

  const uniqueUrls = new Set();
  const processedRowIndices = [];

  for (let i = 0; i < urlsWithIndices.length && uniqueUrls.size < MAX_URLS; i++) {
    const { url, rowIndex } = urlsWithIndices[i];
    if (uniqueUrls.has(url)) {
      await logData(authClient, `Duplicate URL found: ${url}. Clearing row data.`);
      processedRowIndices.push(rowIndex);
      continue;
    }

    uniqueUrls.add(url);
    processedRowIndices.push(rowIndex);

    await logData(authClient, `Processing URL: ${url}`);
    try {
      await processUrl(url, authClient);
    } catch (error) {
      await logData(authClient, `Error processing URL: ${url}. Error: ${error.message}`);
    }
  }

  await clearFetchedRows(authClient, processedRowIndices);
  await logData(authClient, "Processing complete.");
}

updateTestCasesInLibrary().catch(console.error);
