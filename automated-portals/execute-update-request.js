import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Logs';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues', "HELP"];
const MAX_URLS = 5;
const RATE_LIMIT_DELAY = 10000; // seconds delay between requests

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
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24', 'B27', 'C32', 'C11'];
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
    B27: data['B27'],
    C5: data['C5'],
    C6: data['C6'],
    C7: data['C7'],
    C11: data['C11'],
    C32: data['C32'],
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

async function processUrl(url, auth) {
  const targetSpreadsheetId = url.match(/[-\w]{25,}/)[0];
  const sheets = google.sheets({ version: 'v4', auth });

  const meta = await sheets.spreadsheets.get({ spreadsheetId: targetSpreadsheetId });
  const sheetTitles = meta.data.sheets
    .map(s => s.properties.title)
    .filter(title => !SHEETS_TO_SKIP.includes(title));

  const dataPromises = sheetTitles.map(sheetTitle => collectSheetData(auth, targetSpreadsheetId, sheetTitle));
  const collectedData = await Promise.all(dataPromises);

  const validData = collectedData.filter(data => data !== null && Object.values(data).some(v => v !== null && v !== ''));

  for (const data of validData) {
    await logData(auth, `Fetched data from sheet: ${data.sheetName}`);
    await validateAndInsertData(auth, data);
  }
}

async function validateAndInsertData(auth, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  const targetSheetTitles = await getTargetSheetTitles(auth);
  let processed = false;

  for (const sheetTitle of targetSheetTitles) {
    if (SHEETS_TO_SKIP.includes(sheetTitle)) continue;

    const isAllTestCases = sheetTitle === "ALL TEST CASES";
    const validateCol1 = isAllTestCases ? 'C' : 'B';
    const validateCol2 = isAllTestCases ? 'D' : 'C';

    const firstColumn = await getColumnValues(auth, sheetTitle, validateCol1);
    const secondColumn = await getColumnValues(auth, sheetTitle, validateCol2);

    let lastC24Index = -1;
    let existingC3Index = -1;

    for (let i = 0; i < firstColumn.length; i++) {
      if (firstColumn[i] === data.C24) lastC24Index = i + 1;
      if (secondColumn[i] === data.C3) {
        existingC3Index = i + 1;
        break;
      }
    }

    if (existingC3Index !== -1) {
      await clearRowData(auth, sheetTitle, existingC3Index, isAllTestCases);
      await insertDataInRow(auth, sheetTitle, existingC3Index, data, isAllTestCases ? 'C' : 'B', isAllTestCases ? 'R' : 'Q');
      await logData(auth, `Updated row ${existingC3Index} in sheet '${sheetTitle}'`);
      processed = true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await insertRowWithFormat(auth, sheetTitle, lastC24Index);
      await insertDataInRow(auth, sheetTitle, newRowIndex, data, isAllTestCases ? 'C' : 'B', isAllTestCases ? 'R' : 'Q');
      await logData(auth, `Inserted row after ${lastC24Index} in sheet '${sheetTitle}'`);
      processed = true;
    }

    // Rate limiting to avoid quota issues
    await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY));
  }

  if (!processed) {
    await logData(auth, `No matches found for C24 ('${data.C24}') or C3 ('${data.C3}') in any sheet.`);
  }

  return processed;
}

async function insertRowWithFormat(auth, sheetTitle, sourceRowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: await getSheetId(auth, sheetTitle),
              dimension: 'ROWS',
              startIndex: sourceRowIndex, // zero-based
              endIndex: sourceRowIndex + 1
            },
            inheritFromBefore: true // Inherit formulas and data validation from the row above
          }
        }
      ]
    }
  });

  console.log(`Inserted new row after row ${sourceRowIndex} in sheet '${sheetTitle}' with formatting.`);
}

async function getSheetId(auth, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetTitle);
  return sheet.properties.sheetId;
}

// Converts a column letter like 'A' to a number (A=1, B=2, ..., Z=26, AA=27, etc.)
function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num *= 26;
    num += col.charCodeAt(i) - 64;
  }
  return num;
}

// Converts a column number like 27 back to a letter (27 = 'AA')
function numberToColumn(n) {
  let result = '';
  while (n > 0) {
    let remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

async function insertDataInRow(auth, sheetTitle, row, data, startCol) {
  const sheets = google.sheets({ version: 'v4', auth });

  const isAllTestCases = sheetTitle === "ALL TEST CASES";

  const rowValues = [
    data.C24,
    data.C3,
    `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
    data.B27,
    data.C5,
    data.C6,
    data.C7,
    '', // This blank cell is intentionally left
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

  // Add one more value if sheet is 'ALL TEST CASES'
  if (isAllTestCases) {
    rowValues.push(data.C21);
  }

  // Calculate range dynamically
  const startIndex = columnToNumber(startCol);
  const endIndex = startIndex + rowValues.length - 1;
  const endCol = numberToColumn(endIndex);

  const range = `${sheetTitle}!${startCol}${row}:${endCol}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [rowValues]
    }
  });
}


async function clearRowData(auth, sheetTitle, row, isAllTestCases) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = isAllTestCases ? `D${row}:T${row}` : `C${row}:S${row}`;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!${range}`
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

  await logData(authClient, 'Starting processing URLs...');

  const uniqueUrls = new Set();
  const processedRowIndices = [];

  for (const { url, rowIndex } of urlsWithIndices) {
    if (uniqueUrls.has(url)) {
      await logData(authClient, `Duplicate URL found: ${url}.`);
      continue;
    }

    uniqueUrls.add(url);
    processedRowIndices.push(rowIndex);
    await clearFetchedRows(authClient, processedRowIndices);
    await logData(authClient, `Cleared row for URL: ${url}`);

    await logData(authClient, `Processing URL: ${url}`);
    try {
      await processUrl(url, authClient);
    } catch (error) {
      if (error.message.includes('Quota exceeded')) {
        await logData(authClient, `Quota exceeded for URL: ${url}. Retrying...`);
        await new Promise(resolve => setTimeout(resolve, 90000)); // Wait for 1.5 minute before retrying
        try {
          await processUrl(url, authClient);
        } catch (retryError) {
          await logData(authClient, `Error processing URL on retry: ${url}. Error: ${retryError.message}`);
        }
      } else {
        await logData(authClient, `Error processing URL: ${url}. Error: ${error.message}`);
      }
    }
  }

  await logData(authClient, "Processing complete.");
}

updateTestCasesInLibrary().catch(console.error);
