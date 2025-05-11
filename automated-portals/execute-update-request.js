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
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly',
  ],
});

async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!B3:B`;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];
  return values
    .map((row, index) => ({ url: row[0], rowIndex: index + 3 }))
    .filter(entry => entry.url);
}

async function clearFetchedRows(auth, rowIndices) {
  if (!rowIndices.length) return;

  const sheets = google.sheets({ version: 'v4', auth });
  const ranges = rowIndices.map(i => `${SHEET_NAME}!${i}:${i}`);
  await sheets.spreadsheets.values.batchClear({
    spreadsheetId: SHEET_ID,
    requestBody: { ranges },
  });

  console.log(`Cleared ${ranges.length} rows from Logs sheet.`);
}

async function logData(auth, message) {
  const sheets = google.sheets({ version: 'v4', auth });
  const logCell = 'B1';
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!${logCell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[message]] },
  });
  console.log(message);
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24'];
  const ranges = cellRefs.map(ref => `${sheetTitle}!${ref}`);

  const res = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges });

  const data = {};
  res.data.valueRanges.forEach((range, i) => {
    data[cellRefs[i]] = range.values?.[0]?.[0] || null;
  });

  if (!data.C24) return null;

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.title === sheetTitle);
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.properties.sheetId}`;

  return { ...data, sheetUrl, sheetName: sheetTitle };
}

async function getCellValue(auth, spreadsheetId, range) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values?.[0]?.[0] || null;
}

async function processUrl(url, auth) {
  try {
    const spreadsheetId = url.match(/[-\w]{25,}/)[0];
    const sheetsApi = google.sheets({ version: 'v4', auth });
    const spreadsheet = await sheetsApi.spreadsheets.get({ spreadsheetId });

    const allData = [];

    for (const sheet of spreadsheet.data.sheets) {
      const title = sheet.properties.title;
      if (SHEETS_TO_SKIP.includes(title)) continue;

      const data = await collectSheetData(auth, spreadsheetId, title);
      const isValid = data && Object.values(data).some(v => v);

      if (!isValid) {
        await logData(auth, `Skipped '${title}' – no content.`);
        continue;
      }

      allData.push(data);
    }

    if (allData.length) {
      for (const d of allData) {
        await validateAndInsertData(auth, d);
      }
    } else {
      await logData(auth, `No valid data from: ${url}`);
    }
  } catch (err) {
    console.error(`Error processing URL ${url}:`, err);
  }
}

async function getTargetSheetTitles(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return res.data.sheets.map(s => s.properties.title);
}

async function getColumnValues(auth, sheetTitle, col) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${sheetTitle}!${col}2:${col}`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  return (res.data.values || []).flat();
}

async function clearRowData(auth, sheetTitle, rowIndex, isAllTestCases) {
  const sheets = google.sheets({ version: 'v4', auth });
  const startCol = isAllTestCases ? 'C' : 'B';
  const endCol = isAllTestCases ? 'T' : 'S';
  const range = `${sheetTitle}!${startCol}${rowIndex}:${endCol}${rowIndex}`;
  await sheets.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range });
}

async function insertDataInRow(auth, sheetTitle, row, data, startCol, endCol) {
  const sheets = google.sheets({ version: 'v4', auth });

  const values = [
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
  ];

  if (sheetTitle === 'ALL TEST CASES') values.push(data.C21);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!${startCol}${row}:${endCol}${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] },
  });
}

async function insertRowWithFormat(auth, sheetTitle, sourceRowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = await getSheetId(auth, sheetTitle);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        insertDimension: {
          range: {
            sheetId,
            dimension: 'ROWS',
            startIndex: sourceRowIndex,
            endIndex: sourceRowIndex + 1,
          },
          inheritFromBefore: true,
        },
      }],
    },
  });

  console.log(`Inserted row at ${sourceRowIndex} in '${sheetTitle}' with formatting.`);
}

async function getSheetId(auth, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetTitle);
  return sheet?.properties.sheetId;
}

async function validateAndInsertData(auth, data) {
  const targetSheets = await getTargetSheetTitles(auth);
  let processed = false;

  for (const title of targetSheets) {
    if (SHEETS_TO_SKIP.includes(title)) continue;

    const isAllTestCases = title === 'ALL TEST CASES';
    const col1 = isAllTestCases ? 'C' : 'B';
    const col2 = isAllTestCases ? 'D' : 'C';

    const col1Vals = await getColumnValues(auth, title, col1);
    const col2Vals = await getColumnValues(auth, title, col2);

    let lastC24Index = -1;
    let existingC3Index = -1;

    for (let i = 0; i < col1Vals.length; i++) {
      if (col1Vals[i] === data.C24) lastC24Index = i + 1;
      if (col2Vals[i] === data.C3) {
        existingC3Index = i + 1;
        break;
      }
    }

    if (existingC3Index !== -1) {
      await clearRowData(auth, title, existingC3Index, isAllTestCases);
      await insertDataInRow(auth, title, existingC3Index, data, col1, isAllTestCases ? 'T' : 'S');
      await logData(auth, `Updated row ${existingC3Index} in '${title}'`);
      processed = true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await insertRowWithFormat(auth, title, lastC24Index);
      await insertDataInRow(auth, title, newRowIndex, data, col1, isAllTestCases ? 'T' : 'S');
      await logData(auth, `Inserted row after ${lastC24Index} in '${title}'`);
      processed = true;
    }
  }

  if (!processed) {
    await logData(auth, `No match for C24 ('${data.C24}') or C3 ('${data.C3}') in any sheet.`);
  }
  return processed;
}
