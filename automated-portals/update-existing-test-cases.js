import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';
import pLimit from 'p-limit';

dotenv.config();

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly'],
});

const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24', 'B27', 'C32', 'C11'];

const limit = pLimit(1); // 1 concurrent task
const delay = (ms) => new Promise((res) => setTimeout(res, ms));

async function getSheetData(sheetId, range) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range });
  return res.data.values?.[0]?.[0] || '';
}

async function extractHyperlinkFormula(sheetId, range) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const res = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    ranges: [range],
    includeGridData: true,
  });

  const cell = res.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0];
  const formula = cell?.userEnteredValue?.formulaValue;
  if (!formula) return null;

  const match = formula.match(/=HYPERLINK\("([^"]+)"/);
  return match ? match[1] : null;
}

async function fetchAndMapData(sheetUrl) {
  const match = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) return null;

  const sheetId = match[1];
  const cellValues = {};
  for (const cell of cellRefs) {
    const val = await getSheetData(sheetId, cell);
    cellValues[cell] = val;
  }

  return {
    sheetUrl,
    ...cellValues,
  };
}

async function insertDataInRow(auth, sheetId, row, data) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

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

  const range = `Boards Test Cases!B${row}:R${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] },
  });
}

async function processSheet() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  const sheetId = process.env.CDS_PORTAL_SPREADSHEET_ID;
  const readRange = 'Boards Test Cases!D2:D';

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: readRange,
  });

  const rows = res.data.values || [];

  for (let i = 0; i < rows.length; i++) {
    const rowNumber = i + 2; // accounting for starting at row 2
    const value = rows[i][0];

    if (!value) continue;

    const sheetUrl = await extractHyperlinkFormula(sheetId, `Boards Test Cases!D${rowNumber}`);
    if (!sheetUrl) continue;

    try {
      const data = await fetchAndMapData(sheetUrl);
      if (data) {
        await insertDataInRow(auth, sheetId, rowNumber, data);
        console.log(`✅ Row ${rowNumber} processed.`);
      }
    } catch (err) {
      console.error(`❌ Failed to process row ${rowNumber}:`, err.message);
    }

    await delay(1000); // 1 second delay to avoid hitting API limits
  }

  console.log('✅ All rows processed.');
}

processSheet();
