import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases'; //Fetch from "Boards Test Cases"
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues'];
const MAX_URLS = 20;
const RATE_LIMIT_DELAY = 5000; // 5 seconds delay between requests

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

// Fetch URLs from the D column of "Boards Test Cases"
async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!D3:D`;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];

  const urls = await Promise.all(values.map(async (row, index) => {
    const cell = `D${index + 3}`;
    if (index + 3 > 17) return null; // Skip if index exceeds max rows

    try {
      // Check if the row contains text
      const text = row[0] || null; // Get the text from the row
      if (!text) {
        console.error(`No text found for cell ${cell}`);
        return null; // Return null if no text is found
      }

      // Attempt to extract a URL from the text if it's not a hyperlink
      const linkResponse = await sheets.spreadsheets.get({
        spreadsheetId: SHEET_ID,
        ranges: [cell],
        fields: 'sheets.data.rowData.values.hyperlink'
      });

      const hyperlink = linkResponse.data.sheets &&
                        linkResponse.data.sheets[0] &&
                        linkResponse.data.sheets[0].data &&
                        linkResponse.data.sheets[0].data[0] &&
                        linkResponse.data.sheets[0].data[0].rowData &&
                        linkResponse.data.sheets[0].data[0].rowData[0] &&
                        linkResponse.data.sheets[0].data[0].rowData[0].values &&
                        linkResponse.data.sheets[0].data[0].rowData[0].values[0] &&
                        linkResponse.data.sheets[0].data[0].rowData[0].values[0].hyperlink;

      // If hyperlink is null, return the text as a fallback
      return { url: hyperlink || text, rowIndex: index + 3 };
    } catch (error) {
      console.error(`Error processing URL for cell ${cell}:`, error);
      return null; // Return null if there's an error
    }
  }));

  return urls.filter(entry => entry && entry.url);
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

async function insertDataInRow(auth, sheetTitle, row, data, startCol, endCol) {
  const sheets = google.sheets({ version: 'v4', auth });

  const isAllTestCases = sheetTitle === "ALL TEST CASES";

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
    data.C20
  ];

  if (isAllTestCases) {
    values.push(data.C21);
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!${startCol}${row}:${endCol}${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values]
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

  for (const { url } of urlsWithIndices) {
    if (uniqueUrls.has(url)) {
      await logData(authClient, `Duplicate URL found: ${url}.`);
      continue;
    }

    uniqueUrls.add(url);
    await logData(authClient, `Processing URL: ${url}`);
    try {
      await processUrl(url, authClient);
    } catch (error) {
      if (error.message.includes('Quota exceeded')) {
        await logData(authClient, `Quota exceeded for URL: ${url}. Retrying...`);
        await new Promise(resolve => setTimeout(resolve, 90000)); // Wait for 1.5 minutes before retrying
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
