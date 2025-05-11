import { google } from 'googleapis';

const sheetData = JSON.parse(process.env.SHEET_DATA); // Can be an object or array
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);
const targetSpreadsheetId = process.env.AUTOMATED_PORTALS;

async function sendUpdateSignal() {
  try {
    console.log('📤 Starting signal send to Logs sheet');

    if (!targetSpreadsheetId) {
      throw new Error('Missing required environment variable: AUTOMATED_PORTALS');
    }

    const targetSheetName = 'Logs';

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const currentDate = new Date().toISOString();
    const updates = Array.isArray(sheetData) ? sheetData : [sheetData];

    const logEntries = [];

    for (const entry of updates) {
      const spreadsheetId = entry.spreadsheetId;
      const sheetName = entry.sheetName;
      const editedRange = entry.editedRange || '';

      if (!spreadsheetId || !sheetName) {
        throw new Error(`Missing spreadsheetId or sheetName in entry: ${JSON.stringify(entry)}`);
      }

      // Fetch sheet metadata to get sheetId (gid)
      const metadata = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties',
      });

      const matchingSheet = metadata.data.sheets?.find(
        (s) => s.properties?.title === sheetName
      );

      if (!matchingSheet || matchingSheet.properties?.sheetId === undefined) {
        throw new Error(`Could not find sheet ID for "${sheetName}" in spreadsheet ${spreadsheetId}`);
      }

      const sheetId = matchingSheet.properties.sheetId;
      const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;
      const logMessage = `Sheet: ${sheetName} | Range: ${editedRange}`;

      logEntries.push([currentDate, sheetUrl, logMessage]);
    }

    // Append to the Logs sheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: `${targetSheetName}!A:C`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: logEntries,
      },
    });

    console.log(`✅ Log(s) added to Logs sheet at ${targetSpreadsheetId}`);
  } catch (error) {
    console.error('❌ Error sending log signal:', error.message);
  }
}

sendUpdateSignal();
