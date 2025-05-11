import { google } from 'googleapis';

const sheetData = JSON.parse(process.env.SHEET_DATA); // includes source info, possibly multiple sheets
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
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const currentDate = new Date().toISOString();

    const updates = Array.isArray(sheetData) ? sheetData : [sheetData];

    const logEntries = updates.map((entry) => {
      if (!entry.spreadsheetId || !entry.sheetId) {
        throw new Error(`Missing spreadsheetId or sheetId for entry: ${JSON.stringify(entry)}`);
      }
    
      const spreadsheetId = entry.spreadsheetId;
      const sheetId = entry.sheetId;
      const sheetName = entry.sheetName || '';
      const editedRange = entry.editedRange || '';
      const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;
      const logMessage = `Sheet: ${sheetName} | Range: ${editedRange}`;
    
      return [currentDate, sheetUrl, logMessage];
    });


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
