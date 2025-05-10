import { google } from 'googleapis';

const sheetData = JSON.parse(process.env.SHEET_DATA); // includes source info
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON); // service account with access
const targetSpreadsheetId = process.env.AUTOMATED_PORTALS; // destination sheet ID where Logs tab is

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
    const sourceUrl = sheetData.spreadsheetUrl;
    const sheetName = sheetData.sheetName || '';
    const editedRange = sheetData.editedRange || '';

    const logMessage = `Sheet: ${sheetName} | Range: ${editedRange}`;

    await sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: `${targetSheetName}!A:C`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[currentDate, sourceUrl, logMessage]],
      },
    });

    console.log(`✅ Log added to Logs sheet at ${targetSpreadsheetId}`);
  } catch (error) {
    console.error('❌ Error sending log signal:', error.message);
  }
}

sendUpdateSignal();
