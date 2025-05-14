import { google } from 'googleapis';

const sheetData = JSON.parse(process.env.SHEET_DATA); 
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON); 
const targetSpreadsheetId = process.env.AUTOMATED_PORTALS; 

async function sendUpdateSignal() {
  try {
    console.log('📤 Starting signal send to Logs sheet');

    if (!targetSpreadsheetId) {
      throw new Error('Missing required environment variable: AUTOMATED_PORTALS');
    }

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const updates = Array.isArray(sheetData) ? sheetData : [sheetData];

    const logEntries = [];

    for (const entry of updates) {
      const { spreadsheetUrl, sheetName, editedRange } = entry;

      if (!spreadsheetUrl || !sheetName) {
        throw new Error(`Missing spreadsheetUrl or sheetName in entry: ${JSON.stringify(entry)}`);
      }

      const spreadsheetIdMatch = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!spreadsheetIdMatch) {
        throw new Error(`Invalid spreadsheet URL: ${spreadsheetUrl}`);
      }

      const spreadsheetId = spreadsheetIdMatch[1];

      // Get spreadsheet metadata to find the gid (sheetId)
      const metadata = await sheets.spreadsheets.get({ spreadsheetId });
      const matchingSheet = metadata.data.sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );

      if (!matchingSheet) {
        throw new Error(`Sheet name "${sheetName}" not found in spreadsheet: ${spreadsheetUrl}`);
      }

      const gid = matchingSheet.properties.sheetId;
      const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit?gid=${gid}#gid=${gid}`;
      const currentDate = new Date().toISOString();
      const logMessage = `Sheet: ${sheetName} | Range: ${editedRange || 'N/A'}`;

      logEntries.push([currentDate, sheetUrl, logMessage]);
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: `Logs!A:C`,
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
