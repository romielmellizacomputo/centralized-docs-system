const { google } = require('googleapis');

const sheetData = JSON.parse(process.env.SHEET_DATA); // Includes source metadata (like sourceUrl or sheetId)
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON); // Source sheet credentials
const targetSpreadsheetId = process.env.AUTOMATED_PORTALS; // Target sheet (Logs destination)

async function sendUpdateSignal() {
  try {
    console.log('📤 Starting signal send to Logs sheet');

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const currentDate = new Date().toISOString();
    const sourceUrl = sheetData.sourceUrl || `https://docs.google.com/spreadsheets/d/${sheetData.spreadsheetId}`;

    await sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: 'Logs!A:B',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[currentDate, sourceUrl]],
      },
    });

    console.log(`✅ Log entry added to Logs sheet in target spreadsheet (${targetSpreadsheetId})`);
  } catch (error) {
    console.error('❌ Error sending log signal:', error.message);
    process.exit(1);
  }
}

sendUpdateSignal();
