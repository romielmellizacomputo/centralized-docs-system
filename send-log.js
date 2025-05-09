const { google } = require('googleapis');

const targetSpreadsheetId = process.env.AUTOMATED_PORTALS; // Target sheet where "Logs" sheet is
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON); // Source service account

async function sendUpdateSignal() {
  try {
    console.log('📤 Starting signal send to Logs');

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const currentDate = new Date().toISOString();
    const sourceUrl = `https://docs.google.com/spreadsheets/d/${credentials.spreadsheet_id}`;

    // Append log entry to Logs sheet in the target spreadsheet
    await sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: 'Logs!A:B',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[currentDate, sourceUrl]],
      },
    });

    console.log(`✅ Log added to Logs sheet in ${targetSpreadsheetId}`);
  } catch (error) {
    console.error('❌ Error sending update signal:', error.message);
  }
}

sendUpdateSignal();
