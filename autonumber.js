const { google } = require('googleapis');

async function main() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const spreadsheetUrl = process.env.SHEET_URL;

  console.log("Received Spreadsheet URL:", spreadsheetUrl);
  if (!spreadsheetUrl) {
    console.error("Spreadsheet URL is not defined!");
    process.exit(1);
  }

  const match = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    console.error("Invalid Spreadsheet URL format.");
    process.exit(1);
  }

  const spreadsheetId = match[1];

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = metadata.data.sheets.map(s => s.properties.title);
    const skip = ['ToC', 'Roster', 'Issues'];

    for (const name of sheetNames) {
      if (skip.includes(name)) continue;

      const range = `'${name}'!E12:F`;
      const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
      const rows = res.data.values || [];

      let num = 1;
      const values = rows.map(([_, f]) => [(f || '').trim() ? num++ : '']);

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${name}'!E12:E${12 + values.length - 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values }
      });

      console.log(`Updated sheet: ${name}`);
    }
  } catch (err) {
    console.error('ERROR:', err);
    process.exit(1);
  }
}

main();
