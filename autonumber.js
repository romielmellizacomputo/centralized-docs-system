const { google } = require('googleapis');

async function main() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = extractSheetId(process.env.SHEET_URL);

  // Get all sheet names
  const metadata = await sheets.spreadsheets.get({ spreadsheetId });
  const allSheets = metadata.data.sheets.map(s => s.properties.title);
  const skipSheets = ['ToC', 'Roster', 'Issues'];

  for (const title of allSheets) {
    if (skipSheets.includes(title)) continue;

    const range = `'${title}'!E12:F`; // Read E12:F end of sheet
    const result = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const rows = result.data.values || [];

    const updates = [];
    let number = 1;

    for (let i = 0; i < rows.length; i++) {
      const [_, fValue] = rows[i];
      if (fValue && fValue.trim() !== '') {
        updates.push([number++]);
      } else {
        updates.push(['']);
      }
    }

    const updateRange = `'${title}'!E12:E${12 + updates.length - 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updateRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: updates }
    });

    console.log(`Updated sheet: ${title}`);
  }
}

function extractSheetId(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

main().catch(err => {
  console.error('Error running autonumber:', err);
  process.exit(1);
});
