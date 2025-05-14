import { google } from 'googleapis';
import axios from 'axios';

const sheetData = JSON.parse(process.env.SHEET_DATA);
const credentials = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);

async function updateToC() {
  const spreadsheetUrl = sheetData.spreadsheetUrl;
  const spreadsheetId = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  const skip = ['ToC', 'Issues', 'Roster'];

  try {
    const metadata = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheets = metadata.data.sheets || [];
    const tocSheet = allSheets.find(s => s.properties.title === 'ToC');

    if (!tocSheet) {
      console.error("ToC sheet not found.");
      return;
    }

    const tocTitle = tocSheet.properties.title;

    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${tocTitle}!A2:A`,
    });

    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${tocTitle}!B2:K`,
    });

    let tocRows = [];
    for (const sheet of allSheets) {
      const name = sheet.properties.title;
      if (skip.includes(name)) continue;

      // Read value in C4
      const c4Range = `'${name}'!C4`;
      const c4Res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: c4Range,
      });

      const c4Value = c4Res.data.values?.[0]?.[0];
      if (!c4Value) continue;

      if (tocRows.some(row => row[0].includes(c4Value))) continue;

      const sheetId = sheet.properties.sheetId;
      const hyperlink = `=HYPERLINK("${spreadsheetUrl}#gid=${sheetId}", "${c4Value}")`;

      const cellsToRead = ['C5', 'C7', 'C15', 'C18', 'C19', 'C20', 'C21', 'C14', 'C13', 'C6'];
      const batchRanges = cellsToRead.map(cell => `'${name}'!${cell}`);

      const dataRes = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges: batchRanges,
      });

      const values = dataRes.data.valueRanges.map(r => r.values?.[0]?.[0] || '');

      tocRows.push([hyperlink, ...values]);
      console.log(`Inserted hyperlink for: ${c4Value}`);
    }

    if (tocRows.length > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${tocTitle}!A2:K${tocRows.length + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: tocRows,
        },
      });
    }

    try {
      const webAppUrl = 'https://script.google.com/macros/s/AKfycbzR3hWvfItvEOKjadlrVRx5vNTz4QH04WZbz2ufL8fAdbiZVsJbkzueKfmMCfGsAO62/exec';

      await axios.post(webAppUrl, {
        sheetUrl: spreadsheetUrl,
      }, {
        headers: {
          'Content-Type': 'application/json',
        },
      });

      console.log("POST request sent to web app.");
    } catch (err) {
      console.error("Error sending POST request:", err.message);
    }
  } catch (err) {
    console.error("ERROR:", err.message);
    process.exit(1);
  }
}

updateToC();
