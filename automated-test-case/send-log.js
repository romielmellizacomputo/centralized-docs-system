import { google } from 'googleapis';

const a = JSON.parse(process.env.SHEET_DATA); 
const b = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON); 
const c = process.env.AUTOMATED_PORTALS; 

async function d() {
  try {
    if (!c) {
      throw new Error('Missing required environment variable: AUTOMATED_PORTALS');
    }

    const e = new google.auth.GoogleAuth({
      credentials: b,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/spreadsheets'],
    });

    const f = google.sheets({ version: 'v4', auth: e });
    const g = Array.isArray(a) ? a : [a];

    const h = [];

    for (const i of g) {
      const { spreadsheetUrl: j, sheetName: k, editedRange: l } = i;

      if (!j || !k) {
        throw new Error(`Missing spreadsheetUrl or sheetName in entry: ${JSON.stringify(i)}`);
      }

      const m = j.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!m) {
        throw new Error(`Invalid spreadsheet URL: ${j}`);
      }

      const n = m[1];

      const o = await f.spreadsheets.get({ spreadsheetId: n });
      const p = o.data.sheets.find(
        (q) => q.properties.title === k
      );

      if (!p) {
        throw new Error(`Sheet name "${k}" not found in spreadsheet: ${j}`);
      }

      const r = p.properties.sheetId;
      const s = `https://docs.google.com/spreadsheets/d/${n}/edit?gid=${r}#gid=${r}`;
      const t = new Date().toISOString();
      const u = `Sheet: ${k} | Range: ${l || 'N/A'}`;

      h.push([t, s, u]);
    }

    await f.spreadsheets.values.append({
      spreadsheetId: c,
      range: `Logs!A:C`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: h,
      },
    });
  } catch (v) {
  }
}

d();
