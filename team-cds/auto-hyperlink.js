import { google } from 'googleapis';
import 'dotenv/config';

const GITLAB_ROOT_URL = 'https://forge.bposeats.com/';
const PROJECT_URLS_MAP = {
  155: `${GITLAB_ROOT_URL}bposeats/hqzen.com`,
  23: `${GITLAB_ROOT_URL}bposeats/bposeats`,
  124: `${GITLAB_ROOT_URL}bposeats/android-app`,
  123: `${GITLAB_ROOT_URL}bposeats/bposeats-desktop`,
  88: `${GITLAB_ROOT_URL}bposeats/applybpo.com`,
  141: `${GITLAB_ROOT_URL}bposeats/ministry-vuejs`,
  147: `${GITLAB_ROOT_URL}bposeats/scalema.com`,
  89: `${GITLAB_ROOT_URL}bposeats/bposeats.com`
};

const PROJECT_NAME_ID_MAP = {
  'HQZen': 155,
  'Backend': 23,
  'Android': 124,
  'Desktop': 123,
  'ApplyBPO': 88,
  'Ministry': 141,
  'Scalema': 147,
  'BPOSeats.com': 89
};

async function authenticate() {
  try {
    const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON || '');
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    return auth;
  } catch (err) {
    throw new Error('❌ Failed to parse TEAM_CDS_SERVICE_ACCOUNT_JSON: ' + err.message);
  }
}

async function getSheetData(sheets, sheetId, range) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: range
  });
  return data.values;
}

async function getSheetTitles(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets.map(sheet => sheet.properties.title);
}

function buildHyperlink(title, projectName, iid) {
  const projectId = PROJECT_NAME_ID_MAP[projectName];
  const baseUrl = PROJECT_URLS_MAP[projectId];
  if (title && iid && projectId && baseUrl && !title.includes('http')) {
    const hyperlink = `${baseUrl}/-/merge_requests/${iid}`;
    const escapedTitle = title.replace(/"/g, '""');
    return `=HYPERLINK("${hyperlink}", "${escapedTitle}")`;
  }
  return title;
}

async function convertTitlesToHyperlinks(sheetsApi, sheetId, sheetName) {
  const titleRange = `${sheetName}!E4:E`;
  const idRange = `${sheetName}!C4:C`;
  const projectRange = `${sheetName}!N4:N`;

  const [titles = [], ids = [], projects = []] = await Promise.all([
    getSheetData(sheetsApi, sheetId, titleRange),
    getSheetData(sheetsApi, sheetId, idRange),
    getSheetData(sheetsApi, sheetId, projectRange)
  ]);

  const hyperlinks = titles.map((row, i) => {
    const title = row[0];
    const iid = ids[i]?.[0];
    const project = projects[i]?.[0];
    return [buildHyperlink(title, project, iid)];
  });

  if (hyperlinks.length > 0) {
    await sheetsApi.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: titleRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: hyperlinks }
    });
    console.log(`🔗 Updated hyperlinks in ${sheetName}`);
  }
}

async function processSheets(sheetsApi, sheetId) {
  const sheetTitles = await getSheetTitles(sheetsApi, sheetId);
  const relevantSheets = ['NTC', 'G-Issues', 'G-MR'];

  for (const sheetName of relevantSheets) {
    if (sheetTitles.includes(sheetName)) {
      await convertTitlesToHyperlinks(sheetsApi, sheetId, sheetName);
    } else {
      console.warn(`⚠️ Sheet "${sheetName}" not found in ${sheetId}`);
    }
  }
}

async function main() {
  try {
    const auth = await authenticate();
    const sheetsApi = google.sheets({ version: 'v4', auth });

    const utilsSheetId = process.env.UTILS_SHEET_ID;
    const sheetIds = await getSheetData(sheetsApi, utilsSheetId, 'UTILS!B2:B');

    if (!sheetIds.length) {
      console.error('❌ No sheet IDs found in UTILS!B2:B');
      return;
    }

    for (const [sheetId] of sheetIds) {
      if (!sheetId) continue;
      console.log(`📄 Processing: ${sheetId}`);
      await processSheets(sheetsApi, sheetId);
      console.log(`✅ Finished processing: ${sheetId}`);
    }
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
