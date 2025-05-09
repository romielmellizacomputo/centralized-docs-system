import { google } from 'googleapis';

const TEAM_CDS_SERVICE_ACCOUNT_JSON = process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON;
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

const sheetsToProcess = [
  { name: 'NTC', projectColumn: 'N', iidColumn: 'D', titleColumn: 'E', urlType: 'issues' },
  { name: 'G-Issues', projectColumn: 'N', iidColumn: 'D', titleColumn: 'E', urlType: 'issues' },
  { name: 'G-MR', projectColumn: 'O', iidColumn: 'D', titleColumn: 'E', urlType: 'merge_requests' }
];

async function getAuthClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(TEAM_CDS_SERVICE_ACCOUNT_JSON),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  const authClient = await auth.getClient();
  return authClient;
}

async function convertTitlesToHyperlinks() {
  const masterSheetUrl = 'https://docs.google.com/spreadsheets/d/1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k/edit?gid=1536197668#gid=1536197668';

  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const masterSheetId = masterSheetUrl.split('/d/')[1]?.split('/')[0];
    if (!masterSheetId) {
      console.error('Invalid master sheet URL.');
      return;
    }
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: 'UTILS!B2:B', // Replace with actual range if needed
    });

    const urls = res.data.values?.flat()?.filter(url => url);

    if (!urls || urls.length === 0) {
      console.error('No valid URLs found.');
      return;
    }

    for (let sheetId of urls) {
      if (!sheetId) {
        console.log('Skipping invalid or empty sheet ID');
        continue;
      }

      try {
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
        const doc = await sheets.spreadsheets.values.get({
          spreadsheetId: sheetId,
          range: 'Sheet1', // Fetch the entire sheet
        });

        if (!doc.data.values || doc.data.values.length === 0) {
          console.error(`No data found for sheet ID ${sheetId}`);
          continue;
        }

        processSpreadsheet(doc.data.values, sheetId);
      } catch (e) {
        console.error(`Failed to open or process sheet with ID: ${sheetId}, error: ${e}`);
      }
    }
  } catch (e) {
    console.error(`Failed to open master sheet: ${e}`);
  }
}

function processSpreadsheet(sheetData, sheetId) {
  sheetsToProcess.forEach(sheetInfo => {
    const sheet = sheetData.find(sheet => sheet[0] === sheetInfo.name);
    if (sheet) {
      const lastRow = sheet.length;
      if (lastRow >= 4) {
        convertSheetTitlesToHyperlinks(sheet, lastRow, sheetInfo.projectColumn, sheetInfo.iidColumn, sheetInfo.titleColumn, sheetInfo.urlType, sheetId);
      }
    }
  });
}

function convertSheetTitlesToHyperlinks(sheet, lastRow, projectColumn, iidColumn, titleColumn, urlType, sheetId) {
  const projectColumnIndex = projectColumn.charCodeAt(0) - 65;
  const iidColumnIndex = iidColumn.charCodeAt(0) - 65;
  const titleColumnIndex = titleColumn.charCodeAt(0) - 65;

  const hyperlinks = [];

  // Start from row 4 (index 3) and process only non-empty column E
  for (let i = 3; i < lastRow; i++) {
    const title = sheet[i][titleColumnIndex];
    const projectName = sheet[i][projectColumnIndex];
    const iid = sheet[i][iidColumnIndex];

    if (title && iid && !title.includes("http")) {
      const projectId = PROJECT_NAME_ID_MAP[projectName];
      if (projectId && PROJECT_URLS_MAP[projectId]) {
        const hyperlink = `${PROJECT_URLS_MAP[projectId]}/-/${urlType}/${iid}`;
        const escapedTitle = title.replace(/"/g, '""');
        hyperlinks.push([`=HYPERLINK("${hyperlink}", "${escapedTitle}")`]);
      } else {
        hyperlinks.push([title]); 
      }
    } else {
      hyperlinks.push([title]);
    }
  }

  console.log(`Processing sheet ID ${sheetId} with hyperlinks:`, hyperlinks);

  // Update sheet with hyperlinks (adjust per sheet setup)
  // e.g., sheet.getRange(4, 5, hyperlinks.length, 1).setValues(hyperlinks); 
}

convertTitlesToHyperlinks();
