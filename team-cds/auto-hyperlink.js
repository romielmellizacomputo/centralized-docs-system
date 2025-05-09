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
  const credentials = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

async function getSheetTitles(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  const titles = res.data.sheets.map(sheet => sheet.properties.title);
  console.log(`📄 Sheets in ${spreadsheetId}:`, titles);
  return titles;
}

async function getSheetData(sheets, sheetId, range) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: range,
  });
  return data.values;
}

async function convertToHyperlinks(sheet, lastRow, titleColumn, projectColumn, iidColumn) {
  const titles = sheet.getRange(`${titleColumn}4:${titleColumn}${lastRow}`).getValues();
  const projectNames = sheet.getRange(`${projectColumn}4:${projectColumn}${lastRow}`).getValues();
  const iids = sheet.getRange(`${iidColumn}4:${iidColumn}${lastRow}`).getValues();

  const hyperlinks = [];

  for (let i = 0; i < titles.length; i++) {
    const title = titles[i][0];
    const projectName = projectNames[i][0];
    const iid = iids[i][0];

    if (title && iid && !title.includes("http")) { // Check if title already contains a hyperlink
      const projectId = PROJECT_NAME_ID_MAP[projectName];
      if (projectId && PROJECT_URLS_MAP[projectId]) {
        const hyperlink = `${PROJECT_URLS_MAP[projectId]}/-/issues/${iid}`; // Updated URL structure
        // Escape double quotes in the title to avoid formula parse errors
        const escapedTitle = title.replace(/"/g, '""');
        hyperlinks.push([`=HYPERLINK("${hyperlink}", "${escapedTitle}")`]);
      } else {
        hyperlinks.push([title]); // Add title as is if no URL
      }
    } else {
      hyperlinks.push([title]); // Skip rows with no title or IID, or already linked
    }
  }

  if (hyperlinks.length > 0) {
    sheet.getRange(4, sheet.getRange(`${titleColumn}1`).getColumn(), hyperlinks.length, 1).setValues(hyperlinks);
  }
}

async function processSheets(sheets) {
  // Define the sheets we want to process
  const sheetsToProcess = ['NTC', 'G-Issues', 'G-MR'];
  
  for (const sheetName of sheetsToProcess) {
    const sheet = sheets.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 4) {
        // Process column E4:E (hyperlinking based on project and iid)
        await convertToHyperlinks(sheet, lastRow, 'E', 'N', 'C');
      }
    }
  }
}

async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    // Get list of Google Sheet IDs from UTILS!B2:B
    const sheetIds = await getSheetData(sheets, 'UTILS_SHEET_ID', 'UTILS!B2:B');
    if (!sheetIds.length) {
      console.error('❌ No Google Sheets found in UTILS!B2:B');
      return;
    }

    // Iterate over each sheet ID, process the relevant sheets (NTC, G-Issues, G-MR)
    for (const sheetId of sheetIds) {
      console.log(`🔄 Processing: ${sheetId}`);
      
      // Get sheet titles for this sheet ID
      const sheetTitles = await getSheetTitles(sheets, sheetId);

      // Check if the required sheets exist
      if (sheetTitles.includes('NTC') && sheetTitles.includes('G-Issues') && sheetTitles.includes('G-MR')) {
        await processSheets(sheets, sheetId);
        console.log(`✅ Finished processing sheets for ${sheetId}`);
      } else {
        console.warn(`⚠️ Skipping ${sheetId} — missing one or more required sheets.`);
      }
    }
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
