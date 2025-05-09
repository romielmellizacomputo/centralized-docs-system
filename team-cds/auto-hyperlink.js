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

async function getSheetIds(sheets) {
  const { data } = await sheets.spreadsheets.values.get({
    spreadsheetId: '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k',
    range: 'UTILS!B2:B',
  });
  return data.values?.flat().filter(Boolean) || [];
}

async function convertTitlesToHyperlinks(sheet, lastRow, titleColumn, projectColumn, iidColumn, urlType) {
  const titles = sheet.getRange(`${titleColumn}4:${titleColumn}${lastRow}`).getValues();
  const projectNames = sheet.getRange(`${projectColumn}4:${projectColumn}${lastRow}`).getValues();
  const iids = sheet.getRange(`${iidColumn}4:${iidColumn}${lastRow}`).getValues();

  const hyperlinks = [];

  for (let i = 0; i < titles.length; i++) {
    const title = titles[i][0];
    const projectName = projectNames[i][0];
    const iid = iids[i][0];

    if (title && iid && !title.includes("http")) { // Only process if title is not already a link
      const projectId = PROJECT_NAME_ID_MAP[projectName];
      if (projectId && PROJECT_URLS_MAP[projectId]) {
        const hyperlink = `${PROJECT_URLS_MAP[projectId]}/-/${urlType}/${iid}`;
        const escapedTitle = title.replace(/"/g, '""');
        hyperlinks.push([`=HYPERLINK("${hyperlink}", "${escapedTitle}")`]);
      } else {
        hyperlinks.push([title]); // If no matching project URL, keep the title as is
      }
    } else {
      hyperlinks.push([title]); // Skip rows with no title or IID
    }
  }

  if (hyperlinks.length > 0) {
    sheet.getRange(4, sheet.getRange(`${titleColumn}1`).getColumn(), hyperlinks.length, 1).setValues(hyperlinks);
  }
}

async function processSheets(sheets) {
  const sheetIds = await getSheetIds(sheets);
  const sheetsToProcess = sheetIds.map(async (sheetId) => {
    try {
      const sheet = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId });
      const lastRow = sheet.data.values.length;
      if (lastRow >= 4) {
        await convertTitlesToHyperlinks(sheet, lastRow, 'E', 'B', 'C', 'issues'); // E: title, B: project, C: iid
        console.log(`✅ Processed sheet: ${sheetId}`);
      }
    } catch (err) {
      console.error(`❌ Error processing sheet ${sheetId}: ${err.message}`);
    }
  });

  await Promise.all(sheetsToProcess); // Run in parallel
}

async function main() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });

    await processSheets(sheets); // Process all sheets in parallel

    console.log('✅ All sheets processed successfully');
  } catch (err) {
    console.error(`❌ Main failure: ${err.message}`);
  }
}

main();
