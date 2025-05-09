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

function convertTitlesToHyperlinksAcrossSheets() {
  const masterSheet = SpreadsheetApp.openById('1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k');
  const utilsSheet = masterSheet.getSheetByName('UTILS');
  const urls = utilsSheet.getRange('B2:B').getValues().flat().filter(url => url);

  urls.forEach(url => {
    try {
      const doc = SpreadsheetApp.openByUrl(url);
      processSpreadsheet(doc);
    } catch (e) {
      Logger.log(`Failed to open or process: ${url}, error: ${e}`);
    }
  });
}

function processSpreadsheet(spreadsheet) {
  sheetsToProcess.forEach(sheetInfo => {
    const sheet = spreadsheet.getSheetByName(sheetInfo.name);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 4) {
        convertSheetTitlesToHyperlinks(sheet, lastRow, sheetInfo.projectColumn, sheetInfo.iidColumn, sheetInfo.titleColumn, sheetInfo.urlType);
      }
    }
  });
}

function convertSheetTitlesToHyperlinks(sheet, lastRow, projectColumn, iidColumn, titleColumn, urlType) {
  const titles = sheet.getRange(`${titleColumn}4:${titleColumn}${lastRow}`).getValues();
  const projectNames = sheet.getRange(`${projectColumn}4:${projectColumn}${lastRow}`).getValues();
  const iids = sheet.getRange(`${iidColumn}4:${iidColumn}${lastRow}`).getValues();

  const hyperlinks = [];

  for (let i = 0; i < titles.length; i++) {
    const title = titles[i][0];
    const projectName = projectNames[i][0];
    const iid = iids[i][0];

    // Ensure the title is not empty, does not contain an existing URL, and has a valid iid
    if (title && iid && !title.includes("http")) {
      const projectId = PROJECT_NAME_ID_MAP[projectName];
      if (projectId && PROJECT_URLS_MAP[projectId]) {
        const hyperlink = `${PROJECT_URLS_MAP[projectId]}/-/${urlType}/${iid}`;
        const escapedTitle = title.replace(/"/g, '""'); // Escape the title text properly
        hyperlinks.push([`=HYPERLINK("${hyperlink}", "${escapedTitle}")`]);
      } else {
        hyperlinks.push([title]); // If no valid project ID, just keep the title as plain text
      }
    } else {
      hyperlinks.push([title]); // If already contains a URL or no valid data, keep as is
    }
  }

  // Update the sheet with the hyperlinks
  if (hyperlinks.length > 0) {
    sheet.getRange(4, sheet.getRange(`${titleColumn}1`).getColumn(), hyperlinks.length, 1).setValues(hyperlinks);
  }
}


// Placeholder for a function that fetches sheet URLs from a master sheet
function getSheetUrls() {
  return [
    'NTC',
    'G-Issues',
    'G-MR'
  ];
}

// Call the main function to initiate the process
convertTitlesToHyperlinks();
