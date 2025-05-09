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

// Example function to simulate fetching sheet data
function fetchSheetData(sheetName) {
  // This is just a placeholder to represent your sheet data. You would need to adapt
  // it to your real data-fetching mechanism (e.g., via API or local JSON).
  return [
    { project: 'Backend', iid: 1, title: 'Issue 1' },
    { project: 'Android', iid: 2, title: 'Issue 2' }
  ];
}

function convertTitlesToHyperlinks() {
  const sheetUrls = getSheetUrls();  // Get URLs of sheets to process
  sheetUrls.forEach(url => {
    try {
      const sheetData = fetchSheetData(url);
      processSpreadsheet(sheetData);
    } catch (e) {
      console.log(`Failed to open or process: ${url}, error: ${e}`);
    }
  });
}

function processSpreadsheet(sheetData) {
  sheetsToProcess.forEach(sheetInfo => {
    const data = fetchSheetData(sheetInfo.name); // Get data for each sheet
    if (data && data.length > 0) {
      convertSheetTitlesToHyperlinks(data, sheetInfo.projectColumn, sheetInfo.iidColumn, sheetInfo.titleColumn, sheetInfo.urlType);
    }
  });
}

function convertSheetTitlesToHyperlinks(data, projectColumn, iidColumn, titleColumn, urlType) {
  const hyperlinks = [];

  data.forEach(row => {
    const { title, project, iid } = row;

    if (title && iid && !title.includes("http")) {
      const projectId = PROJECT_NAME_ID_MAP[project];
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
  });

  // Example output to log the hyperlinks
  console.log(hyperlinks);
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
