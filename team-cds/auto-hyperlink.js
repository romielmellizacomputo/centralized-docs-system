import { google } from 'googleapis'; // Import googleapis
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import * as dotenv from 'dotenv';

// Load environment variables from .env file (if you use it)
dotenv.config();

// Get the directory name equivalent of __dirname in ES module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Save the secret service account JSON to a file
const serviceAccountJson = process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON;
if (!serviceAccountJson) {
  console.error('SERVICE_ACCOUNT_JSON secret not found!');
  process.exit(1);
}

const serviceAccountPath = path.join(__dirname, 'service-account.json');
fs.writeFileSync(serviceAccountPath, serviceAccountJson);

// Google Sheets API authentication using the service account
async function getAuthClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: serviceAccountPath, // Path to the service account JSON file
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return auth.getClient();
}

// URLs of the projects in GitLab
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

// Mapping project names to their respective IDs
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

// List of sheets to process
const sheetsToProcess = [
  { name: 'NTC', projectColumn: 'N', iidColumn: 'D', titleColumn: 'E', urlType: 'issues' },
  { name: 'G-Issues', projectColumn: 'N', iidColumn: 'D', titleColumn: 'E', urlType: 'issues' },
  { name: 'G-MR', projectColumn: 'O', iidColumn: 'D', titleColumn: 'E', urlType: 'merge_requests' }
];

// Function to convert titles to hyperlinks across sheets
async function convertTitlesToHyperlinks() {
  // Replace with the actual URL of the master sheet
  const masterSheetUrl = 'https://docs.google.com/spreadsheets/d/1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k/edit?gid=1536197668#gid=1536197668';
  
  const auth = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  try {
    const masterSheetId = masterSheetUrl.split('/d/')[1]?.split('/')[0]; // Safe splitting
    if (!masterSheetId) {
      console.error('Invalid master sheet URL.');
      return;
    }
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: 'UTILS!B2:B',
    });

    const urls = res.data.values?.flat()?.filter(url => url); // Ensure no empty URLs

    // Check if the URLs array is valid and contains data
    if (!urls || urls.length === 0) {
      console.error('No valid URLs found.');
      return;
    }

    // Process each URL in the sheet
    for (let url of urls) {
      if (!url) {
        console.log('Skipping invalid or empty URL');
        continue;
      }

      try {
        const spreadsheetId = url.split('/d/')[1]?.split('/')[0]; // Safe splitting
        if (!spreadsheetId) {
          console.log(`Invalid URL format: ${url}`);
          continue;
        }

        const doc = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: 'Sheet1!A1:Z1000', // Adjust this range to process your sheet
        });
        processSpreadsheet(doc.data.values);
      } catch (e) {
        console.log(`Failed to open or process: ${url}, error: ${e}`);
      }
    }
  } catch (e) {
    console.log(`Failed to open master sheet: ${e}`);
  }
}


// Process each spreadsheet
function processSpreadsheet(spreadsheetData) {
  sheetsToProcess.forEach(sheetInfo => {
    // You would likely want to pass in the sheet data you retrieve from each URL to process further
    convertSheetTitlesToHyperlinks(spreadsheetData, sheetInfo);
  });
}

// Convert sheet titles to hyperlinks
function convertSheetTitlesToHyperlinks(sheetData, sheetInfo) {
  const { projectColumn, iidColumn, titleColumn, urlType } = sheetInfo;
  
  const hyperlinks = sheetData.map((row, index) => {
    const title = row[titleColumn];
    const projectName = row[projectColumn];
    const iid = row[iidColumn];

    if (title && iid && !title.includes("http")) {
      const projectId = PROJECT_NAME_ID_MAP[projectName];
      if (projectId && PROJECT_URLS_MAP[projectId]) {
        const hyperlink = `${PROJECT_URLS_MAP[projectId]}/-/${urlType}/${iid}`;
        const escapedTitle = title.replace(/"/g, '""'); // Escape double quotes in title
        return [`=HYPERLINK("${hyperlink}", "${escapedTitle}")`]; // Create hyperlink
      } else {
        return [title]; // If no matching project URL, keep the title as plain text
      }
    }
    return [title]; // If no title or URL, keep the title as plain text
  });

  // Log the hyperlinks for now (you can use the Sheets API to update the sheet as needed)
  console.log(hyperlinks); // Replace this with a call to update the sheet using the Sheets API

  // Example: Update the sheet with hyperlinks (adjust this part to your actual sheet update logic)
  // const updateRange = `E4:E${sheetData.length}`;
  // await sheets.spreadsheets.values.update({
  //   spreadsheetId: YOUR_SPREADSHEET_ID,
  //   range: updateRange,
  //   valueInputOption: 'RAW',
  //   requestBody: {
  //     values: hyperlinks,
  //   },
  // });
}

// Call the main function to process everything
convertTitlesToHyperlinks();
