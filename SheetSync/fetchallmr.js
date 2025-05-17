import { config } from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';

config();

const requiredEnv = ['GITLAB_URL', 'GITLAB_TOKEN', 'SHEET_SYNC_SID', 'SHEET_SYNC_SAJ'];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

const GITLAB_URL = process.env.GITLAB_URL.endsWith('/') ? process.env.GITLAB_URL : process.env.GITLAB_URL + '/';
const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
const SHEET_SYNC_SID = process.env.SHEET_SYNC_SID;

let PROJECT_CONFIG;
try {
  PROJECT_CONFIG = JSON.parse(process.env.PROJECT_CONFIG);
} catch (err) {
  console.error('❌ Failed to parse PROJECT_CONFIG JSON from environment variable:', err);
  process.exit(1);
}

function loadServiceAccount() {
  try {
    return JSON.parse(process.env.SHEET_SYNC_SAJ);
  } catch (err) {
    console.error('❌ Error parsing service account JSON:', err.message);
    process.exit(1);
  }
}

const serviceAccount = loadServiceAccount();

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function formatDate(dateString) {
  if (!dateString) return '';
  return new Intl.DateTimeFormat('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  }).format(new Date(dateString));
}


async function fetchMRsForProject(projectId, config) {
  let page = 1;
  let mrs = [];
  console.log(`🔄 Fetching MRs for ${config.name}...`);

  while (true) {
    const url = `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests?state=all&per_page=100&page=${page}`;
    let response;

    try {
      response = await axios.get(url, {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      });
    } catch (error) {
      console.error(`❌ Error fetching MRs for ${config.name} page ${page}:`, error.message);
      break; // Stop on error
    }

    if (response.status !== 200) {
      console.error(`❌ Failed to fetch MRs for ${config.name} (page ${page}), status: ${response.status}`);
      break;
    }

    const fetchedMRs = response.data;
    if (fetchedMRs.length === 0) {
      // No more pages
      break;
    }

    for (const mr of fetchedMRs) {
      const reviewers = (mr.reviewers || []).map(r => r.name).join(', ') || 'Unassigned';

      mrs.push([
        mr.id ?? '',
        mr.iid ?? '',
        mr.title && mr.web_url
          ? `=HYPERLINK("${mr.web_url}", "${mr.title.replace(/"/g, '""')}")`
          : 'No Title',
        mr.author?.name ?? 'Unknown Author',
        mr.assignee?.name ?? 'Unassigned',
        reviewers,
        (mr.labels || []).join(', '),
        mr.milestone?.title ?? 'No Milestone',
        capitalize(mr.state ?? ''),
        formatDate(mr.created_at),
        formatDate(mr.closed_at),
        formatDate(mr.merged_at),
        config.name,
      ]);
    }

    console.log(`✅ Page ${page} fetched (${fetchedMRs.length} MRs) for ${config.name}`);

    // Use x-next-page header to decide if continue or stop
    const nextPage = response.headers['x-next-page'];
    if (!nextPage) break;

    page = parseInt(nextPage, 10);
    if (isNaN(page) || page === 0) break;
  }

  return mrs;
}

async function fetchAndUpdateMRsForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  console.log('🔄 Fetching MRs for all projects...');

  const mrPromises = Object.entries(PROJECT_CONFIG).map(([projectId, config]) =>
    fetchMRsForProject(projectId, config)
  );

  const allMRsResults = await Promise.all(mrPromises);
  const allMRs = allMRsResults.flat();

  if (allMRs.length === 0) {
    console.log('ℹ️ No MRs to insert or update.');
    return;
  }

  const readRange = 'ALL MRs!C4:O';
  const existingResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_SYNC_SID,
    range: readRange,
  });

  const existingValues = existingResponse.data.values || [];
  const existingMap = new Map();
  existingValues.forEach((row, i) => {
    const id = row[0] ?? '';
    const iid = row[1] ?? '';
    const project = row[12] ?? '';
    const key = `${id}|${iid}|${project}`;
    existingMap.set(key, i);
  });

  const updates = [];
  const inserts = [];

  allMRs.forEach((mr) => {
    const id = String(mr[0] ?? '');
    const iid = String(mr[1] ?? '');
    const project = String(mr[12] ?? '');
    const key = `${id}|${iid}|${project}`;

    if (existingMap.has(key)) {
      const rowIndex = existingMap.get(key) + 4;
      updates.push({ row: rowIndex, values: mr });
    } else {
      inserts.push(mr);
    }
  });

  const batchUpdateRequests = updates.map(update => ({
    range: `ALL MRs!C${update.row}:O${update.row}`,
    values: [update.values],
  }));

  try {
    if (batchUpdateRequests.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_SYNC_SID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: batchUpdateRequests,
        },
      });
      console.log(`✅ Updated ${batchUpdateRequests.length} existing MRs.`);
    } else {
      console.log('ℹ️ No existing MRs to update.');
    }

    if (inserts.length > 0) {
      // Insert rows before updating values
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_SYNC_SID,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: await getSheetIdByName(sheets, SHEET_SYNC_SID, 'ALL MRs'),
                  dimension: 'ROWS',
                  startIndex: 3, // Index is zero-based, row 4 = index 3
                  endIndex: 3 + inserts.length,
                },
                inheritFromBefore: false,
              },
            },
          ],
        },
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_SYNC_SID,
        range: `ALL MRs!C4`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: inserts,
        },
      });

      console.log(`✅ Inserted ${inserts.length} new MRs.`);
    } else {
      console.log('ℹ️ No new MRs to insert.');
    }
    
  } catch (err) {
    console.error('❌ Error updating/inserting MRs:', err.stack || err.message);
  }
}

async function getSheetIdByName(sheets, spreadsheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    includeGridData: false,
  });
  const sheet = meta.data.sheets.find((s) => s.properties.title === sheetName);
  if (!sheet) {
    throw new Error(`Sheet with name "${sheetName}" not found`);
  }
  return sheet.properties.sheetId;
}

fetchAndUpdateMRsForAllProjects();
