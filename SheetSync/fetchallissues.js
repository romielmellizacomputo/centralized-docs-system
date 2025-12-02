import { config } from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import { GITLAB_URL, GITLAB_TOKEN, SHEET_SYNC_SID } from '../constants.js';

config();

const requiredEnv = ['GITLAB_URL', 'GITLAB_TOKEN', 'SHEET_SYNC_SID', 'SHEET_SYNC_SAJ'];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

let PROJECT_CONFIG;
try {
  PROJECT_CONFIG = JSON.parse(process.env.PROJECT_CONFIG);
} catch (err) {
  console.error('❌ Failed to parse PROJECT_CONFIG JSON from environment variable:', err);
  process.exit(1);
}

function loadServiceAccount() {
  if (process.env.GITHUB_ACTIONS && process.env.SHEET_SYNC_SAJ) {
    try {
      return JSON.parse(process.env.SHEET_SYNC_SAJ);
    } catch (error) {
      console.error('❌ Error parsing service account JSON:', error.message);
      throw error;
    }
  } else {
    console.error('❌ Script must run in GitHub Actions with SHEET_SYNC_SAJ');
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
  const date = new Date(dateString);
  return new Intl.DateTimeFormat('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  }).format(date);
}

async function validateSheetExistsAndHasPermissions(sheets, spreadsheetId, sheetName) {
  console.log(`🔍 Validating sheet "${sheetName}" exists...`);
  
  try {
    // Check if spreadsheet and sheet exist
    const meta = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: false,
    });
    
    const sheet = meta.data.sheets.find((s) => s.properties.title === sheetName);
    if (!sheet) {
      throw new Error(`Sheet with name "${sheetName}" not found in spreadsheet`);
    }
    
    console.log(`✅ Sheet "${sheetName}" exists`);
    
    // Test write permission by writing test data
    console.log(`🔍 Testing write permissions...`);
    const testRange = `${sheetName}!A4`;
    const testValue = `TEST_${Date.now()}`;
    
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: testRange,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[testValue]],
      },
    });
    
    console.log(`✅ Write permission confirmed`);
    
    // Verify the test data was written
    console.log(`🔍 Verifying test data...`);
    const readResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: testRange,
    });
    
    if (readResponse.data.values && readResponse.data.values[0][0] === testValue) {
      console.log(`✅ Read permission confirmed`);
    } else {
      throw new Error('Failed to verify test data read');
    }
    
    // Clean up test data
    console.log(`🧹 Cleaning up test data...`);
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: testRange,
    });
    
    console.log(`✅ Test data cleaned up successfully`);
    console.log(`✅ All validations passed for sheet "${sheetName}"`);
    
    return sheet.properties.sheetId;
    
  } catch (error) {
    console.error(`❌ Sheet validation failed: ${error.message}`);
    if (error.code === 403) {
      console.error(`❌ Permission denied. Make sure the service account has edit access to the spreadsheet.`);
    } else if (error.code === 404) {
      console.error(`❌ Spreadsheet or sheet not found. Check SHEET_SYNC_SID and sheet name.`);
    }
    throw error;
  }
}

async function fetchIssuesForProject(projectId, config) {
  let page = 1;
  let issues = [];
  console.log(`🔄 Fetching issues for project ID ${projectId} (${config.name})...`);

  while (true) {
    const response = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/issues?state=all&per_page=100&page=${page}`,
      {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      }
    );

    if (response.status !== 200) {
      console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
      break;
    }

    const fetchedIssues = response.data;
    if (fetchedIssues.length === 0) break;

    fetchedIssues.forEach((issue) => {
      const issueData = [
        issue.id ?? '',
        issue.iid ?? '',
        issue.title && issue.web_url
          ? `=HYPERLINK("${issue.web_url}", "${issue.title.replace(/"/g, '""')}")`
          : 'No Title',
        issue.author?.name ?? 'Unknown Author',
        issue.assignee?.name ?? 'Unassigned',
        (issue.labels || []).join(', '),
        issue.milestone?.title ?? 'No Milestone',
        capitalize(issue.state ?? ''),
        issue.created_at ? formatDate(issue.created_at) : '',
        issue.closed_at ? formatDate(issue.closed_at) : '',
        issue.closed_by?.name ?? '',
        config.name,
      ];

      issues.push(issueData);
    });

    console.log(`✅ Page ${page} fetched (${fetchedIssues.length} issues) for ${config.name}`);
    page++;
  }

  return issues;
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

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  // ✅ VALIDATE SHEET FIRST - before fetching any GitLab data
  console.log('🔍 Starting sheet validation...');
  try {
    await validateSheetExistsAndHasPermissions(sheets, SHEET_SYNC_SID, 'ALL ISSUES');
  } catch (error) {
    console.error('❌ Cannot proceed: Sheet validation failed');
    console.error('❌ No GitLab data will be fetched to avoid wasted API calls');
    process.exit(1);
  }

  console.log('🔄 Fetching issues for all projects...');

  // Fetch all new data after validation passes
  const issuesPromises = Object.keys(PROJECT_CONFIG).map(async (key) => {
    const config = PROJECT_CONFIG[key];
    config.name = key;
    const projectId = config.id;
    return fetchIssuesForProject(projectId, config);
  });

  const allIssuesResults = await Promise.all(issuesPromises);
  const allIssues = allIssuesResults.flat();

  if (allIssues.length === 0) {
    console.log('ℹ️ No issues to insert or update.');
    return;
  }

  const readRange = 'ALL ISSUES!C4:N';

  const existingResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_SYNC_SID,
    range: readRange,
  });

  const existingValues = existingResponse.data.values || [];

  const existingMap = new Map();
  existingValues.forEach((row, i) => {
    const id = row[0] ? String(row[0]) : '';
    const iid = row[1] ? String(row[1]) : '';
    const project = row[11] ? String(row[11]) : '';
    const key = `${id}|${iid}|${project}`;
    existingMap.set(key, i);
  });

  const updates = [];
  const inserts = [];

  allIssues.forEach((issueData) => {
    const id = String(issueData[0] ?? '');
    const iid = String(issueData[1] ?? '');
    const project = String(issueData[11] ?? '');
    const key = `${id}|${iid}|${project}`;

    if (existingMap.has(key)) {
      const zeroBasedIndex = existingMap.get(key);
      const sheetRowNumber = zeroBasedIndex + 4;
      updates.push({ row: sheetRowNumber, values: issueData });
    } else {
      inserts.push(issueData);
    }
  });

  const batchUpdateRequests = updates.map((update) => {
    return {
      range: `ALL ISSUES!C${update.row}:N${update.row}`,
      values: [update.values],
    };
  });

  try {
    if (batchUpdateRequests.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_SYNC_SID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: batchUpdateRequests,
        },
      });

      console.log(`✅ Updated ${batchUpdateRequests.length} existing issues.`);
    } else {
      console.log('ℹ️ No existing issues to update.');
    }

    if (inserts.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_SYNC_SID,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: await getSheetIdByName(sheets, SHEET_SYNC_SID, 'ALL ISSUES'),
                  dimension: 'ROWS',
                  startIndex: 3,
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
        range: `ALL ISSUES!C4`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: inserts,
        },
      });

      console.log(`✅ Inserted ${inserts.length} new issues.`);
    } else {
      console.log('ℹ️ No new issues to insert.');
    }
  } catch (err) {
    console.error('❌ Error updating/inserting rows:', err.stack || err.message);
  }
}

fetchAndUpdateIssuesForAllProjects();
