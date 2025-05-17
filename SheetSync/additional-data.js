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

const GITLAB_URL = process.env.GITLAB_URL;
const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
const SHEET_SYNC_SID = process.env.SHEET_SYNC_SID;

const PROJECT_CONFIG = {
  155: { name: 'HQZen', path: 'bposeats/hqzen.com' },
  88: { name: 'ApplyBPO', path: 'bposeats/applybpo.com' },
  23: { name: 'Backend', path: 'bposeats/bposeats' },
  123: { name: 'Desktop', path: 'bposeats/bposeats-desktop' },
  141: { name: 'Ministry', path: 'bposeats/ministry-vuejs' },
  147: { name: 'Scalema', path: 'bposeats/scalema.com' },
  89: { name: 'BPOSeats.com', path: 'bposeats/bposeats.com' },
  124: { name: 'Android', path: 'bposeats/android-app' },
};

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

async function fetchIssuesFromSheet() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL ISSUES!C4:N', // Adjust the range as needed
  });

  return response.data.values || [];
}

async function fetchAdditionalDataForIssue(issue) {
  const issueId = issue[0]; // Assuming the issue ID is in the first column (C)
  const projectId = issue[1]; // Change this if the project ID is in a different column

  // Find the project configuration based on the project ID
  const projectConfig = PROJECT_CONFIG[projectId];
  if (!projectConfig) {
    console.error(`❌ Project configuration not found for project ID ${projectId}`);
    return ['', 'No', 'Unknown', ''];
  }

  // Fetch comments for the issue
  const commentsResponse = await axios.get(
    `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes`,
    {
      headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
    }
  );

  const comments = commentsResponse.data;
  let firstLgtmCommenter = '';
  let reopenedStatus = 'No';
  let lastReopenedBy = '';
  let lastReopenedAt = '';

  // Process comments to find the required data
  for (const comment of comments) {
    if (firstLgtmCommenter === '' && comment.body.includes('LGTM')) {
      firstLgtmCommenter = comment.author.name;
    }
  }

  // Fetch the issue details to check reopened status
  const issueResponse = await axios.get(
    `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}`,
    {
      headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
    }
  );

  const issueDetails = issueResponse.data;

  if (issueDetails.reopened_at) {
    reopenedStatus = 'Yes';
    lastReopenedBy = issueDetails.reopened_by?.name ?? 'Unknown';
    lastReopenedAt = formatDate(issueDetails.reopened_at);
  }

  return [firstLgtmCommenter, reopenedStatus, lastReopenedBy, lastReopenedAt];
}

async function updateSheetWithAdditionalData(updatedRows) {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  // Clear the target range from C4 to N (to avoid affecting other columns)
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL ISSUES!C4:N',
  });

  // Overwrite the target sheet starting from C4
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL ISSUES!C4',
    valueInputOption: 'USER_ENTERED',
    resource: { values: updatedRows },
  });

  console.log(`✅ Updated the sheet with additional data.`);
}

async function fetchAndUpdateIssues() {
  const issues = await fetchIssuesFromSheet();
  const updatedRows = [];

  for (const issue of issues) {
    const additionalData = await fetchAdditionalDataForIssue(issue);
    updatedRows.push([...issue, ...additionalData]); // Combine original issue with additional data
  }

  await updateSheetWithAdditionalData(updatedRows);
}

fetchAndUpdateIssues();
