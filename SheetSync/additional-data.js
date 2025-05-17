import pLimit from 'p-limit';  // npm install p-limit
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
  HQZen: { id: 155, path: 'bposeats/hqzen.com' },
  ApplyBPO: { id: 88, path: 'bposeats/applybpo.com' },
  Backend: { id: 23, path: 'bposeats/bposeats' },
  Desktop: { id: 123, path: 'bposeats/bposeats-desktop' },
  Ministry: { id: 141, path: 'bposeats/ministry-vuejs' },
  Scalema: { id: 147, path: 'bposeats/scalema.com' },
  'BPOSeats.com': { id: 89, path: 'bposeats/bposeats.com' },
  Android: { id: 124, path: 'bposeats/android-app' },
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

async function fetchIssuesFromSheet(authClient) {
  console.log('⏳ Fetching issues from sheet...');
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL ISSUES!C4:N',
  });
  console.log(`✅ Fetched ${response.data.values?.length || 0} issues`);
  return response.data.values || [];
}

async function clearTargetColumns(authClient) {
  console.log('⏳ Clearing target columns O:R...');
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL ISSUES!O4:R',
  });
  console.log('✅ Cleared target columns');
}

async function fetchAdditionalDataForIssue(issue) {
  const issueId = issue[1]; // Assuming the issue ID is in the second column (D)
  const projectName = issue[11];
  const projectConfig = PROJECT_CONFIG[projectName];

  if (!projectConfig) {
    console.warn(`⚠️ Project config not found for ${projectName}, skipping issue ${issueId}`);
    return ['', 'No', 'Unknown', ''];
  }

  const projectId = projectConfig.id;
  try {
    const commentsResponse = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    const comments = commentsResponse.data;
    let firstLgtmCommenter = '';
    let reopenedStatus = 'No';
    let lastReopenedBy = '';
    let lastReopenedAt = '';

    for (const comment of comments) {
      if (
        firstLgtmCommenter === '' &&
        (comment.body.includes('LGTM') || comment.body.includes('**LGTM**'))
      ) {
        firstLgtmCommenter = comment.author.name;
      }
      if (comment.body.includes('reopened')) {
        reopenedStatus = 'Yes';
        lastReopenedBy = comment.author.name;
        lastReopenedAt = comment.created_at;
      }
    }
    return [firstLgtmCommenter, reopenedStatus, lastReopenedBy, lastReopenedAt];
  } catch (error) {
    console.error(`❌ Error fetching comments for issue ${issueId} in project ${projectName}:`, error.message);
    return ['', 'Error', 'Error', 'Error'];
  }
}

async function updateSheet(authClient, rows) {
  console.log('⏳ Updating sheet with fetched data...');
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  // Prepare the data in the range O4:R
  // Assuming rows.length matches the number of issues, and each row is an array of length 4
  const resource = {
    values: rows,
  };
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_SYNC_SID,
    range: `ALL ISSUES!O4:R${rows.length + 3}`, // 4th row is the first data row, so offset by 3
    valueInputOption: 'RAW',
    resource,
  });
  console.log('✅ Sheet updated');
}

async function main() {
  const authClient = await auth.getClient();
  await clearTargetColumns(authClient);
  const issues = await fetchIssuesFromSheet(authClient);

  const limit = pLimit(5); // Limit concurrency to 5 requests at a time
  console.log(`⏳ Processing ${issues.length} issues with concurrency limit 5`);

  const results = await Promise.allSettled(
    issues.map((issue, index) =>
      limit(async () => {
        console.log(`🛠️ Processing issue ${index + 1}/${issues.length} (ID: ${issue[1]})`);
        const data = await fetchAdditionalDataForIssue(issue);
        console.log(`✅ Completed issue ${index + 1} (ID: ${issue[1]})`);
        return data;
      })
    )
  );

  // Map results to data rows, default to empty arrays on failure
  const rows = results.map((res, i) => {
    if (res.status === 'fulfilled') return res.value;
    console.error(`❌ Failed to fetch data for issue ${i + 1}`, res.reason);
    return ['', 'Error', 'Error', 'Error'];
  });

  await updateSheet(authClient, rows);
}

main().catch((error) => {
  console.error('❌ Fatal error:', error);
  process.exit(1);
});
