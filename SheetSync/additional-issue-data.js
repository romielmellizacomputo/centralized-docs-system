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

async function fetchAdditionalDataForIssue(issue) {
  const issueId = issue[1]; // Assuming the issue ID is in column D (index 1)
  const projectName = issue[11];
  const projectConfig = PROJECT_CONFIG[projectName];

  if (!projectConfig) {
    console.warn(`⚠️ Project config not found for ${projectName}, skipping issue ${issueId}`);
    return ['', 'No', 'Unknown', '', 'No Local Status', ''];
  }

  const projectId = projectConfig.id;
  try {
    const [commentsResponse, issueDetailsResponse] = await Promise.all([
      axios.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes`, {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      }),
      axios.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}`, {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      }),
    ]);

    const comments = commentsResponse.data;
    const issueDetails = issueDetailsResponse.data;

    let firstLgtmCommenter = '';
    let reopenedStatus = 'No';
    let lastReopenedBy = '';
    let lastReopenedAt = '';

    const lgtmComments = comments
      .filter(comment => comment.body.match(/\b(?:\*\*LGTM\*\*|LGTM)\b/i))
      .sort((a, b) => new Date(a.created_at) - new Date(b.created_at));

    if (lgtmComments.length > 0) {
      firstLgtmCommenter = lgtmComments[0].author.name;
    }

    for (const comment of comments) {
      if (comment.body.toLowerCase().includes('reopened')) {
        reopenedStatus = 'Yes';
        lastReopenedBy = comment.author.name;
        lastReopenedAt = new Date(comment.created_at).toLocaleDateString('en-US', {
          weekday: 'short',
          year: 'numeric',
          month: 'short',
          day: 'numeric',
        });
      }
    }

    // Label checks for column S and T
    const labelEventsResponse = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/resource_label_events`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    const labelEvents = labelEventsResponse.data;
    const targetLabels = ['Testing::Local::Passed', 'Testing::Epic::Passed'];

    let statusLabel = 'No Local Status';
    let labelAddedDate = '';

    for (const target of targetLabels) {
      const matched = labelEvents
        .filter(e => e.label?.name === target && e.action === 'add')
        .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
      if (matched.length > 0) {
        statusLabel = target;
        labelAddedDate = new Date(matched[0].created_at).toLocaleDateString('en-US', {
          weekday: 'short',
          year: 'numeric',
          month: 'short',
          day: 'numeric',
        });
        break;
      }
    }

    return [firstLgtmCommenter, reopenedStatus, lastReopenedBy, lastReopenedAt, statusLabel, labelAddedDate];
  } catch (error) {
    console.error(`❌ Error fetching data for issue ${issueId} in ${projectName}:`, error.message);
    return ['', 'Error', 'Error', 'Error', 'Error', 'Error'];
  }
}

async function updateSheet(authClient, startRow, rows) {
  console.log(`⏳ Updating sheet rows O${startRow}:T${startRow + rows.length - 1}...`);
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const resource = { values: rows };
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_SYNC_SID,
    range: `ALL ISSUES!O${startRow}:T${startRow + rows.length - 1}`,
    valueInputOption: 'RAW',
    resource,
  });
  console.log('✅ Sheet updated');
}

async function processBatch(authClient, issuesBatch, batchStartIndex) {
  const limit = pLimit(5);
  console.log(`⏳ Processing batch of ${issuesBatch.length} issues starting at row ${batchStartIndex + 4}...`);

  const results = await Promise.allSettled(
    issuesBatch.map((issue, index) =>
      limit(async () => {
        console.log(`🛠️ Processing issue ${batchStartIndex + index + 1} (ID: ${issue[1]})`);
        const data = await fetchAdditionalDataForIssue(issue);
        console.log(`✅ Completed issue ${batchStartIndex + index + 1} (ID: ${issue[1]})`);
        return data;
      })
    )
  );

  const rows = results.map((res, i) => {
    if (res.status === 'fulfilled') return res.value;
    console.error(`❌ Failed to fetch data for issue ${batchStartIndex + i + 1}`, res.reason);
    return ['', 'Error', 'Error', 'Error', 'Error', 'Error'];
  });

  await updateSheet(authClient, batchStartIndex + 4, rows);
}

async function main() {
  const authClient = await auth.getClient();
  const issues = await fetchIssuesFromSheet(authClient);

  const batchSize = 1000;
  for (let i = 0; i < issues.length; i += batchSize) {
    const batch = issues.slice(i, i + batchSize);
    await processBatch(authClient, batch, i);
  }
}

main().catch((error) => {
  console.error('❌ Fatal error:', error);
  process.exit(1);
});
