import pLimit from 'p-limit';
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
    return JSON.parse(process.env.SHEET_SYNC_SAJ);
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

async function fetchMRsFromSheet(authClient) {
  console.log('⏳ Fetching MRs from sheet...');
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_SYNC_SID,
    range: 'ALL MRs!C4:O', // ID is in D (index 1), Project Name is in O (index 12)
  });
  console.log(`✅ Fetched ${response.data.values?.length || 0} MRs`);
  return response.data.values || [];
}

async function fetchAdditionalDataForMR(mr) {
  const mrId = mr[1]; // column D
  const projectName = mr[12]; // column O
  const projectConfig = PROJECT_CONFIG[projectName];

  if (!projectConfig) {
    console.warn(`⚠️ Project config not found for ${projectName}, skipping MR ${mrId}`);
    return ['Unknown', 'Unknown', 'Unknown', 'Unknown'];
  }

  const projectId = projectConfig.id;

  try {
    // Get MR details
    const mrDetails = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    const { merged_by, target_branch } = mrDetails.data;

    // Get related issues
    const closesIssues = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}/closes_issues`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    const linked = closesIssues.data.map(issue => `#${issue.iid} (ID: ${issue.id})`).join(', ') || 'None';

    // Get comments
    const commentsRes = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}/notes`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    const lgtmComments = commentsRes.data
      .filter(comment => comment.body.match(/\b(?:\*\*LGTM\*\*|LGTM)\b/i))
      .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

    const lastLgtmBy = lgtmComments[0]?.author?.name || 'N/A';

    return [
      merged_by?.name || 'N/A',
      target_branch || 'N/A',
      linked,
      lastLgtmBy
    ];

  } catch (err) {
    console.error(`❌ Error fetching MR ${mrId} (${projectName}):`, err.message);
    return ['Error', 'Error', 'Error', 'Error'];
  }
}

async function updateSheet(authClient, startRow, rows) {
  console.log(`⏳ Updating sheet rows P${startRow}:S${startRow + rows.length - 1}...`);
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_SYNC_SID,
    range: `ALL MRs!P${startRow}:S${startRow + rows.length - 1}`,
    valueInputOption: 'RAW',
    resource: { values: rows },
  });
  console.log('✅ Sheet updated');
}

async function processBatch(authClient, mrsBatch, batchStartIndex) {
  const limit = pLimit(5);
  console.log(`🔄 Processing batch starting at row ${batchStartIndex + 4}...`);

  const results = await Promise.allSettled(
    mrsBatch.map((mr, index) =>
      limit(async () => {
        console.log(`📦 Processing MR row ${batchStartIndex + index + 4} (ID: ${mr[1]})`);
        return await fetchAdditionalDataForMR(mr);
      })
    )
  );

  const rows = results.map(res =>
    res.status === 'fulfilled' ? res.value : ['Error', 'Error', 'Error', 'Error']
  );

  await updateSheet(authClient, batchStartIndex + 4, rows);
}

async function main() {
  const authClient = await auth.getClient();
  const mrs = await fetchMRsFromSheet(authClient);

  const batchSize = 1000;
  for (let i = 0; i < mrs.length; i += batchSize) {
    await processBatch(authClient, mrs.slice(i, i + batchSize), i);
  }
}

main().catch((err) => {
  console.error('❌ Fatal error:', err);
  process.exit(1);
});
