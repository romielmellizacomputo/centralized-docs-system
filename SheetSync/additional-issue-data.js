import pLimit from 'p-limit';
import { config } from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import http from 'http';
import https from 'https';

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

// Axios instance with connection pooling and retry logic
const axiosInstance = axios.create({
  headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
  timeout: 30000,
  maxRedirects: 5,
  httpAgent: new http.Agent({ keepAlive: true, maxSockets: 10 }),
  httpsAgent: new https.Agent({ keepAlive: true, maxSockets: 10 }),
});

// Add retry interceptor with better rate limit handling
axiosInstance.interceptors.response.use(
  (response) => response,
  async (error) => {
    const config = error.config;
    if (!config || !config.retry) config.retry = 0;
    
    if (config.retry >= 3) return Promise.reject(error);
    
    // Handle rate limiting (429) and connection errors
    if (error.response?.status === 429 || error.code === 'ECONNRESET') {
      config.retry += 1;
      const delay = Math.min(1000 * Math.pow(2, config.retry), 10000);
      console.log(`⏸️  Rate limited, waiting ${delay}ms before retry...`);
      await new Promise(resolve => setTimeout(resolve, delay));
      return axiosInstance(config);
    }
    
    return Promise.reject(error);
  }
);

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

// Cache for issue data to avoid redundant API calls
const issueCache = new Map();

async function fetchAdditionalDataForIssue(issue) {
  const issueId = issue[1];
  const projectName = issue[11];
  const projectConfig = PROJECT_CONFIG[projectName];

  if (!projectConfig) {
    console.warn(`⚠️ Project config not found for ${projectName}, skipping issue ${issueId}`);
    return ['', 'No', 'Unknown', '', 'No Local Status', ''];
  }

  const projectId = projectConfig.id;
  const cacheKey = `${projectId}:${issueId}`;

  // Check cache first
  if (issueCache.has(cacheKey)) {
    return issueCache.get(cacheKey);
  }

  try {
    // Fetch all data in parallel
    const [commentsResponse, issueDetailsResponse, labelEventsResponse] = await Promise.all([
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes?per_page=100`),
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}`),
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/resource_label_events?per_page=100`),
    ]);

    const comments = commentsResponse.data;
    const labelEvents = labelEventsResponse.data;

    let firstLgtmCommenter = '';
    let reopenedStatus = 'No';
    let lastReopenedBy = '';
    let lastReopenedAt = '';

    // Process LGTM comments
    const lgtmComments = comments
      .filter(comment => /\b(?:\*\*LGTM\*\*|LGTM)\b/i.test(comment.body))
      .sort((a, b) => new Date(a.created_at) - new Date(b.created_at));

    if (lgtmComments.length > 0) {
      firstLgtmCommenter = lgtmComments[0].author.name;
    }

    // Process reopened status
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

    // Process label events
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

    const result = [firstLgtmCommenter, reopenedStatus, lastReopenedBy, lastReopenedAt, statusLabel, labelAddedDate];
    
    // Cache the result
    issueCache.set(cacheKey, result);
    
    return result;
  } catch (error) {
    console.error(`❌ Error fetching data for issue ${issueId} in ${projectName}:`, error.message);
    return ['', 'Error', 'Error', 'Error', 'Error', 'Error'];
  }
}

async function updateSheet(authClient, startRow, rows) {
  if (rows.length === 0) return;
  
  console.log(`⏳ Updating sheet rows O${startRow}:T${startRow + rows.length - 1}...`);
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const resource = { values: rows };
  
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_SYNC_SID,
      range: `ALL ISSUES!O${startRow}:T${startRow + rows.length - 1}`,
      valueInputOption: 'RAW',
      resource,
    });
    console.log(`✅ Sheet updated (${rows.length} rows)`);
  } catch (error) {
    console.error('❌ Error updating sheet:', error.message);
    throw error;
  }
}

async function processBatch(authClient, issuesBatch, batchStartIndex) {
  // Lower concurrency to respect GitLab rate limits (7200/hour = 2 req/sec)
  const limit = pLimit(2);
  console.log(`⏳ Processing batch of ${issuesBatch.length} issues starting at row ${batchStartIndex + 4}...`);

  const results = await Promise.allSettled(
    issuesBatch.map((issue, index) =>
      limit(async () => {
        const globalIndex = batchStartIndex + index + 1;
        if (globalIndex % 100 === 0) {
          console.log(`🛠️ Processing issue ${globalIndex} (ID: ${issue[1]})`);
        }
        const data = await fetchAdditionalDataForIssue(issue);
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
  
  // Longer delay between batches to stay well under rate limits
  // This gives GitLab breathing room and other users can work normally
  await new Promise(resolve => setTimeout(resolve, 2000));
}

async function main() {
  console.log('🚀 Starting optimized sync process...');
  console.log('⚙️  Rate-limited mode: 2 concurrent requests, 2s batch delay');
  const startTime = Date.now();
  
  const authClient = await auth.getClient();
  const issues = await fetchIssuesFromSheet(authClient);

  console.log(`📊 Total issues to process: ${issues.length}`);
  console.log(`⏱️  Estimated time: ${Math.ceil(issues.length / 100 * 1.5)} minutes (approximate)\n`);

  // Smaller batch size for more frequent updates and better memory management
  const batchSize = 100;
  
  for (let i = 0; i < issues.length; i += batchSize) {
    const batch = issues.slice(i, i + batchSize);
    const progress = ((i / issues.length) * 100).toFixed(1);
    const elapsed = ((Date.now() - startTime) / 1000 / 60).toFixed(1);
    console.log(`\n📈 Progress: ${progress}% (${i}/${issues.length}) - ${elapsed}min elapsed`);
    
    await processBatch(authClient, batch, i);
    
    // Longer pause every 1000 rows to give GitLab extended breathing room
    if (i % 1000 === 0 && i > 0) {
      console.log('😴 Taking a 10-second break to be nice to GitLab...');
      await new Promise(resolve => setTimeout(resolve, 10000));
    }
    
    // Clear cache periodically to manage memory
    if (i % 5000 === 0 && i > 0) {
      console.log('🧹 Clearing cache to manage memory...');
      issueCache.clear();
    }
  }

  const duration = ((Date.now() - startTime) / 1000 / 60).toFixed(2);
  console.log(`\n✅ All issues processed successfully in ${duration} minutes!`);
  console.log(`📊 Cache hits saved ${issueCache.size} API calls`);
}

main().catch((error) => {
  console.error('❌ Fatal error:', error);
  process.exit(1);
});
