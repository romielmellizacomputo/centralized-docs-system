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

// Axios instance with connection pooling and retry logic
const axiosInstance = axios.create({
  headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
  timeout: 30000,
  maxRedirects: 5,
  httpAgent: new http.Agent({ keepAlive: true, maxSockets: 10 }),
  httpsAgent: new https.Agent({ keepAlive: true, maxSockets: 10 }),
});

// Add retry interceptor
axiosInstance.interceptors.response.use(
  (response) => response,
  async (error) => {
    const config = error.config;
    if (!config || !config.retry) config.retry = 0;
    
    if (config.retry >= 3) return Promise.reject(error);
    
    if (error.response?.status === 429 || error.code === 'ECONNRESET') {
      config.retry += 1;
      const delay = Math.min(1000 * Math.pow(2, config.retry), 10000);
      await new Promise(resolve => setTimeout(resolve, delay));
      return axiosInstance(config);
    }
    
    return Promise.reject(error);
  }
);

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

// Cache for MR data to avoid redundant API calls
const mrCache = new Map();

async function fetchAdditionalDataForMR(mr) {
  const mrId = mr[1]; // column D
  const projectName = mr[12]; // column O
  const projectConfig = PROJECT_CONFIG[projectName];

  if (!projectConfig) {
    console.warn(`⚠️ Project config not found for ${projectName}, skipping MR ${mrId}`);
    return ['Unknown', 'Unknown', 'Unknown', 'Unknown'];
  }

  const projectId = projectConfig.id;
  const cacheKey = `${projectId}:${mrId}`;

  // Check cache first
  if (mrCache.has(cacheKey)) {
    return mrCache.get(cacheKey);
  }

  try {
    // Fetch all data in parallel for better performance
    const [mrDetailsRes, closesIssuesRes, commentsRes] = await Promise.all([
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}`),
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}/closes_issues?per_page=100`),
      axiosInstance.get(`${GITLAB_URL}api/v4/projects/${projectId}/merge_requests/${mrId}/notes?per_page=100`),
    ]);

    const { merged_by, target_branch } = mrDetailsRes.data;

    // Process linked issues
    const linked = closesIssuesRes.data.map(issue => `#${issue.iid} (ID: ${issue.id})`).join(', ') || 'None';

    // Process LGTM comments
    const lgtmComments = commentsRes.data
      .filter(comment => /\b(?:\*\*LGTM\*\*|LGTM)\b/i.test(comment.body))
      .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

    const lastLgtmBy = lgtmComments[0]?.author?.name || 'N/A';

    const result = [
      merged_by?.name || 'N/A',
      target_branch || 'N/A',
      linked,
      lastLgtmBy
    ];

    // Cache the result
    mrCache.set(cacheKey, result);

    return result;

  } catch (err) {
    console.error(`❌ Error fetching MR ${mrId} (${projectName}):`, err.message);
    return ['Error', 'Error', 'Error', 'Error'];
  }
}

async function updateSheet(authClient, startRow, rows) {
  if (rows.length === 0) return;

  console.log(`⏳ Updating sheet rows P${startRow}:S${startRow + rows.length - 1}...`);
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_SYNC_SID,
      range: `ALL MRs!P${startRow}:S${startRow + rows.length - 1}`,
      valueInputOption: 'RAW',
      resource: { values: rows },
    });
    console.log(`✅ Sheet updated (${rows.length} rows)`);
  } catch (error) {
    console.error('❌ Error updating sheet:', error.message);
    throw error;
  }
}

async function processBatch(authClient, mrsBatch, batchStartIndex) {
  // Reduce concurrency to avoid overwhelming GitLab
  const limit = pLimit(3);
  console.log(`🔄 Processing batch of ${mrsBatch.length} MRs starting at row ${batchStartIndex + 4}...`);

  const results = await Promise.allSettled(
    mrsBatch.map((mr, index) =>
      limit(async () => {
        const globalIndex = batchStartIndex + index + 1;
        if (globalIndex % 100 === 0) {
          console.log(`📦 Processing MR ${globalIndex} (ID: ${mr[1]})`);
        }
        return await fetchAdditionalDataForMR(mr);
      })
    )
  );

  const rows = results.map((res, i) => {
    if (res.status === 'fulfilled') return res.value;
    console.error(`❌ Failed to fetch data for MR ${batchStartIndex + i + 1}`, res.reason);
    return ['Error', 'Error', 'Error', 'Error'];
  });

  await updateSheet(authClient, batchStartIndex + 4, rows);
  
  // Small delay between batches to be nice to GitLab
  await new Promise(resolve => setTimeout(resolve, 500));
}

async function main() {
  console.log('🚀 Starting optimized MR sync process...');
  const startTime = Date.now();

  const authClient = await auth.getClient();
  const mrs = await fetchMRsFromSheet(authClient);

  console.log(`📊 Total MRs to process: ${mrs.length}`);

  // Smaller batch size for more frequent updates and better memory management
  const batchSize = 250;

  for (let i = 0; i < mrs.length; i += batchSize) {
    const batch = mrs.slice(i, i + batchSize);
    const progress = ((i / mrs.length) * 100).toFixed(1);
    console.log(`\n📈 Progress: ${progress}% (${i}/${mrs.length})`);
    
    await processBatch(authClient, batch, i);
    
    // Clear cache periodically to manage memory
    if (i % 5000 === 0 && i > 0) {
      console.log('🧹 Clearing cache to manage memory...');
      mrCache.clear();
    }
  }

  const duration = ((Date.now() - startTime) / 1000 / 60).toFixed(2);
  console.log(`\n✅ All MRs processed successfully in ${duration} minutes!`);
  console.log(`📊 Cache hits would have saved ${mrCache.size} API calls`);
}

main().catch((err) => {
  console.error('❌ Fatal error:', err);
  process.exit(1);
});
