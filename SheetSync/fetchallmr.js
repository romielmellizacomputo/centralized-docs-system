import dotenv from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();

const requiredEnv = ['GITLAB_URL', 'GITLAB_TOKEN', 'SPREADSHEET_ID', 'GOOGLE_SERVICE_ACCOUNT_JSON'];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

const GITLAB_URL = process.env.GITLAB_URL;
const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const PROJECT_CONFIG = {
  155: { name: 'HQZen', sheet: 'HQZEN', path: 'bposeats/hqzen.com' },
  88: { name: 'ApplyBPO', sheet: 'APPLYBPO', path: 'bposeats/applybpo.com' },
  23: { name: 'Backend', sheet: 'BACKEND', path: 'bposeats/bposeats' },
  123: { name: 'Desktop', sheet: 'DESKTOP', path: 'bposeats/bposeats-desktop' },
  141: { name: 'Ministry', sheet: 'MINISTRY', path: 'bposeats/ministry-vuejs' },
  147: { name: 'Scalema', sheet: 'SCALEMA', path: 'bposeats/scalema.com' },
  89: { name: 'BPOSeats.com', sheet: 'BPOSEATS', path: 'bposeats/bposeats.com' },
  124: { name: 'Android', sheet: 'ANDROID', path: 'bposeats/android-app' },
};

function loadServiceAccount() {
  if (process.env.GITHUB_ACTIONS && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    try {
      return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    } catch (error) {
      console.error('❌ Error parsing service account JSON:', error.message);
      process.exit(1);
    }
  } else {
    console.error('❌ This script should be run in GitHub Actions.');
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

async function fetchExistingMergeRequestKeys(sheets) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ALL MRs!C4:O',
    });

    const rows = res.data.values || [];
    const mrKeys = new Map();
    for (const row of rows) {
      const id = row[0]?.trim();
      const iid = row[1]?.trim();
      if (id && iid) {
        mrKeys.set(`${id}_${iid}`, row);
      }
    }
    return mrKeys;
  } catch (err) {
    console.error('❌ Failed to read existing MRs from sheet:', err.message);
    return new Map();
  }
}

async function fetchMRsForProject(projectId, config, existingMRs) {
  let page = 1;
  const allProjectMRs = [];

  console.log(`🔄 Fetching merge requests for ${config.name}...`);

  while (true) {
    try {
      const response = await axios.get(
        `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests?state=all&per_page=100&page=${page}`,
        { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
      );

      if (response.status !== 200) {
        console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
        break;
      }

      const mrs = response.data;
      if (mrs.length === 0) break;

      for (const mr of mrs) {
        const key = `${mr.id}_${mr.iid}`;
        const reviewers = (mr.reviewers || []).map(r => r.name).join(', ') || 'Unassigned';

        const mrData = [
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
        ];

        if (existingMRs.has(key)) {
          existingMRs.set(key, mrData);
        } else {
          allProjectMRs.push(mrData);
        }
      }

      console.log(`✅ Page ${page} fetched (${mrs.length} MRs) for ${config.name}`);
      page++;
    } catch (err) {
      console.error(`❌ Error fetching MRs for ${config.name}:`, err.message);
      break;
    }
  }

  return allProjectMRs;
}

async function fetchAndUpdateMRsForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const existingMRs = await fetchExistingMergeRequestKeys(sheets);
  console.log('🔄 Fetching merge requests for all projects in parallel...');

  const fetchPromises = Object.entries(PROJECT_CONFIG).map(([projectId, config]) =>
    fetchMRsForProject(projectId, config, existingMRs)
  );

  const results = await Promise.all(fetchPromises);
  const allNewMRs = results.flat();

  const updatedRows = Array.from(existingMRs.values());

  const cleanData = (rows) =>
    rows.map((row) =>
      row.map((cell) => {
        if (cell === null || cell === undefined) return '';
        if (typeof cell === 'object') return JSON.stringify(cell);
        return String(cell);
      })
    );

  if (updatedRows.length > 0) {
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL MRs!C4',
        valueInputOption: 'USER_ENTERED',
        resource: { values: cleanData(updatedRows) },
      });
      console.log(`✅ Updated ${updatedRows.length} merge requests.`);
    } catch (err) {
      console.error('❌ Error updating data:', err.stack || err.message);
    }
  }

  if (allNewMRs.length > 0) {
    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL MRs!C4',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: cleanData(allNewMRs) },
      });
      console.log(`✅ Inserted ${allNewMRs.length} new merge requests.`);
    } catch (err) {
      console.error('❌ Error inserting new merge requests:', err.stack || err.message);
    }
  }
}

// Run
fetchAndUpdateMRsForAllProjects();
