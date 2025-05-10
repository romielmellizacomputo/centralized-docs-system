import { config } from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';

// Load environment variables
config();

const requiredEnv = ['GITLAB_URL', 'GITLAB_TOKEN', 'SPREADSHEET_ID', 'GOOGLE_SERVICE_ACCOUNT_JSON'];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

const GITLAB_URL = process.env.GITLAB_URL!;
const GITLAB_TOKEN = process.env.GITLAB_TOKEN!;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID!;

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
      console.error('❌ Error parsing service account JSON:', (error as Error).message);
      throw error;
    }
  } else {
    console.error('❌ Script must run in GitHub Actions with GOOGLE_SERVICE_ACCOUNT_JSON');
    process.exit(1);
  }
}

const serviceAccount = loadServiceAccount();

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

function capitalize(str: string) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function formatDate(dateString?: string) {
  if (!dateString) return '';
  const date = new Date(dateString);
  return new Intl.DateTimeFormat('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  }).format(date);
}

async function fetchExistingIssueKeys(sheets: any) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ALL ISSUES!C4:N',
    });

    const rows = res.data.values || [];
    const issueKeys = new Map();
    for (const row of rows) {
      const id = row[0]?.trim();
      const iid = row[1]?.trim();
      if (id && iid) {
        issueKeys.set(`${id}_${iid}`, row);
      }
    }
    return issueKeys;
  } catch (err: any) {
    console.error('❌ Failed to read existing issues from sheet:', err.message);
    return new Map();
  }
}

async function fetchIssuesForProject(projectId: string, config: any, existingIssues: Map<string, any[]>) {
  let allIssues: any[] = [];
  let page = 1;

  console.log(`🔄 Fetching issues for ${config.name}...`);

  while (true) {
    const response = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/issues?state=all&per_page=100&page=${page}`,
      { headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN } }
    );

    if (response.status !== 200) {
      console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
      break;
    }

    const issues = response.data;
    if (issues.length === 0) break;

    for (const issue of issues) {
      const key = `${issue.id}_${issue.iid}`;
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

      if (existingIssues.has(key)) {
        existingIssues.set(key, issueData);
      } else {
        allIssues.push(issueData);
      }
    }

    console.log(`✅ Page ${page} fetched (${issues.length} issues) for ${config.name}`);
    page++;
  }

  return allIssues;
}

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const existingIssues = await fetchExistingIssueKeys(sheets);
  const allNewIssues: any[] = [];

  const fetchPromises = Object.entries(PROJECT_CONFIG).map(([projectId, config]) =>
    fetchIssuesForProject(projectId, config, existingIssues)
  );

  const newIssuesArrays = await Promise.all(fetchPromises);
  newIssuesArrays.forEach((issues) => allNewIssues.push(...issues));

  const updatedRows = Array.from(existingIssues.values()).map(row =>
    row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
  );

  if (updatedRows.length > 0) {
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        resource: { values: updatedRows },
      });

      console.log(`✅ Updated ${updatedRows.length} issues.`);
    } catch (err: any) {
      console.error('❌ Error updating existing issues:', err.message);
    }
  }

  if (allNewIssues.length > 0) {
    const safeNewRows = allNewIssues.map(row =>
      row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
    );

    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: safeNewRows },
      });

      console.log(`✅ Inserted ${safeNewRows.length} new issues.`);
    } catch (err: any) {
      console.error('❌ Error inserting new issues:', err.message);
    }
  } else {
    console.log('ℹ️ No new issues to insert.');
  }
}

fetchAndUpdateIssuesForAllProjects();
