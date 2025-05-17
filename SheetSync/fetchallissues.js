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

async function fetchIssuesForProject(projectId, config, page) {
  const response = await axios.get(
    `${GITLAB_URL}api/v4/projects/${projectId}/issues?state=all&per_page=100&page=${page}`,
    {
      headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
    }
  );

  if (response.status !== 200) {
    console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
    return [];
  }

  const fetchedIssues = response.data;
  const issues = [];

  // Collect issue IDs for batch comment fetching
  const issueIds = fetchedIssues.map(issue => issue.iid);
  const commentsResponses = await Promise.all(
    issueIds.map(issueId =>
      axios.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes`, {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      }).catch(err => ({ data: [] })) // Handle errors gracefully
    )
  );

  for (let i = 0; i < fetchedIssues.length; i++) {
    const issue = fetchedIssues[i];
    const comments = commentsResponses[i].data;

    // Find first LGTM commenter
    const firstLgtmCommenter = comments.find(comment => comment.body.includes('LGTM'))?.author.name || 'Unknown';
    
    // Check if the issue was reopened
    const reopenedStatus = issue.state === 'reopened' ? 'Yes' : 'No';
    const lastReopenedBy = reopenedStatus === 'Yes' ? issue.reopened_by?.name || 'Unknown' : 'Unknown';

    // Prepare labels
    const labels = (issue.labels || []).map(label => {
      if (label === 'Bug-issue') return 'Bug Issue';
      if (label === 'Usability Suggestions') return 'Usability Suggestion';
      return label;
    }).join(', ');

    const issueData = [
      issue.id ?? '',
      issue.iid ?? '',
      issue.title && issue.web_url
        ? `=HYPERLINK("${issue.web_url}", "${issue.title.replace(/"/g, '""')}")`
        : 'No Title',
      issue.author?.name ?? 'Unknown Author',
      issue.assignee?.name ?? 'Unassigned',
      labels,
      issue.milestone?.title ?? 'No Milestone',
      capitalize(issue.state ?? ''),
      issue.created_at ? formatDate(issue.created_at) : '',
      issue.closed_at ? formatDate(issue.closed_at) : '',
      issue.closed_by?.name ?? '',
      config.name,
      firstLgtmCommenter,
      reopenedStatus,
      lastReopenedBy
    ];

    issues.push(issueData);
  }

  console.log(`✅ Page ${page} fetched (${fetchedIssues.length} issues) for ${config.name}`);
  return issues;
}

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  console.log('🔄 Fetching issues for all projects...');

  let page = 1;
  let allIssues = [];

  while (true) {
    const issuesPromises = Object.keys(PROJECT_CONFIG).map(async (projectId) => {
      const config = PROJECT_CONFIG[projectId];
      return fetchIssuesForProject(projectId, config, page);
    });

    const allIssuesResults = await Promise.all(issuesPromises);
    const currentPageIssues = allIssuesResults.flat();

    if (currentPageIssues.length === 0) break; // Stop if no issues are returned

    allIssues = [...allIssues, ...currentPageIssues];
    page++;
  }

  if (allIssues.length > 0) {
    const safeRows = allIssues.map((row) =>
      row.map((cell) =>
        cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)
      )
    );

    try {
      // Overwrite the target sheet starting from C4
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_SYNC_SID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        resource: { values: safeRows },
      });

      console.log(`✅ Cleared and inserted ${safeRows.length} issues.`);
    } catch (err) {
      console.error('❌ Error writing to sheet:', err.stack || err.message);
    }
  } else {
    console.log('ℹ️ No issues to insert.');
  }
}

fetchAndUpdateIssuesForAllProjects();
