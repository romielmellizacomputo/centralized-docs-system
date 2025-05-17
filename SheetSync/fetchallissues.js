import { config } from 'dotenv';
import axios from 'axios';
import pLimit from 'p-limit';
import { google } from 'googleapis'; // Ensure you have this import for Google Auth

// Load environment variables
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

// Load service account credentials
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

// Utility functions
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

// Limit the number of concurrent requests
const limit = pLimit(5);

// Fetch comments for issues
async function fetchCommentsForIssues(projectId, issues) {
  const commentPromises = issues.map(issue => 
    limit(() => axios.get(`${GITLAB_URL}api/v4/projects/${projectId}/issues/${issue.iid}/notes`, {
      headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
    }))
  );
  
  const commentsResponses = await Promise.all(commentPromises);
  return commentsResponses.map(response => response.data);
}

// Fetch issues for a project
async function fetchIssuesForProject(projectId, config) {
  let page = 1;
  let issues = [];
  console.log(`🔄 Fetching issues for ${config.name}...`);

  while (true) {
    const response = await axios.get(
      `${GITLAB_URL}api/v4/projects/${projectId}/issues?state=all&per_page=100&page=${page}`,
      {
        headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      }
    );

    if (response.data.length === 0) break; // No more issues to fetch

    issues = issues.concat(response.data);
    page++;
  }

  // Fetch comments for all issues
  const comments = await fetchCommentsForIssues(projectId, issues);
  return { issues, comments };
}

// Main function
async function main() {
  for (const [projectId, config] of Object.entries(PROJECT_CONFIG)) {
    try {
      const { issues, comments } = await fetchIssuesForProject(projectId, config);
      
      // Process issues and comments as needed
      console.log(`🔍 Fetched ${issues.length} issues and ${comments.flat().length} comments for ${config.name}.`);
    } catch (error) {
      console.error(`❌ Error fetching data for project ${config.name}:`, error);
    }
  }
}

// Execute main function
main().catch(error => {
  console.error('❌ An error occurred:', error);
  process.exit(1);
});



async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  console.log('🔄 Fetching issues for all projects...');

  // Fetch issues for all projects
  const issuesPromises = Object.keys(PROJECT_CONFIG).map(async (projectId) => {
    const config = PROJECT_CONFIG[projectId];
    return fetchIssuesForProject(projectId, config);
  });

  const allIssuesResults = await Promise.all(issuesPromises);
  const allIssues = allIssuesResults.flat();

  // Fetch comments for all issues in parallel
  const comments = await fetchCommentsForIssues(allIssues.map(issue => issue.project_id), allIssues);

  const processedIssues = allIssues.map((issue, index) => {
    let firstLgtmCommenter = 'Unknown';
    let lastReopenedBy = 'Unknown';
    let reopenedStatus = 'No';

    comments[index].forEach(comment => {
      if (comment.body.includes('LGTM') && firstLgtmCommenter === 'Unknown') {
        firstLgtmCommenter = comment.author.name;
      }
    });

    if (issue.state === 'reopened') {
      reopenedStatus = 'Yes';
      lastReopenedBy = issue.reopened_by?.name ?? 'Unknown';
    }

    const labels = (issue.labels || []).map(label => {
      if (label === 'Bug-issue') return 'Bug Issue';
      if (label === 'Usability Suggestions') return 'Usability Suggestion';
      return label;
    }).join(', ');

    return [
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
      PROJECT_CONFIG[issue.project_id]?.name ?? 'Unknown Project',
      firstLgtmCommenter,
      reopenedStatus,
      lastReopenedBy
    ];
  });

  if (processedIssues.length > 0) {
    const safeRows = processedIssues.map(row =>
      row.map(cell =>
        cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)
      )
    );

    try {
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
