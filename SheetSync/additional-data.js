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

async function fetchIssuesForProject(projectId) {
  const response = await axios.get(
    `${GITLAB_URL}api/v4/projects/${projectId}/issues`,
    {
      headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
      params: {
        per_page: 100, // Adjust as needed
        page: 1,
      },
    }
  );

  return response.data;
}

async function fetchAdditionalDataForIssues(issues) {
  const additionalData = [];

  for (const issue of issues) {
    // Debug log to see the structure of the issue
    console.log('Issue:', issue);

    const issueId = issue.id; // Accessing the issue ID directly
    const projectName = issue.project_id; // Change this line to access the correct project name

    // Find the project ID based on the project name
    const projectId = Object.keys(PROJECT_CONFIG).find(id => PROJECT_CONFIG[id].name === projectName);
    if (!projectId) {
      console.error(`❌ Project not found for issue ID ${issueId} with project name ${projectName}`);
      additionalData.push(['', 'No', '', '']);
      continue;
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

    additionalData.push([
      firstLgtmCommenter,
      reopenedStatus,
      lastReopenedBy,
      lastReopenedAt,
    ]);
  }

  return additionalData;
}

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  console.log('🔄 Fetching issues for all projects...');

  const issuesPromises = Object.keys(PROJECT_CONFIG).map(async (projectId) => {
    const issues = await fetchIssuesForProject(projectId);
    
    // Debug log to see the fetched issues
    console.log(`Fetched issues for project ${projectId}:`, issues);

    const additionalData = await fetchAdditionalDataForIssues(issues);

    // Combine the original issues with additional data
    if (issues.length === 0) {
      console.log(`ℹ️ No issues found for project ${projectId}.`);
      return;
    }

    const combinedData = issues.map((issue, index) => [...issue, ...additionalData[index]]);

    if (combinedData.length > 0) {
      const safeRows = combinedData.map((row) =>
        row.map((cell) =>
          cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)
        )
      );

      try {
        // Clear the target range from C4 to R (to avoid affecting other columns)
        await sheets.spreadsheets.values.clear({
          spreadsheetId: SHEET_SYNC_SID,
          range: 'ALL ISSUES!C4:R',
        });

        // Overwrite the target sheet starting from C4
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_SYNC_SID,
          range: 'ALL ISSUES!C4',
          valueInputOption: 'USER_ENTERED',
          resource: { values: safeRows },
        });

        console.log(`✅ Cleared and inserted ${safeRows.length} issues with additional data.`);
      } catch (err) {
        console.error('❌ Error writing to sheet:', err.stack || err.message);
      }
    } else {
      console.log('ℹ️ No issues to insert.');
    }
  });

  await Promise.all(issuesPromises);
}

fetchAndUpdateIssuesForAllProjects();
