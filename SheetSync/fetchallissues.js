// GitHub Actions version - no need for dotenv
const { google } = require('googleapis');
const axios = require('axios');

// Validate required GitHub secrets (set in GitHub Actions)
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

// Load and parse service account credentials from GitHub secret
function loadServiceAccount() {
  try {
    if (process.env.GITHUB_ACTIONS && process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    } else {
      console.error('❌ This script is intended to be run in GitHub Actions with GOOGLE_SERVICE_ACCOUNT_JSON secret.');
      process.exit(1);
    }
  } catch (err) {
    console.error('❌ Failed to parse GOOGLE_SERVICE_ACCOUNT_JSON:', err.message);
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

async function fetchExistingIssueKeys(sheets) {
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
  } catch (err) {
    console.error('❌ Failed to read existing issues:', err.message);
    return new Map();
  }
}

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const existingIssues = await fetchExistingIssueKeys(sheets);
  const newIssues = [];

  console.log('🔄 Starting issue fetch and update process...');

  for (const projectId in PROJECT_CONFIG) {
    const config = PROJECT_CONFIG[projectId];
    let page = 1;

    console.log(`📂 Fetching issues for ${config.name}...`);

    while (true) {
      try {
        const response = await axios.get(
          `${GITLAB_URL}api/v4/projects/${projectId}/issues?state=all&per_page=100&page=${page}`,
          {
            headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
          }
        );

        if (response.status !== 200) {
          console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
          break;
        }

        const issues = response.data;
        if (issues.length === 0) break;

        for (const issue of issues) {
          const key = `${issue.id}_${issue.iid}`;
          const rowData = [
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
            existingIssues.set(key, rowData); // Update existing
          } else {
            newIssues.push(rowData); // Collect new
          }
        }

        console.log(`✅ Page ${page} fetched (${issues.length} issues) for ${config.name}`);
        page++;
      } catch (err) {
        console.error(`❌ Error fetching issues for ${config.name}:`, err.message);
        break;
      }
    }
  }

  const updatedRows = Array.from(existingIssues.values());

  if (updatedRows.length > 0) {
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: updatedRows.map(row =>
            row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
          ),
        },
      });
      console.log(`✅ Updated ${updatedRows.length} rows.`);
    } catch (err) {
      console.error('❌ Error updating rows:', err.message);
    }
  } else {
    console.log('ℹ️ No existing issues to update.');
  }

  if (newIssues.length > 0) {
    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: newIssues.map(row =>
            row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
          ),
        },
      });
      console.log(`✅ Inserted ${newIssues.length} new issues.`);
    } catch (err) {
      console.error('❌ Error inserting new issues:', err.message);
    }
  } else {
    console.log('ℹ️ No new issues to insert.');
  }
}

// Execute the update
fetchAndUpdateIssuesForAllProjects().catch((err) => {
  console.error('❌ Fatal error:', err.stack || err.message);
  process.exit(1);
});
