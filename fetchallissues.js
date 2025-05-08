const { google } = require('googleapis');
const fs = require('fs');
const axios = require('axios');
const path = require('path');

const keyFile = './service-account.json';
const GITLAB_URL = 'https://forge.bposeats.com/';
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
  try {
    const keyFilePath = path.resolve(keyFile);
    const serviceAccount = JSON.parse(fs.readFileSync(keyFilePath, 'utf8'));
    if (serviceAccount.private_key) {
      serviceAccount.private_key = serviceAccount.private_key.replace(/\\n/g, '\n');
    }
    return serviceAccount;
  } catch (error) {
    console.error('❌ Error loading service account:', error.message);
    throw error;
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
  const formatter = new Intl.DateTimeFormat('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  });
  
  return formatter.format(date);
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
    console.error('❌ Failed to read existing issues from sheet:', err.message);
    return new Map();
  }
}

async function fetchAndUpdateIssuesForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const existingIssues = await fetchExistingIssueKeys(sheets); 
  let allIssues = [];

  console.log('🔄 Fetching issues for all projects...');

  for (const projectId in PROJECT_CONFIG) {
    const config = PROJECT_CONFIG[projectId];
    let page = 1;

    console.log(`🔄 Fetching issues for ${config.name}...`);

    while (true) {
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

      issues.forEach(issue => {
        const key = `${issue.id}_${issue.iid}`;
        const existingIssue = existingIssues.get(key);

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

        if (existingIssue) {
          existingIssues.set(key, issueData);
        } else {
          allIssues.push(issueData);
        }
      });

      console.log(`✅ Page ${page} fetched (${issues.length} issues) for ${config.name}`);
      page++;
    }
  }

  const updatedRows = Array.from(existingIssues.values());

  if (updatedRows.length > 0) {
    const safeRows = updatedRows.map(row =>
      row.map(cell => {
        if (cell === null || cell === undefined) return '';
        if (typeof cell === 'object') return JSON.stringify(cell);
        return String(cell);
      })
    );

    try {
      const res = await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4', 
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: safeRows,
        },
      });

      console.log(`✅ Updated ${safeRows.length} issues.`);
    } catch (err) {
      console.error('❌ Error updating data:', err.stack || err.message);
    }
  } else {
    console.log('ℹ️ No updates to existing issues.');
  }

  // Insert new issues if any
  if (allIssues.length > 0) {
    const safeNewRows = allIssues.map(row =>
      row.map(cell => {
        if (cell === null || cell === undefined) return '';
        if (typeof cell === 'object') return JSON.stringify(cell);
        return String(cell);
      })
    );

    try {
      const res = await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL ISSUES!C4',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: {
          values: safeNewRows,
        },
      });

      console.log(`✅ Inserted ${safeNewRows.length} new issues.`);
    } catch (err) {
      console.error('❌ Error inserting new issues:', err.stack || err.message);
    }
  } else {
    console.log('ℹ️ No new issues to insert.');
  }
}

fetchAndUpdateIssuesForAllProjects();
