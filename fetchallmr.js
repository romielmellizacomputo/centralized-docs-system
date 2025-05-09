require('dotenv').config();
const { google } = require('googleapis');
const fs = require('fs');
const axios = require('axios');
const path = require('path');

// Load env variables
const {
  GITLAB_URL,
  GITLAB_TOKEN,
  SPREADSHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_KEY_PATH,
} = process.env;

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
    const keyFilePath = path.resolve(GOOGLE_SERVICE_ACCOUNT_KEY_PATH);
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

// Initialize Google Auth
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

async function fetchExistingMergeRequestKeys(sheets) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'ALL MRs!C4:N',
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
    console.error('❌ Failed to read existing merge requests from sheet:', err.message);
    return new Map();
  }
}

async function fetchAndUpdateMRsForAllProjects() {
  const authClient = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const existingMRs = await fetchExistingMergeRequestKeys(sheets);
  let allMRs = [];

  console.log('🔄 Fetching merge requests for all projects...');

  for (const projectId in PROJECT_CONFIG) {
    const config = PROJECT_CONFIG[projectId];
    let page = 1;

    console.log(`🔄 Fetching merge requests for ${config.name}...`);

    while (true) {
      const response = await axios.get(
        `${GITLAB_URL}api/v4/projects/${projectId}/merge_requests?state=all&per_page=100&page=${page}`,
        {
          headers: { 'PRIVATE-TOKEN': GITLAB_TOKEN },
        }
      );

      if (response.status !== 200) {
        console.error(`❌ Failed to fetch page ${page} for ${config.name}`);
        break;
      }

      const mrs = response.data;
      if (mrs.length === 0) break;

      mrs.forEach(mr => {
        const key = `${mr.id}_${mr.iid}`;
        const existingMR = existingMRs.get(key);

        const mrData = [
          mr.id ?? '',
          mr.iid ?? '',
          mr.title && mr.web_url
            ? `=HYPERLINK("${mr.web_url}", "${mr.title.replace(/"/g, '""')}")`
            : 'No Title',
          mr.author?.name ?? 'Unknown Author',
          mr.assignee?.name ?? 'Unassigned',
          (mr.labels || []).join(', '),
          mr.milestone?.title ?? 'No Milestone',
          capitalize(mr.state ?? ''),
          mr.created_at ? formatDate(mr.created_at) : '',
          mr.merged_at ? formatDate(mr.merged_at) : '',
          mr.merged_by?.name ?? '',
          config.name,
        ];

        if (existingMR) {
          existingMRs.set(key, mrData);
        } else {
          allMRs.push(mrData);
        }
      });

      console.log(`✅ Page ${page} fetched (${mrs.length} MRs) for ${config.name}`);
      page++;
    }
  }

  const updatedRows = Array.from(existingMRs.values());

  if (updatedRows.length > 0) {
    const safeRows = updatedRows.map(row =>
      row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
    );

    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL MRs!C4',
        valueInputOption: 'USER_ENTERED',
        resource: { values: safeRows },
      });

      console.log(`✅ Updated ${safeRows.length} merge requests.`);
    } catch (err) {
      console.error('❌ Error updating data:', err.stack || err.message);
    }
  } else {
    console.log('ℹ️ No updates to existing merge requests.');
  }

  if (allMRs.length > 0) {
    const safeNewRows = allMRs.map(row =>
      row.map(cell => (cell == null ? '' : typeof cell === 'object' ? JSON.stringify(cell) : String(cell)))
    );

    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'ALL MRs!C4',
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: safeNewRows },
      });

      console.log(`✅ Inserted ${safeNewRows.length} new merge requests.`);
    } catch (err) {
      console.error('❌ Error inserting new merge requests:', err.stack || err.message);
    }
  } else {
    console.log('ℹ️ No new merge requests to insert.');
  }
}

// Run the function
fetchAndUpdateMRsForAllProjects();
