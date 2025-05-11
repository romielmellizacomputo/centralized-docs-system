import dotenv from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

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
const SERVICE_ACCOUNT = JSON.parse(fs.readFileSync(path.resolve(__dirname, process.env.GOOGLE_SERVICE_ACCOUNT_JSON)));

const LABELS_TO_PROCESS = [
  "To Do", "Doing", "Changes Requested",
  "Manual QA For Review", "QA Lead For Review",
  "Automation QA For Review", "Done", "On Hold", "Deprecated",
  "Automation Team For Review",
];

const PROJECT_IDS = {
  155: 'HQZen',
  23: 'Backend',
  124: 'Android',
  123: 'Desktop',
  88: 'ApplyBPO',
  141: 'Ministry',
  147: 'Scalema',
  89: 'BPOSeats.com'
};

async function authorizeGoogle() {
  const auth = new google.auth.GoogleAuth({
    credentials: SERVICE_ACCOUNT,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth.getClient();
}

function extractProjectIdFromUrl(url) {
  const urlParts = url.split('/');
  const projectName = urlParts[4];

  for (const [id, name] of Object.entries(PROJECT_IDS)) {
    if (projectName.toLowerCase().includes(name.toLowerCase())) {
      return { id, name };
    }
  }
  return null;
}

async function fetchIssue(projectId, issueId) {
  const apiUrl = `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}`;
  const res = await axios.get(apiUrl, {
    headers: {
      'PRIVATE-TOKEN': GITLAB_TOKEN,
    },
  });
  return res.data;
}

async function reviewMetricsLabels() {
  const authClient = await authorizeGoogle();
  const sheets = google.sheets({ version: 'v4', auth: authClient });

  const range = 'TC Review!A2:J'; // Adjust as needed
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range,
  });

  const rows = response.data.values || [];
  const updates = [];

  console.log(`Processing ${rows.length} rows...`);

  for (let i = 0; i < rows.length; i++) {
    const rowIndex = i + 2; // Because spreadsheet is 1-based and starts from row 2
    const url = rows[i][7] || ''; // Column H (index 7)

    if (!url || !/^https:\/\/forge\.bposeats\.com\/[^\/]+\/[^\/]+\/-\/issues\/\d+$/.test(url)) {
      console.log(`Row ${rowIndex}: Skipped due to invalid or missing URL`);
      continue;
    }

    const issueId = url.split('/').pop();
    const project = extractProjectIdFromUrl(url);

    if (!project) {
      console.log(`Row ${rowIndex}: Project not found for URL`);
      continue;
    }

    try {
      const issue = await fetchIssue(project.id, issueId);
      const label = LABELS_TO_PROCESS.find((l) => issue.labels.includes(l));

      if (label) {
        const note = [
          `Title: ${issue.title || 'No Title'}`,
          `Author: ${issue.author?.name || 'Unknown'}`,
          `Assignee: ${issue.assignee?.name || 'Unassigned'}`,
          `Status: ${issue.state.charAt(0).toUpperCase() + issue.state.slice(1)}`,
          `Created At: ${new Date(issue.created_at).toLocaleString()}`,
          `Labels: ${issue.labels.join(', ') || 'None'}`,
          `URL: ${url}`,
        ].join('\n');

        updates.push({
          range: `J${rowIndex}`,
          values: [[label]],
          note,
        });

        updates.push({
          range: `H${rowIndex}`,
          values: [[`=HYPERLINK("${url}", "#${issueId}")`]],
        });

        console.log(`Row ${rowIndex}: Label set to "${label}"`);
      }
    } catch (err) {
      console.error(`Row ${rowIndex}: Error fetching issue - ${err.message}`);
    }

    if (updates.length >= 300) break;
  }

  // Batch update values
  if (updates.length > 0) {
    const valueUpdates = updates.map(({ range, values }) => ({ range, values }));
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: valueUpdates,
      },
    });

    // Add cell notes separately (Google Sheets API v4 doesn’t support adding notes via batch)
    for (const update of updates.filter(u => u.note)) {
      await sheets.spreadsheets.developerMetadata.create({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          metadataKey: `note-${update.range}`,
          metadataValue: update.note,
          location: {
            dimensionRange: {
              sheetId: 0,
              dimension: 'ROWS',
              startIndex: parseInt(update.range.match(/\d+/)[0]) - 1,
              endIndex: parseInt(update.range.match(/\d+/)[0])
            }
          },
          visibility: 'DOCUMENT',
        }
      }).catch(() => {}); // Notes API workaround, often skipped
    }
  }

  console.log(`✅ Finished processing ${updates.length} updates`);
}

reviewMetricsLabels().catch((err) => console.error(`❌ Error: ${err.message}`));
