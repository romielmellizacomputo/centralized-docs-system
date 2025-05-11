import dotenv from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();

const requiredEnv = [
  'GITLAB_URL',
  'GITLAB_TOKEN',
  'CDS_PORTAL_SPREADSHEET_ID',
  'CDS_PORTALS_SERVICE_ACCOUNT_JSON',
];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

const GITLAB_URL = process.env.GITLAB_URL;
const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
const SPREADSHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;

// Parse the service account JSON from the environment variable
const SERVICE_ACCOUNT = JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON);

const LABELS_TO_PROCESS = [
  'To Do',
  'Doing',
  'Changes Requested',
  'Manual QA For Review',
  'QA Lead For Review',
  'Automation QA For Review',
  'Done',
  'On Hold',
  'Deprecated',
  'Automation Team For Review',
];

const PROJECT_IDS = {
  155: 'HQZen',
  23: 'Backend',
  124: 'Android',
  123: 'Desktop',
  88: 'ApplyBPO',
  141: 'Ministry',
  147: 'Scalema',
  89: 'BPOSeats.com',
};

const EXCLUDED_SHEETS = [
  'Metrics Comparison',
  'Test Scenario Portal',
  'Scenario Extractor',
  'TEMPLATE',
  'Template',
  'Help',
  'Feature Change Log',
  'Logs',
  'UTILS',
];

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

async function fetchIssueNotes(projectId, issueId) {
  const apiUrl = `${GITLAB_URL}api/v4/projects/${projectId}/issues/${issueId}/notes`;
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

  // Get all sheet names
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });
  const sheetTitles = spreadsheet.data.sheets
    .map((sheet) => sheet.properties.title)
    .filter((title) => !EXCLUDED_SHEETS.includes(title));

  for (const title of sheetTitles) {
    console.log(`Processing sheet: ${title}`);
    const range = `'${title}'!E3:E`;
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
    });

    const rows = response.data.values || [];
    const updates = [];

    for (let i = 0; i < rows.length; i++) {
      const rowIndex = i + 3; // Adjusting for starting from row 3
      const url = rows[i][0] || '';

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
          // Fetch the notes (comments) for the issue
          const notes = await fetchIssueNotes(project.id, issueId);
          const lastComment = notes.length > 0 ? notes[notes.length - 1] : null;

          let note = '';
          if (issue.state === 'opened') {
            note += `The ticket was created by ${issue.author?.name || 'Unknown'}\n`;
            note += `The ticket is still open\n`;
            if (issue.assignee) {
              note += `${issue.assignee.name} is the current assignee\n`;
            } else {
              note += `No assignee on the ticket\n`;
            }
          } else if (issue.state === 'closed') {
            note += `The ticket was created by ${issue.author?.name || 'Unknown'}\n`;
            note += `The ticket was closed\n`;
            if (issue.assignee) {
              note += `${issue.assignee.name} was the assignee\n`;
            } else {
              note += `No assignee on the ticket\n`;
            }
          }

          // Add the latest comment to the note if available
          if (lastComment) {
            note += `Last activity: ${lastComment.body}\n`;
            note += `Commented by: ${lastComment.author.name}\n`;
          }

          updates.push({
            range: `'${title}'!I${rowIndex}`,
            values: [[label]],
            note,
          });

          updates.push({
            range: `'${title}'!E${rowIndex}`,
            values: [[`=HYPERLINK("${url}", "#${issueId}")`]],
          });

          console.log(`Row ${rowIndex}: Label set to "${label}"`);
        }
      } catch (err) {
        console.error(`Row ${rowIndex}: Error fetching issue - ${err.message}`);
      }
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

      // Note: Adding notes to cells is not directly supported via the Google Sheets API v4.
      // This functionality would require the use of Google Apps Script or other workarounds.
    }

    console.log(`✅ Finished processing sheet: ${title}`);
  }
}

reviewMetricsLabels().catch((err) => console.error(`❌ Error: ${err.message}`));
