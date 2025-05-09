import { google } from 'googleapis';
import { readFile } from 'fs/promises';

export async function authenticate() {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json', // Google service account credentials
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return await auth.getClient();
}

export async function getAllIssues(sheets) {
  const sheetId = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';
  const range = `${'ALL ISSUES'}!C4:N`;

  const { data } = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    ranges: [range],
    includeGridData: true,
  });

  const rows = data.sheets[0].data[0].rowData || [];

  return rows.map(row => {
    return (row.values || []).map(cell => ({
      text: cell.formattedValue || '',
      hyperlink: cell.hyperlink || '',
    }));
  });
}

export async function getSelectedMilestones(sheets, sheetId) {
  const range = `${'G-Milestones'}!C4:G`;
  const { data } = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range });

  return data.values
    .filter(row => row[0] === 'TRUE' && row[4])
    .map(row => row[4]);
}

export async function clearSheet(sheets, sheetId) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `${'G-Issues'}!C4:N`,
  });
}

export async function insertData(sheets, sheetId, values) {
  if (!values.length) return;

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${'G-Issues'}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values },
  });
}

export async function updateTimestamp(sheets, sheetId) {
  const now = new Date().toLocaleString('en-US', { timeZone: 'Asia/Manila' });
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `${'Dashboard'}!AB9`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [[`Data was updated on ${now}`]],
    },
  });
}

