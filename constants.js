// Source 001
export const UTILS_SHEET_ID = process.env.LEADS_CDS_SID;

// Source 002
export const CENTRAL_ISSUE_SHEET_ID = process.env.SHEET_SYNC_SID;
export const ALL_ISSUES = 'ALL ISSUES!C4:T';
export const ALL_NTC = 'ALL ISSUES!C4:N';
export const ALL_MR = 'ALL MRs!C4:S';

// Target
export const DASHBOARD_SHEET = 'Dashboard';
export const G_MILESTONES = 'G-Milestones';
export const G_ISSUES_SHEET = 'G-Issues';
export const G_MR_SHEET = 'G-MR';
export const NTC_SHEET = 'NTC';

// Platform & Database
const rawGitlabUrl = process.env.GITLAB_URL || '';

export const GITLAB_URL = rawGitlabUrl.endsWith('/')
  ? rawGitlabUrl
  : rawGitlabUrl + '/';

export const GITLAB_TOKEN = process.env.GITLAB_TOKEN || '';
export const SHEET_SYNC_SID = process.env.SHEET_SYNC_SID || '';


// Timestamp
export function generateTimestampString() {
  const now = new Date();
  const timeZoneEAT = 'Africa/Nairobi';
  const timeZonePHT = 'Asia/Manila';

  const optionsDate = {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  };

  const optionsTime = {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  };

  const formattedDateEAT = new Intl.DateTimeFormat('en-US', {
    ...optionsDate,
    timeZone: timeZoneEAT
  }).format(now);

  const formattedDatePHT = new Intl.DateTimeFormat('en-US', {
    ...optionsDate,
    timeZone: timeZonePHT
  }).format(now);

  const formattedEAT = new Intl.DateTimeFormat('en-US', {
    ...optionsTime,
    timeZone: timeZoneEAT
  }).format(now);

  const formattedPHT = new Intl.DateTimeFormat('en-US', {
    ...optionsTime,
    timeZone: timeZonePHT
  }).format(now);

  return `Sync on ${formattedDateEAT}, ${formattedEAT} (UG) / ${formattedDatePHT}, ${formattedPHT} (PH)`;
}
