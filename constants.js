// Source 001
export const UTILS_SHEET_ID = process.env.LEADS_CDS_SID;

// Source 002
export const CENTRAL_ISSUE_SHEET_ID = process.env.SHEET_SYNC_SID;
export const ALL_ISSUES = 'ALL ISSUES!C4:U';
export const ALL_NTC = 'ALL ISSUES!C4:N';
export const ALL_MR = 'ALL MRs!C4:O';

// Target
export const DASHBOARD_SHEET = 'Dashboard';
export const G_MILESTONES = 'G-Milestones';
export const G_ISSUES_SHEET = 'G-Issues';
export const G_MR_SHEET = 'G-MR';
export const NTC_SHEET = 'NTC';

// Platform & Database
export const GITLAB_URL = process.env.GITLAB_URL.endsWith('/')
  ? process.env.GITLAB_URL
  : process.env.GITLAB_URL + '/';

export const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
export const SHEET_SYNC_SID = process.env.SHEET_SYNC_SID;
