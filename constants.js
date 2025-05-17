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
import { config } from 'dotenv';
config();

const requiredEnv = ['GITLAB_URL', 'GITLAB_TOKEN', 'SHEET_SYNC_SID', 'SHEET_SYNC_SAJ', 'PROJECT_CONFIG'];
requiredEnv.forEach((key) => {
  if (!process.env[key]) {
    console.error(`❌ Missing required environment variable: ${key}`);
    process.exit(1);
  }
});

export const GITLAB_URL = process.env.GITLAB_URL.endsWith('/')
  ? process.env.GITLAB_URL
  : process.env.GITLAB_URL + '/';

export const GITLAB_TOKEN = process.env.GITLAB_TOKEN;
export const SHEET_SYNC_SID = process.env.SHEET_SYNC_SID;

export const PROJECT_CONFIG = (() => {
  try {
    return JSON.parse(process.env.PROJECT_CONFIG);
  } catch (err) {
    console.error('❌ Failed to parse PROJECT_CONFIG JSON from environment variable:', err);
    process.exit(1);
  }
})();

export function loadServiceAccount() {
  try {
    return JSON.parse(process.env.SHEET_SYNC_SAJ);
  } catch (err) {
    console.error('❌ Error parsing service account JSON:', err.message);
    process.exit(1);
  }
}

