import os

# Google Sheet IDs
UTILS_SHEET_ID = os.getenv('LEADS_CDS_SID')
CENTRAL_ISSUE_SHEET_ID = os.getenv('SHEET_SYNC_SID')

# Data Ranges
ALL_ISSUES = 'ALL ISSUES!C4:T'
ALL_NTC = 'ALL ISSUES!C4:N'
ALL_MR = 'ALL MRs!C4:S'

# Sheet Names
DASHBOARD_SHEET = 'Dashboard'
G_MILESTONES = 'G-Milestones'
G_ISSUES_SHEET = 'G-Issues'
G_MR_SHEET = 'G-MR'
NTC_SHEET = 'NTC'

# GitLab
raw_gitlab_url = os.getenv('GITLAB_URL', '')
GITLAB_URL = raw_gitlab_url if raw_gitlab_url.endswith('/') else raw_gitlab_url + '/'
GITLAB_TOKEN = os.getenv('GITLAB_TOKEN', '')
SHEET_SYNC_SID = os.getenv('SHEET_SYNC_SID', '')
