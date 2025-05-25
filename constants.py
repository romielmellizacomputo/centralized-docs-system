import os
from datetime import datetime
import pytz  # Make sure to install pytz: pip install pytz

# Constants
UTILS_SHEET_ID = os.getenv('LEADS_CDS_SID')
CDS_MASTER_ROSTER = os.getenv('CDS_MASTER_ROSTER')

CENTRAL_ISSUE_SHEET_ID = os.getenv('SHEET_SYNC_SID')
ALL_ISSUES = 'ALL ISSUES!C4:T'
ALL_NTC = 'ALL ISSUES!C4:N'
ALL_MR = 'ALL MRs!C4:S'

DASHBOARD_SHEET = 'Dashboard'
G_MILESTONES = 'G-Milestones'
G_ISSUES_SHEET = 'G-Issues'
G_MR_SHEET = 'G-MR'
NTC_SHEET = 'NTC'

# Platform & Database
raw_gitlab_url = os.getenv('GITLAB_URL', '')

GITLAB_URL = raw_gitlab_url if raw_gitlab_url.endswith('/') else raw_gitlab_url + '/'
GITLAB_TOKEN = os.getenv('GITLAB_TOKEN', '')
SHEET_SYNC_SID = os.getenv('SHEET_SYNC_SID', '')

# Timestamp Function
def generate_timestamp_string():
    now = datetime.now()
    
    time_zone_eat = pytz.timezone('Africa/Nairobi')
    time_zone_pht = pytz.timezone('Asia/Manila')

    formatted_date_eat = now.astimezone(time_zone_eat).strftime('%B %d, %Y')
    formatted_date_pht = now.astimezone(time_zone_pht).strftime('%B %d, %Y')

    formatted_eat = now.astimezone(time_zone_eat).strftime('%I:%M:%S %p')
    formatted_pht = now.astimezone(time_zone_pht).strftime('%I:%M:%S %p')

    return f"Sync on {formatted_date_eat}, {formatted_eat} (UG) / {formatted_date_pht}, {formatted_pht} (PH)"
