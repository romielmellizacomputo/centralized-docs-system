import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build


def authenticate():
    credentials_info = json.loads(os.getenv('TEAM_CDS_SERVICE_ACCOUNT_JSON'))
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return credentials


def get_sheet_titles(sheets, spreadsheet_id):
    res = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [sheet['properties']['title'] for sheet in res.get('sheets', [])]
    print(f"📄 Sheets in {spreadsheet_id}:", titles)
    return titles


def get_all_team_cds_sheet_ids(sheets, utils_sheet_id):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=utils_sheet_id,
        range='UTILS!B2:B100'  # <-- fixed range with explicit end row
    ).execute()
    values = result.get('values', [])
    return [item for sublist in values for item in sublist if item]


def get_selected_milestones(sheets, sheet_id, g_milestones):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f'{g_milestones}!G4:G100'  # <-- fixed range
    ).execute()
    values = result.get('values', [])
    return [item for sublist in values for item in sublist if item]


# =========================
# Task Reminders functions
# =========================

import datetime
from constants import UTILS_SHEET_ID, CDS_MASTER_ROSTER

def get_sheet_ids(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=UTILS_SHEET_ID,
        range="UTILS!B2:B"
    ).execute()
    values = result.get("values", [])
    sheet_ids = [row[0].strip() for row in values if row and row[0].strip()]
    print(f"🔗 Found {len(sheet_ids)} valid sheet IDs in UTILS sheet")
    return sheet_ids

def get_assignee_email_map(sheets):
    try:
        result = sheets.spreadsheets().values().get(
            spreadsheetId=CDS_MASTER_ROSTER,
            range="Roster!A4:B"
        ).execute()
        rows = result.get("values", [])
        mapping = {}
        for row in rows:
            if len(row) >= 2:
                name = row[0].strip()
                email = row[1].strip()
                if name and email:
                    mapping[name] = email
        print(f"📋 Loaded {len(mapping)} assignee-email mappings from CDS_MASTER_ROSTER")
        return mapping
    except Exception as e:
        print(f"❌ Failed to load assignee-email mapping: {e}")
        return {}

# def days_since(date_str):
#     try:
#         return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%m/%d/%Y")).days
#     except ValueError:
#         try:
#             return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%a, %b %d, %Y")).days
#         except Exception as e:
#             print(f"⚠️ Could not parse date '{date_str}': {e}")
#             return None

def days_since(date_str):
    try:
        # Normalize unicode (replace narrow no-break space with normal space)
        date_str = date_str.replace("\u202f", " ").strip()
        
        # Try format with full datetime and time
        try:
            parsed = datetime.datetime.strptime(date_str, "%a, %b %d, %Y, %I:%M:%S %p")
        except ValueError:
            try:
                parsed = datetime.datetime.strptime(date_str, "%a, %b %d, %Y")
            except ValueError:
                parsed = datetime.datetime.strptime(date_str, "%m/%d/%Y")
        
        return (datetime.datetime.now() - parsed).days
    except Exception as e:
        print(f"⚠️ Could not parse date '{date_str}': {e}")
        return None
