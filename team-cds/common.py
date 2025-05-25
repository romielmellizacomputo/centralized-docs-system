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

def get_task_reminders(sheets, spreadsheet_id, sheet_name='TaskReminders', data_range='A2:E100'):
    """
    Fetch task reminders data from the specified sheet and range.
    """
    result = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f'{sheet_name}!{data_range}'
    ).execute()
    values = result.get('values', [])
    return values


def get_task_reminder_titles(sheets, spreadsheet_id, sheet_name='TaskReminders'):
    """
    Get titles of all sheets related to Task Reminders.
    """
    res = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [sheet['properties']['title'] for sheet in res.get('sheets', [])]
    # Filter for TaskReminders sheets only (optional, if you have multiple)
    task_reminder_titles = [title for title in titles if 'TaskReminders' in title]
    print(f"📝 Task Reminders Sheets in {spreadsheet_id}:", task_reminder_titles)
    return task_reminder_titles
