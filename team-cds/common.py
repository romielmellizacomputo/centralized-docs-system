# common.py
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
        range='UTILS!B2:B'
    ).execute()
    values = result.get('values', [])
    return [item for sublist in values for item in sublist if item]


def get_selected_milestones(sheets, sheet_id, g_milestones):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f'{g_milestones}!G4:G'
    ).execute()
    values = result.get('values', [])
    return [item for sublist in values for item in sublist if item]
