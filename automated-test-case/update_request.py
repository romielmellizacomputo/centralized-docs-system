import os
from datetime import datetime
import sys

from config import credentials_info
from constants import SCOPES
from google_auth import get_sheet_service
from sheet_utils import get_spreadsheet_id, get_sheet_metadata
from retry_utils import execute_with_retries

# Expecting this env var to be defined in GitHub Secrets or local env
TARGET_SPREADSHEET_ID = os.getenv('AUTOMATED_PORTALS')

# This should be passed or loaded depending on your setup
# Example mock data (replace or integrate accordingly)
sheet_data = [
    {
        "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/your-id/edit",
        "sheetName": "Sheet1",
        "editedRange": "B2:C5"
    }
]


def get_gid_from_metadata(service, spreadsheet_id, sheet_name):
    metadata = get_sheet_metadata(service, spreadsheet_id)
    for sheet in metadata.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    raise ValueError(f'Sheet "{sheet_name}" not found in spreadsheet ID: {spreadsheet_id}')


def create_log_entries(service, updates):
    log_entries = []

    for entry in updates:
        spreadsheet_url = entry.get('spreadsheetUrl')
        sheet_name = entry.get('sheetName')
        edited_range = entry.get('editedRange', 'N/A')

        if not spreadsheet_url or not sheet_name:
            raise ValueError(f"Missing spreadsheetUrl or sheetName in entry: {entry}")

        spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
        gid = get_gid_from_metadata(service, spreadsheet_id, sheet_name)

        sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit?gid={gid}#gid={gid}"
        log_message = f"Sheet: {sheet_name} | Range: {edited_range}"
        timestamp = datetime.utcnow().isoformat()

        log_entries.append([timestamp, sheet_url, log_message])

    return log_entries


def append_logs_to_sheet(service, log_entries):
    body = {
        'values': log_entries
    }

    request = service.spreadsheets().values().append(
        spreadsheetId=TARGET_SPREADSHEET_ID,
        range='Logs!A:C',
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    )

    execute_with_retries(request)


def main():
    if not TARGET_SPREADSHEET_ID:
        print("❌ ERROR: Missing required env variable: AUTOMATED_PORTALS")
        sys.exit(1)

    print("📤 Starting log updates to Logs sheet")

    try:
        service = get_sheet_service(credentials_info)
        log_entries = create_log_entries(service, sheet_data)
        append_logs_to_sheet(service, log_entries)
        print(f"✅ Log(s) successfully appended to Logs sheet at {TARGET_SPREADSHEET_ID}")
    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
