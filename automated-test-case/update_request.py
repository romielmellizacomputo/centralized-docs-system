import os
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from sheet_utils import get_spreadsheet_id  # Assuming you have this function
import sys

from config import sheet_data, credentials_info  # Your update data and credentials

# Scopes needed for read + write access
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheet_service(credentials_info):
    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def main():
    target_spreadsheet_id = os.environ.get('AUTOMATED_PORTALS')
    if not target_spreadsheet_id:
        print("❌ ERROR: Missing environment variable AUTOMATED_PORTALS")
        sys.exit(1)

    updates = sheet_data if isinstance(sheet_data, list) else [sheet_data]

    try:
        service = get_sheet_service(credentials_info)
        sheets_api = service.spreadsheets()

        log_entries = []

        for entry in updates:
            spreadsheet_url = entry.get('spreadsheetUrl')
            sheet_name = entry.get('sheetName')
            edited_range = entry.get('editedRange', 'N/A')

            if not spreadsheet_url or not sheet_name:
                raise ValueError(f"Missing spreadsheetUrl or sheetName in entry: {entry}")

            spreadsheet_id = get_spreadsheet_id(spreadsheet_url)

            # Get metadata for the source spreadsheet
            metadata = sheets_api.get(spreadsheetId=spreadsheet_id).execute()
            sheets = metadata.get('sheets', [])
            matching_sheet = next((s for s in sheets if s['properties']['title'] == sheet_name), None)

            if not matching_sheet:
                raise ValueError(f'Sheet name "{sheet_name}" not found in spreadsheet: {spreadsheet_url}')

            gid = matching_sheet['properties']['sheetId']
            sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit?gid={gid}#gid={gid}"
            current_date = datetime.utcnow().isoformat() + 'Z'  # UTC ISO format

            log_message = f"Sheet: {sheet_name} | Range: {edited_range}"

            log_entries.append([current_date, sheet_url, log_message])

        # Append logs to the target spreadsheet
        sheets_api.values().append(
            spreadsheetId=target_spreadsheet_id,
            range='Logs!A:C',
            valueInputOption='USER_ENTERED',
            body={'values': log_entries}
        ).execute()

        print(f"✅ Log(s) added to Logs sheet at {target_spreadsheet_id}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
