from config import sheet_data, credentials_info, AUTOMATED_PORTALS
from constants import SCOPES
from google_auth import get_sheet_service
from sheet_utils import get_spreadsheet_id, get_sheet_metadata
from retry_utils import execute_with_retries
from datetime import datetime
import sys


def construct_log_entry(service, entry):
    spreadsheet_url = entry.get('spreadsheetUrl')
    sheet_name = entry.get('sheetName')
    edited_range = entry.get('editedRange', 'N/A')

    if not spreadsheet_url or not sheet_name:
        raise ValueError(f"Missing spreadsheetUrl or sheetName in entry: {entry}")

    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    metadata = get_sheet_metadata(service, spreadsheet_id)

    matching_sheet = next(
        (s for s in metadata.get('sheets', [])
         if s['properties']['title'] == sheet_name),
        None
    )

    if not matching_sheet:
        raise ValueError(f'Sheet name "{sheet_name}" not found in spreadsheet: {spreadsheet_url}')

    gid = matching_sheet['properties']['sheetId']
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit?gid={gid}#gid={gid}"
    timestamp = datetime.utcnow().isoformat()
    log_message = f"Sheet: {sheet_name} | Range: {edited_range}"

    return [timestamp, sheet_url, log_message]


def main():
    try:
        print("📤 Starting log update to Logs sheet")

        service = get_sheet_service(credentials_info)

        updates = sheet_data if isinstance(sheet_data, list) else [sheet_data]

        log_entries = []
        for entry in updates:
            log_entry = construct_log_entry(service, entry)
            log_entries.append(log_entry)

        append_body = {
            "values": log_entries
        }

        execute_with_retries(
            service.spreadsheets().values().append,
            spreadsheetId=AUTOMATED_PORTALS,
            range="Logs!A:C",
            valueInputOption="USER_ENTERED",
            body=append_body
        )

        print(f"✅ Log(s) added to Logs sheet at {AUTOMATED_PORTALS}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
