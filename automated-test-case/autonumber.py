from config import sheet_data, credentials_info
from constants import SCOPES, skip_sheets
from google_auth import get_sheet_service
from sheet_utils import (
    get_spreadsheet_id,
    process_sheet,
    get_sheet_metadata,
)
from retry_utils import execute_with_retries
import sys

def main():
    spreadsheet_url = sheet_data['spreadsheetUrl']
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)

    try:
        metadata = get_sheet_metadata(service, spreadsheet_id)
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]

        for name in sheet_names:
            if name in skip_sheets:
                continue
            process_sheet(service, spreadsheet_id, sheets, name)

        print("✅ All applicable sheets processed.")
    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
