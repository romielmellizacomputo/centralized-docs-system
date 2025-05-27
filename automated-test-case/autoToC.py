from config import sheet_data, credentials_info
from constants import skip_sheets
from google_auth import get_sheet_service
from sheet_utils import (
    get_spreadsheet_id,
    get_sheet_metadata,
    clear_sheet_range,
    batch_get_values,
    update_sheet_values
)
from retry_utils import execute_with_retries
import requests
import sys


def build_toc_rows(service, spreadsheet_id, spreadsheet_url, sheets):
    toc_rows = []
    existing_titles = set()

    for sheet in sheets:
        name = sheet['properties']['title']
        sheet_id = sheet['properties']['sheetId']

        if name in skip_sheets:
            continue

        # Get value of C4
        c4_range = f"'{name}'!C4"
        c4_value = batch_get_values(service, spreadsheet_id, [c4_range])[0]
        c4_text = c4_value[0][0] if c4_value and c4_value[0] else None

        if not c4_text or c4_text in existing_titles:
            continue

        existing_titles.add(c4_text)
        hyperlink = f'=HYPERLINK("{spreadsheet_url}#gid={sheet_id}", "{c4_text}")'

        # Define ranges to read
        cells = ['C5', 'C7', 'C15', 'C18', 'C19', 'C20', 'C21', 'C14', 'C13', 'C6']
        ranges = [f"'{name}'!{cell}" for cell in cells]
        values = batch_get_values(service, spreadsheet_id, ranges)

        row_data = [hyperlink]
        for val in values:
            cell_val = val[0][0] if val and val[0] else ''
            row_data.append(cell_val)

        toc_rows.append(row_data)
        print(f"✅ Inserted hyperlink for: {c4_text}")

    return toc_rows


def main():
    spreadsheet_url = sheet_data['spreadsheetUrl']
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)

    try:
        metadata = get_sheet_metadata(service, spreadsheet_id)
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]

        toc_sheet = next((s for s in sheets if s['properties']['title'] == 'ToC'), None)
        if not toc_sheet:
            raise Exception("ToC sheet not found.")

        # Clear previous data
        clear_sheet_range(service, spreadsheet_id, "'ToC'!A2:A")
        clear_sheet_range(service, spreadsheet_id, "'ToC'!B2:K")

        # Build new ToC rows
        toc_rows = build_toc_rows(service, spreadsheet_id, spreadsheet_url, sheets)

        # Write to sheet
        if toc_rows:
            update_sheet_values(
                service,
                spreadsheet_id,
                f"'ToC'!A2:K{len(toc_rows)+1}",
                toc_rows
            )
            print("✅ ToC updated successfully.")
        else:
            print("⚠️ No rows to insert into ToC.")

        # Optional POST request to web app
        try:
            web_app_url = 'https://script.google.com/macros/s/AKfycbzR3hWvfItvEOKjadlrVRx5vNTz4QH04WZbz2ufL8fAdbiZVsJbkzueKfmMCfGsAO62/exec'
            response = requests.post(web_app_url, json={'sheetUrl': spreadsheet_url})
            response.raise_for_status()
            print("✅ POST request sent to web app.")
        except Exception as post_err:
            print(f"⚠️ Error sending POST request: {post_err}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
