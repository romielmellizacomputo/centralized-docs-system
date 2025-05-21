import os
import re
import json
import time
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables as JSON strings
sheet_data = json.loads(os.getenv('SHEET_DATA'))
credentials_json = json.loads(os.getenv('TEST_CASE_SERVICE_ACCOUNT_JSON'))

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def exponential_backoff_retry(func, max_attempts=5):
    attempts = 0
    while True:
        try:
            return func()
        except HttpError as err:
            if err.resp.status == 429 and attempts < max_attempts:
                attempts += 1
                wait_time = (2 ** attempts)
                print(f"Quota exceeded. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise

def main():
    spreadsheet_url = sheet_data['spreadsheetUrl']
    # Extract spreadsheetId from URL
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', spreadsheet_url)
    if not match:
        raise ValueError("Invalid spreadsheet URL")
    spreadsheet_id = match.group(1)

    credentials = Credentials.from_service_account_info(credentials_json, scopes=SCOPES)
    sheets_service = build('sheets', 'v4', credentials=credentials)

    skip = ['ToC', 'Roster', 'Issues']

    try:
        metadata = exponential_backoff_retry(
            lambda: sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        )
        sheet_names = [s['properties']['title'] for s in metadata['sheets']]

        for name in sheet_names:
            if name in skip:
                continue

            sheet_meta = next(s for s in metadata['sheets'] if s['properties']['title'] == name)
            sheet_id = sheet_meta['properties']['sheetId']
            merges = sheet_meta.get('merges', [])

            # Fetch range E12:F (column 5 and 6) - columns are zero-based index in API, but range string is 1-based
            # Range string 'E12:F' means from E12 to F (all rows downward)
            range_str = f"'{name}'!E12:F"
            res = exponential_backoff_retry(
                lambda: sheets_service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=range_str).execute()
            )
            rows = res.get('values', [])
            start_row = 12

            requests = []
            values = [[''] for _ in range(len(rows))]
            number = 1
            row = 0

            while row < len(rows):
                abs_row = row + start_row  # absolute row number on the sheet

                # Find merge on column F (index 5)
                f_merge = next((
                    m for m in merges if
                    m['startRowIndex'] == abs_row - 1 and
                    m['startColumnIndex'] == 5 and
                    m['endColumnIndex'] == 6
                ), None)

                merge_start = abs_row
                merge_end = abs_row + 1

                if f_merge:
                    merge_start = f_merge['startRowIndex'] + 1
                    merge_end = f_merge['endRowIndex'] + 1

                is_merged_in_f = merge_end > merge_start
                merge_length = merge_end - merge_start

                f_value = rows[row][1].strip() if len(rows[row]) > 1 and rows[row][1] else None
                e_value = rows[row][0].strip() if len(rows[row]) > 0 and rows[row][0] else None

                # Find merge on column E (index 4) in the same merged rows
                e_merge = next((
                    m for m in merges if
                    m['startRowIndex'] == merge_start - 1 and
                    m['endRowIndex'] == merge_end - 1 and
                    m['startColumnIndex'] == 4 and
                    m['endColumnIndex'] == 5
                ), None)

                if f_value:
                    values[row] = [str(number)]

                    if is_merged_in_f and not e_merge:
                        requests.append({
                            "mergeCells": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": merge_start - 1,
                                    "endRowIndex": merge_end - 1,
                                    "startColumnIndex": 4,
                                    "endColumnIndex": 5
                                },
                                "mergeType": "MERGE_ALL"
                            }
                        })

                    if not is_merged_in_f and e_merge:
                        requests.append({
                            "unmergeCells": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": e_merge['startRowIndex'],
                                    "endRowIndex": e_merge['endRowIndex'],
                                    "startColumnIndex": 4,
                                    "endColumnIndex": 5
                                }
                            }
                        })

                    number += 1

                row += merge_length

            # Update values in column E starting at row 12
            update_range = f"'{name}'!E12:E{start_row + len(values) - 1}"

            def update_values():
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=update_range,
                    valueInputOption='USER_ENTERED',
                    body={'values': values}
                ).execute()

            exponential_backoff_retry(update_values)

            if requests:
                batch_update_body = {"requests": requests}
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=batch_update_body
                ).execute()

            print(f"✅ Updated: {name}")

    except Exception as err:
        print(f"❌ ERROR: {err}")
        exit(1)

if __name__ == "__main__":
    main()
