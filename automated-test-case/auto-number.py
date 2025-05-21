import os
import json
import re
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Parse env vars
sheet_data = json.loads(os.environ["SHEET_DATA"])
spreadsheet_url = sheet_data["spreadsheetUrl"]
spreadsheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url).group(1)

credentials_json = json.loads(os.environ["TEST_CASE_SERVICE_ACCOUNT_JSON"])

# Create credentials and sheets client
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
credentials = service_account.Credentials.from_service_account_info(credentials_json, scopes=SCOPES)
sheets_service = build('sheets', 'v4', credentials=credentials)

skip_sheets = ['ToC', 'Roster', 'Issues']

def update_values_with_retry(sheets, spreadsheet_id, range_name, values):
    attempts = 0
    while True:
        try:
            sheets.values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
            return
        except HttpError as error:
            if error.resp.status == 429:
                attempts += 1
                wait_time = (2 ** attempts)
                print(f"Quota exceeded. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise

def main():
    try:
        metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]

        for name in sheet_names:
            if name in skip_sheets:
                continue

            sheet_meta = next(s for s in sheets if s['properties']['title'] == name)
            sheet_id = sheet_meta['properties']['sheetId']
            merges = sheet_meta.get('merges', [])

            range_name = f"'{name}'!E12:F"
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=range_name).execute()
            rows = result.get('values', [])
            start_row = 12

            requests = []
            values = [[''] for _ in range(len(rows))]
            number = 1
            row = 0

            while row < len(rows):
                abs_row = row + start_row

                f_merge = next((
                    m for m in merges
                    if m['startRowIndex'] == abs_row - 1
                    and m['startColumnIndex'] == 5
                    and m['endColumnIndex'] == 6
                ), None)

                merge_start = abs_row
                merge_end = abs_row + 1

                if f_merge:
                    merge_start = f_merge['startRowIndex'] + 1
                    merge_end = f_merge['endRowIndex'] + 1

                is_merged_in_f = merge_end > merge_start
                merge_length = merge_end - merge_start

                f_value = rows[row][1].strip() if len(rows[row]) > 1 else ''
                e_value = rows[row][0].strip() if len(rows[row]) > 0 else ''

                e_merge = next((
                    m for m in merges
                    if m['startRowIndex'] == merge_start - 1
                    and m['endRowIndex'] == merge_end - 1
                    and m['startColumnIndex'] == 4
                    and m['endColumnIndex'] == 5
                ), None)

                if f_value:
                    values[row] = [str(number)]

                    if is_merged_in_f and not e_merge:
                        requests.append({
                            'mergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': merge_start - 1,
                                    'endRowIndex': merge_end - 1,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                },
                                'mergeType': 'MERGE_ALL'
                            }
                        })

                    if not is_merged_in_f and e_merge:
                        requests.append({
                            'unmergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': e_merge['startRowIndex'],
                                    'endRowIndex': e_merge['endRowIndex'],
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                }
                            }
                        })

                    number += 1

                row += merge_length

            update_range = f"'{name}'!E12:E{start_row + len(values) - 1}"
            update_values_with_retry(sheets_service, spreadsheet_id, update_range, values)

            if requests:
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={'requests': requests}
                ).execute()

            print(f"✅ Updated: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)

if __name__ == '__main__':
    main()
