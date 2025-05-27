import os
import json
import re
import time
import datetime
import pytz 
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables (assumes they're set in your environment)
sheet_data = json.loads(os.environ['SHEET_DATA'])
credentials_info = json.loads(os.environ['TEST_CASE_SERVICE_ACCOUNT_JSON'])

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
skip_sheets = ['ToC', 'Roster', 'Issues', 'HELP']

def get_spreadsheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    raise ValueError('Invalid spreadsheet URL')

def update_values_with_retry(service, spreadsheet_id, range_, values):
    attempts = 0
    while True:
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
            return
        except HttpError as e:
            if e.resp.status == 429:
                attempts += 1
                wait_time = (2 ** attempts)
                print(f"Quota exceeded. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise

# Retry helper
def execute_with_retries(request_fn, max_retries=5, base_delay=5):
    for attempt in range(max_retries):
        try:
            return request_fn()
        except HttpError as e:
            if e.resp.status in [429, 500, 503]:
                wait_time = base_delay * (2 ** attempt)
                print(f"⚠️ Rate limited or server error. Retrying in {wait_time} seconds... (Attempt {attempt + 1})")
                time.sleep(wait_time)
            else:
                print("❌ Non-retryable error encountered:")
                raise
    raise Exception("❌ Exceeded maximum retries due to rate limits or server errors.")
    
def get_current_times():
    # Define timezones
    ph_tz = pytz.timezone('Asia/Manila')
    ug_tz = pytz.timezone('Africa/Kampala')
    # Get current UTC time
    now_utc = datetime.datetime.utcnow().replace(tzinfo=pytz.utc)
    # Convert to PH and UG timezones
    now_ph = now_utc.astimezone(ph_tz)
    now_ug = now_utc.astimezone(ug_tz)
    # Format the date and time nicely
    ph_time_str = now_ph.strftime('%B %d, %Y %I:%M %p')
    ug_time_str = now_ug.strftime('%B %d, %Y %I:%M %p')
    return ph_time_str, ug_time_str


def main():
    spreadsheet_url = sheet_data['spreadsheetUrl']
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)

    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    try:
        metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]

        for name in sheet_names:
            if name in skip_sheets:
                continue

            sheet_meta = next(s for s in sheets if s['properties']['title'] == name)
            sheet_id = sheet_meta['properties']['sheetId']
            merges = sheet_meta.get('merges', [])

            range_ = f"'{name}'!E12:F"
            result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
            rows = result.get('values', [])
            start_row = 12

            requests = []
            values = [[''] for _ in range(len(rows))]  # list of single-item lists
            number = 1
            row = 0

            while row < len(rows):
                abs_row = row + start_row

                f_merge = next((
                    m for m in merges
                    if m['startRowIndex'] == abs_row - 1 and
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

                e_merge = next((
                    m for m in merges
                    if m['startRowIndex'] == merge_start - 1 and
                       m['endRowIndex'] == merge_end - 1 and
                       m['startColumnIndex'] == 4 and
                       m['endColumnIndex'] == 5
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
                                'mergeType': 'MERGE_ALL',
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

            # Update values with retry
            end_row = start_row + len(values) - 1
            value_range = f"'{name}'!E12:E{end_row}"
            update_values_with_retry(service, spreadsheet_id, value_range, values)
            if requests:
                # Add a note to cell E10 (row 9 zero-based)
                ph_time_str, ug_time_str = get_current_times()
                note_text = (
                    f"This document was updated on {ph_time_str} (PH Time) / {ug_time_str} (UG Time). "
                    "The autonumbering and unmerging processes were managed by the Centralized Docs System, "
                    "leveraging integrations with GitLab, GitHub, Google API, and Google Apps Script. "
                    "This is a Milestone Project by Romiel Melliza Computo."
                )
                requests.append({
                    "updateCells": {
                        "rows": [{
                            "values": [{
                                "note": note_text
                            }]
                        }],
                        "fields": "note",
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 9,  # E10 is zero-based index 9
                            "endRowIndex": 10,
                            "startColumnIndex": 4,  # Column E → index 4
                            "endColumnIndex": 5
                        }
                    }
                })
                execute_with_retries(lambda: service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())

            print(f"✅ Updated: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)

if __name__ == '__main__':
    main()
