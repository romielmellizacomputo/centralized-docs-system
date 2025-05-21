import os
import json
import re
import time

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables (assumes they're set in your environment)
sheet_data = json.loads(os.environ['SHEET_DATA'])
credentials_info = json.loads(os.environ['TEST_CASE_SERVICE_ACCOUNT_JSON'])

# Constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
skip_sheets = ['ToC', 'Roster', 'Issues']

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
            values = [[''] for _ in range(len(rows))]  # Clear all E column values
            number = 1
            row = 0

            while row < len(rows):
                abs_row = row + start_row

                def get_merge(col_index):
                    return next((m for m in merges if
                                 m['startRowIndex'] <= abs_row - 1 < m['endRowIndex'] and
                                 m['startColumnIndex'] == col_index and
                                 m['endColumnIndex'] == col_index + 1), None)

                e_merge = get_merge(4)
                f_merge = get_merge(5)

                merge_start = abs_row
                merge_end = abs_row + 1
                if f_merge:
                    merge_start = f_merge['startRowIndex'] + 1
                    merge_end = f_merge['endRowIndex'] + 1
                elif e_merge:
                    merge_start = e_merge['startRowIndex'] + 1
                    merge_end = e_merge['endRowIndex'] + 1

                merge_length = merge_end - merge_start
                first_row = merge_start - start_row

                e_value = rows[first_row][0].strip() if len(rows[first_row]) > 0 and rows[first_row][0] else ''
                f_value = rows[first_row][1].strip() if len(rows[first_row]) > 1 and rows[first_row][1] else ''

                has_e = bool(e_value)
                has_f = bool(f_value)

                # Case: F merged and empty -> unmerge F, border it
                if f_merge and not has_f:
                    requests.append({
                        'unmergeCells': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': f_merge['startRowIndex'],
                                'endRowIndex': f_merge['endRowIndex'],
                                'startColumnIndex': 5,
                                'endColumnIndex': 6,
                            }
                        }
                    })
                    # Border F cells
                    requests.append({
                        'updateBorders': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': f_merge['startRowIndex'],
                                'endRowIndex': f_merge['endRowIndex'],
                                'startColumnIndex': 5,
                                'endColumnIndex': 6,
                            },
                            'top':    {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            'bottom': {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            'left':   {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            'right':  {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                        }
                    })

                    # Then if E has data, unmerge and border, clear value
                    if e_merge:
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
                        requests.append({
                            'updateBorders': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': e_merge['startRowIndex'],
                                    'endRowIndex': e_merge['endRowIndex'],
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                },
                                'top':    {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'bottom': {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'left':   {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'right':  {'style': 'SOLID', 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            }
                        })

                        for i in range(merge_length):
                            values[row + i] = ['']

                # Case: E is empty but F has value -> assign number and merge E
                elif not has_e and has_f:
                    values[row] = [str(number)]
                    number += 1
                    if not e_merge:
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

                # Case: F is merged, E is not -> merge E and add number
                elif f_merge and not e_merge and has_f:
                    values[row] = [str(number)]
                    number += 1
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

                row += merge_length

            # Update values
            end_row = start_row + len(values) - 1
            value_range = f"'{name}'!E12:E{end_row}"
            update_values_with_retry(service, spreadsheet_id, value_range, values)

            if requests:
                body = {'requests': requests}
                service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

            print(f"✅ Updated: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)
