import os
import json
import re
import time

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

sheet_data = json.loads(os.environ['SHEET_DATA'])
credentials_info = json.loads(os.environ['TEST_CASE_SERVICE_ACCOUNT_JSON'])

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
skip_sheets = ['ToC', 'Roster', 'Issues']

def get_spreadsheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    raise ValueError('Invalid spreadsheet URL')

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

            # Get data from columns E and F starting row 12
            range_ = f"'{name}'!E12:F"
            result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
            rows = result.get('values', [])
            start_row = 12

            requests = []

            for merge in merges:
                start_row_idx = merge['startRowIndex']
                end_row_idx = merge['endRowIndex']
                start_col_idx = merge['startColumnIndex']
                end_col_idx = merge['endColumnIndex']

                if start_row_idx < 11:  # skip merges above row 12
                    continue

                col_letter = chr(65 + start_col_idx)
                merge_range = f"{col_letter}{start_row_idx+1}:{chr(64 + end_col_idx)}{end_row_idx}"

                # Process column E merges (column index 4)
                if start_col_idx == 4 and end_col_idx == 5:
                    e_range = f"'{name}'!E{start_row_idx+1}:E{end_row_idx}"
                    f_range = f"'{name}'!F{start_row_idx+1}:F{end_row_idx}"
                    e_values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=e_range).execute().get('values', [])
                    f_values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f_range).execute().get('values', [])

                    e_has_data = any(row and row[0].strip() for row in e_values)
                    f_has_data = any(row and row[0].strip() for row in f_values)

                    if e_has_data and not f_has_data:
                        # Unmerge E
                        requests.append({
                            'unmergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row_idx,
                                    'endRowIndex': end_row_idx,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                }
                            }
                        })
                        # Clear E values
                        requests.append({
                            'updateCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row_idx,
                                    'endRowIndex': end_row_idx,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                },
                                'fields': 'userEnteredValue'
                            }
                        })
                        # Add black border
                        requests.append({
                            'updateBorders': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row_idx,
                                    'endRowIndex': end_row_idx,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5,
                                },
                                'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            }
                        })

                # Process column F merges (index 5)
                elif start_col_idx == 5 and end_col_idx == 6:
                    f_range = f"'{name}'!F{start_row_idx+1}:F{end_row_idx}"
                    f_values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f_range).execute().get('values', [])

                    f_has_data = any(row and row[0].strip() for row in f_values)

                    if not f_has_data:
                        # Unmerge F and add black border
                        requests.append({
                            'unmergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row_idx,
                                    'endRowIndex': end_row_idx,
                                    'startColumnIndex': 5,
                                    'endColumnIndex': 6,
                                }
                            }
                        })
                        requests.append({
                            'updateBorders': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row_idx,
                                    'endRowIndex': end_row_idx,
                                    'startColumnIndex': 5,
                                    'endColumnIndex': 6,
                                },
                                'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                            }
                        })

            if requests:
                body = {'requests': requests}
                service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                print(f"✅ Unmerging updated for: {name}")
            else:
                print(f"⏭️ No unmerging needed for: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)

if __name__ == '__main__':
    main()
