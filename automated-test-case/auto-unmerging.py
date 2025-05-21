import os
import json
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

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

            requests = []

            for merge in merges:
                start_row = merge['startRowIndex']
                end_row = merge['endRowIndex']
                start_col = merge['startColumnIndex']
                end_col = merge['endColumnIndex']

                if start_row < 11:  # only process rows from 12 (index 11)
                    continue

                # Handle E column (index 4)
                if start_col == 4 and end_col == 5:
                    row_range = f"{name}!E{start_row+1}:E{end_row}"
                    values = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id, range=row_range).execute().get('values', [])

                    has_data = any(val and val[0].strip() for val in values)

                    if not has_data:
                        # Unmerge
                        requests.append({
                            'unmergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row,
                                    'endRowIndex': end_row,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5
                                }
                            }
                        })
                        # Clear values
                        requests.append({
                            'updateCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row,
                                    'endRowIndex': end_row,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5
                                },
                                'fields': 'userEnteredValue'
                            }
                        })
                        # Add borders
                        requests.append({
                            'updateBorders': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row,
                                    'endRowIndex': end_row,
                                    'startColumnIndex': 4,
                                    'endColumnIndex': 5
                                },
                                'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}}
                            }
                        })

                # Handle F column (index 5)
                elif start_col == 5 and end_col == 6:
                    row_range = f"{name}!F{start_row+1}:F{end_row}"
                    values = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id, range=row_range).execute().get('values', [])

                    has_data = any(val and val[0].strip() for val in values)

                    if not has_data:
                        # Unmerge
                        requests.append({
                            'unmergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row,
                                    'endRowIndex': end_row,
                                    'startColumnIndex': 5,
                                    'endColumnIndex': 6
                                }
                            }
                        })
                        # Add borders
                        requests.append({
                            'updateBorders': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': start_row,
                                    'endRowIndex': end_row,
                                    'startColumnIndex': 5,
                                    'endColumnIndex': 6
                                },
                                'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
                                'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}}
                            }
                        })

            if requests:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()
                print(f"✅ Unmerging updated for: {name}")
            else:
                print(f"⏭️ No unmerging needed for: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)

if __name__ == '__main__':
    main()
