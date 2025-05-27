import re
import time
from retry_utils import execute_with_retries, update_values_with_retry
from time_utils import get_current_times

def get_spreadsheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    raise ValueError('Invalid spreadsheet URL')

def get_sheet_metadata(service, spreadsheet_id):
    return execute_with_retries(lambda: service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute())

def process_sheet(service, spreadsheet_id, sheets, name):
    sheet_meta = next(s for s in sheets if s['properties']['title'] == name)
    sheet_id = sheet_meta['properties']['sheetId']
    merges = sheet_meta.get('merges', [])

    range_ = f"'{name}'!E12:F"
    time.sleep(1)
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_).execute()
    rows = result.get('values', [])
    start_row = 12

    requests = []
    values = [[''] for _ in range(len(rows))]
    number = 1
    row = 0

    while row < len(rows):
        abs_row = row + start_row

        f_merge = next((m for m in merges if m['startRowIndex'] == abs_row - 1 and m['startColumnIndex'] == 5 and m['endColumnIndex'] == 6), None)
        merge_start = abs_row
        merge_end = abs_row + 1

        if f_merge:
            merge_start = f_merge['startRowIndex'] + 1
            merge_end = f_merge['endRowIndex'] + 1

        is_merged_in_f = merge_end > merge_start
        merge_length = merge_end - merge_start

        f_value = rows[row][1].strip() if len(rows[row]) > 1 and rows[row][1] else None
        e_value = rows[row][0].strip() if len(rows[row]) > 0 and rows[row][0] else None

        e_merge = next((m for m in merges if m['startRowIndex'] == merge_start - 1 and m['endRowIndex'] == merge_end - 1 and m['startColumnIndex'] == 4 and m['endColumnIndex'] == 5), None)

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

    end_row = start_row + len(values) - 1
    value_range = f"'{name}'!E12:E{end_row}"
    update_values_with_retry(service, spreadsheet_id, value_range, values)

    if requests:
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
                    "startRowIndex": 9,
                    "endRowIndex": 10,
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                }
            }
        })

        execute_with_retries(lambda: service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())

    print(f"✅ Updated: {name}")



def get_spreadsheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    raise ValueError('Invalid spreadsheet URL')

def create_border_request(sheet_id, start_row_idx, end_row_idx, start_col_idx, end_col_idx):
    return {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row_idx,
                'endRowIndex': end_row_idx,
                'startColumnIndex': start_col_idx,
                'endColumnIndex': end_col_idx,
            },
            'top':    {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
            'bottom': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
            'left':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
            'right':  {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
            'innerHorizontal': {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
            'innerVertical':   {'style': 'SOLID', 'width': 1, 'color': {'red': 0, 'green': 0, 'blue': 0}},
        }
    }

