from config import sheet_data, credentials_info
from constants import SCOPES, skip_sheets
from config import sheet_data
from constants import skip_sheets
from google_auth import get_sheet_service
from sheet_utils import get_spreadsheet_id, create_border_request, get_sheet_metadata
from retry_utils import execute_with_retries

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

            sheet_meta = next(s for s in sheets if s['properties']['title'] == name)
            sheet_id = sheet_meta['properties']['sheetId']
            merges = sheet_meta.get('merges', [])

            requests = []

            for merge in merges:
                start_row_idx = merge['startRowIndex']
                end_row_idx = merge['endRowIndex']
                start_col_idx = merge['startColumnIndex']
                end_col_idx = merge['endColumnIndex']

                if start_row_idx < 11:
                    continue

                if start_col_idx == 4 and end_col_idx == 5:
                    e_range = f"'{name}'!E{start_row_idx+1}:E{end_row_idx}"
                    f_range = f"'{name}'!F{start_row_idx+1}:F{end_row_idx}"

                    e_values = execute_with_retries(lambda: service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id, range=e_range).execute().get('values', []))
                    f_values = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id, range=f_range).execute().get('values', [])

                    e_has_data = any(row and row[0].strip() for row in e_values)
                    f_has_data = any(row and row[0].strip() for row in f_values)

                    if e_has_data and not f_has_data:
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
                        requests.append(create_border_request(sheet_id, start_row_idx, end_row_idx, 4, 5))

                elif start_col_idx == 5 and end_col_idx == 6:
                    f_range = f"'{name}'!F{start_row_idx+1}:F{end_row_idx}"
                    f_values = service.spreadsheets().values().get(
                        spreadsheetId=spreadsheet_id, range=f_range).execute().get('values', [])
                    f_has_data = any(row and row[0].strip() for row in f_values)

                    # Range specs
                    e_range_spec = {
                        'sheetId': sheet_id,
                        'startRowIndex': start_row_idx,
                        'endRowIndex': end_row_idx,
                        'startColumnIndex': 4,
                        'endColumnIndex': 5,
                    }
                    f_range_spec = {
                        'sheetId': sheet_id,
                        'startRowIndex': start_row_idx,
                        'endRowIndex': end_row_idx,
                        'startColumnIndex': 5,
                        'endColumnIndex': 6,
                    }

                    e_merged = any(
                        m['startRowIndex'] == start_row_idx and
                        m['endRowIndex'] == end_row_idx and
                        m['startColumnIndex'] == 4 and
                        m['endColumnIndex'] == 5
                        for m in merges
                    )

                    if f_has_data:
                        if not e_merged:
                            requests.append({
                                'mergeCells': {
                                    'range': e_range_spec,
                                    'mergeType': 'MERGE_ALL'
                                }
                            })

                        # Add borders with explicit args instead of **rng
                        requests.append(create_border_request(
                            sheet_id=sheet_id,
                            start_row_idx=start_row_idx,
                            end_row_idx=end_row_idx,
                            start_col_idx=4,
                            end_col_idx=5
                        ))
                        requests.append(create_border_request(
                            sheet_id=sheet_id,
                            start_row_idx=start_row_idx,
                            end_row_idx=end_row_idx,
                            start_col_idx=5,
                            end_col_idx=6
                        ))

                    else:
                        if e_merged:
                            requests.append({
                                'unmergeCells': {
                                    'range': e_range_spec
                                }
                            })
                        requests.append({
                            'unmergeCells': {
                                'range': f_range_spec
                            }
                        })

                        # Add borders with explicit args instead of **rng
                        requests.append(create_border_request(
                            sheet_id=sheet_id,
                            start_row_idx=start_row_idx,
                            end_row_idx=end_row_idx,
                            start_col_idx=4,
                            end_col_idx=5
                        ))
                        requests.append(create_border_request(
                            sheet_id=sheet_id,
                            start_row_idx=start_row_idx,
                            end_row_idx=end_row_idx,
                            start_col_idx=5,
                            end_col_idx=6
                        ))

            if requests:
                execute_with_retries(lambda: service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id, body={'requests': requests}).execute())
                print(f"✅ Merge/unmerge updated for: {name}")
            else:
                print(f"⏭️ No merge/unmerge changes needed for: {name}")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        exit(1)

if __name__ == '__main__':
    main()
