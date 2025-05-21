from config import CONFIG, DASHBOARD_SHEET, generate_timestamp_string

def get_all_data(sheets, data_type, spreadsheet_id, utils_range=None):
    data_range = utils_range if utils_range else CONFIG[data_type]["range"]
    result = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=data_range
    ).execute()
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {data_range}")
    return values


def clear_target_sheet(sheets, sheet_id, data_type):
    sheet_name = CONFIG[data_type]["sheet_name"]
    max_length = CONFIG[data_type]["max_length"]
    # Calculate last column letter starting from 'C'
    col_end = chr(ord('C') + max_length - 1)
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{sheet_name}!C4:{col_end}"
    ).execute()

def pad_row(row, max_length):
    # Ensure the row has exactly max_length columns by padding with empty strings
    return row[:max_length] + [''] * (max_length - len(row))

def insert_data(sheets, sheet_id, data_type, data):
    sheet_name = CONFIG[data_type]["sheet_name"]
    max_length = CONFIG[data_type]["max_length"]
    padded_data = [pad_row(row, max_length) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {sheet_name}!C4")
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{sheet_name}!C4",
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    timestamp = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{DASHBOARD_SHEET}!W6",
        valueInputOption='RAW',
        body={'values': [[timestamp]]}
    ).execute()
