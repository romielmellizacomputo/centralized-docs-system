from googleapiclient.discovery import build
from google.oauth2 import service_account
from ..constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_MR_SHEET,
    DASHBOARD_SHEET,
    CENTRAL_ISSUE_SHEET_ID,
    ALL_MR,
    generate_timestamp_string
)
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones
)

def get_all_mr(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=CENTRAL_ISSUE_SHEET_ID,
        range=ALL_MR
    ).execute()

    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {ALL_MR}")
    return values

def clear_gmr(sheets, sheet_id):
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f'{G_MR_SHEET}!C4:S'
    ).execute()

def pad_row_to_u(row):
    full_length = 17
    return row + [''] * (full_length - len(row))

def insert_data_to_gmr(sheets, sheet_id, data):
    padded_data = [pad_row_to_u(row[:17]) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_MR_SHEET}!C4")
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{G_MR_SHEET}!C4',
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    formatted = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{DASHBOARD_SHEET}!W6',
        valueInputOption='RAW',
        body={'values': [[formatted]]}
    ).execute()

def main():
    try:
        creds = authenticate()
        sheets = build('sheets', 'v4', credentials=creds)

        get_sheet_titles(sheets, UTILS_SHEET_ID)
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)

        if not sheet_ids:
            print('❌ No Team CDS sheet IDs found in UTILS!B2:B')
            return

        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")

                sheet_titles = get_sheet_titles(sheets, sheet_id)

                if G_MILESTONES not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MILESTONES}' sheet")
                    continue

                if G_MR_SHEET not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MR_SHEET}' sheet")
                    continue

                milestones, issues_data = (
                    get_selected_milestones(sheets, sheet_id, G_MILESTONES),
                    get_all_mr(sheets)
                )

                filtered = [row for row in issues_data if len(row) > 7 and row[7] in milestones]
                processed_data = [row[:17] for row in filtered]

                clear_gmr(sheets, sheet_id)
                insert_data_to_gmr(sheets, sheet_id, processed_data)
                update_timestamp(sheets, sheet_id)

                print(f"✅ Finished: {sheet_id}")
            except Exception as err:
                print(f"❌ Error processing {sheet_id}: {err}")
    except Exception as err:
        print(f"❌ Main failure: {err}")

if __name__ == '__main__':
    main()
