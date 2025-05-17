import os
import json
import pytz
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

UTILS_SHEET_ID = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k'
G_MILESTONES = 'G-Milestones'
NTC_SHEET = 'NTC'
DASHBOARD_SHEET = 'Dashboard'

CENTRAL_ISSUE_SHEET_ID = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY'
ALL_ISSUES_RANGE = 'ALL ISSUES!C4:N'

def authenticate():
    credentials_info = json.loads(os.environ['TEAM_CDS_SERVICE_ACCOUNT_JSON'])
    creds = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return build('sheets', 'v4', credentials=creds)

def get_sheet_titles(sheets, spreadsheet_id):
    res = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [sheet['properties']['title'] for sheet in res['sheets']]
    print(f"📄 Sheets in {spreadsheet_id}:", titles)
    return titles

def get_all_team_cds_sheet_ids(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=UTILS_SHEET_ID,
        range='UTILS!B2:B'
    ).execute()
    values = result.get('values', [])
    return [row[0] for row in values if row]

def get_selected_milestones(sheets, sheet_id):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"{G_MILESTONES}!G4:G"
    ).execute()
    values = result.get('values', [])
    return [row[0] for row in values if row]

def get_all_issues(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=CENTRAL_ISSUE_SHEET_ID,
        range=ALL_ISSUES_RANGE
    ).execute()
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {ALL_ISSUES_RANGE}")
    return values

def clear_ntc_sheet(sheets, sheet_id):
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4:N",
        body={}
    ).execute()

def insert_data_to_ntc_sheet(sheets, sheet_id, data):
    if not data:
        print("No data to insert.")
        return
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4",
        valueInputOption='RAW',
        body={'values': data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    now_utc = datetime.utcnow().replace(tzinfo=pytz.utc)

    tz_eat = pytz.timezone('Africa/Nairobi')
    tz_pht = pytz.timezone('Asia/Manila')

    now_eat = now_utc.astimezone(tz_eat)
    now_pht = now_utc.astimezone(tz_pht)

    date_eat = now_eat.strftime('%B %d, %Y')
    time_eat = now_eat.strftime('%I:%M:%S %p')

    date_pht = now_pht.strftime('%B %d, %Y')
    time_pht = now_pht.strftime('%I:%M:%S %p')

    formatted = f"Sync on {date_eat}, {time_eat} (EAT) / {date_pht}, {time_pht} (PHT)"

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{DASHBOARD_SHEET}!W6",
        valueInputOption='RAW',
        body={'values': [[formatted]]}
    ).execute()

def main():
    try:
        sheets = authenticate()

        get_sheet_titles(sheets, UTILS_SHEET_ID)

        sheet_ids = get_all_team_cds_sheet_ids(sheets)
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return

        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")
                sheet_titles = get_sheet_titles(sheets, sheet_id)

                if G_MILESTONES not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MILESTONES}' sheet")
                    continue

                if NTC_SHEET not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{NTC_SHEET}' sheet")
                    continue

                milestones = get_selected_milestones(sheets, sheet_id)
                issues_data = get_all_issues(sheets)

                filtered = []
                for row in issues_data:
                    milestone_matches = row[6] in milestones if len(row) > 6 else False
                    labels_raw = row[5] if len(row) > 5 else ''
                    labels = [label.strip().lower() for label in labels_raw.split(',')]

                    print(f"Raw labels for row: {labels_raw}")
                    print(f"Processed labels for row: {labels}")

                    labels_match = any(label in [
                        "needs test case", "needs test scenario", "test case needs update"
                    ] for label in labels)

                    if milestone_matches and labels_match:
                        filtered.append(row[:12])

                if filtered:
                    clear_ntc_sheet(sheets, sheet_id)
                    insert_data_to_ntc_sheet(sheets, sheet_id, filtered)
                    update_timestamp(sheets, sheet_id)
                    print(f"✅ Finished: {sheet_id}")
                else:
                    print(f"⚠️ No matching data for {sheet_id}")

            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")

    except Exception as e:
        print(f"❌ Main failure: {str(e)}")

if __name__ == "__main__":
    main()
