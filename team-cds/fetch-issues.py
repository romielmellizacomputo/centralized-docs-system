import sys
import os
import json
import asyncio  # Add this line to import asyncio

# Add the root directory to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from googleapiclient.discovery import build
from google.oauth2 import service_account  # Import service_account for credentials
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_ISSUES_SHEET,
    DASHBOARD_SHEET,
    CENTRAL_ISSUE_SHEET_ID,
    ALL_ISSUES,
    generate_timestamp_string
)
from common import (
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones,
)

async def get_all_issues(sheets):
    result = await sheets.spreadsheets().values().get(
        spreadsheetId=CENTRAL_ISSUE_SHEET_ID,
        range=ALL_ISSUES,
    ).execute()

    data = result.get('values', [])
    
    if not data:
        raise ValueError(f"No data found in range {ALL_ISSUES}")

    return data

async def clear_g_issues(sheets, sheet_id):
    await sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{G_ISSUES_SHEET}!C4:T",
    ).execute()

def pad_row_to_u(row):
    full_length = 18
    return row + [''] * (full_length - len(row))

async def insert_data_to_g_issues(sheets, sheet_id, data):
    padded_data = [pad_row_to_u(row[:18]) for row in data]

    print(f"📤 Inserting {len(padded_data)} rows to {G_ISSUES_SHEET}!C4")
    await sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{G_ISSUES_SHEET}!C4",
        valueInputOption='RAW',
        body={'values': padded_data},
    ).execute()

async def update_timestamp(sheets, sheet_id):
    formatted = generate_timestamp_string()

    await sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{DASHBOARD_SHEET}!W6",
        valueInputOption='RAW',
        body={'values': [[formatted]]},
    ).execute()

async def authenticate():
    # Load credentials from the environment variable
    service_account_info = json.loads(os.environ['TEAM_CDS_SERVICE_ACCOUNT_JSON'])
    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    return credentials

async def main():
    try:
        auth = await authenticate()
        sheets = build('sheets', 'v4', credentials=auth)

        await get_sheet_titles(sheets, UTILS_SHEET_ID)

        sheet_ids = await get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        if not sheet_ids:
            print('❌ No Team CDS sheet IDs found in UTILS!B2:B')
            return

        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")

                sheet_titles = await get_sheet_titles(sheets, sheet_id)

                if G_MILESTONES not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MILESTONES}' sheet")
                    continue

                if G_ISSUES_SHEET not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_ISSUES_SHEET}' sheet")
                    continue

                milestones, issues_data = await asyncio.gather(
                    get_selected_milestones(sheets, sheet_id, G_MILESTONES),
                    get_all_issues(sheets),
                )

                filtered = [row for row in issues_data if row[6] in milestones]  # Column I
                processed_data = [row[:18] for row in filtered]

                await clear_g_issues(sheets, sheet_id)
                await insert_data_to_g_issues(sheets, sheet_id, processed_data)
                await update_timestamp(sheets, sheet_id)

                print(f"✅ Finished: {sheet_id}")
            except Exception as err:
                print(f"❌ Error processing {sheet_id}: {err}")

    except Exception as err:
        print(f"❌ Main failure: {err}")

# Run the main function
if __name__ == "__main__":
    asyncio.run(main())
