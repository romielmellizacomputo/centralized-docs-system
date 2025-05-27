from config import sheet_data, credentials_info
from constants import skip_sheets  # e.g. ['ToC', 'Roster', 'Issues', 'HELP']
from google_auth import get_sheet_service
from sheet_utils import get_spreadsheet_id
from retry_utils import execute_with_retries
import sys

ISSUES_SHEET = "Issues"
DROPDOWN_RANGE = "K3:K"

def main():
    spreadsheet_url = sheet_data['spreadsheetUrl']
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)

    try:
        # Get spreadsheet metadata
        metadata = execute_with_retries(
            lambda: service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        )
        sheets = metadata.get('sheets', [])
        sheet_name_to_gid = {}

        for s in sheets:
            name = s['properties']['title']
            gid = s['properties']['sheetId']
            if name not in skip_sheets:
                sheet_name_to_gid[name] = gid

        sheet_names = list(sheet_name_to_gid.keys())

        # Find the Issues sheet
        issues_sheet = next((s for s in sheets if s['properties']['title'] == ISSUES_SHEET), None)
        if not issues_sheet:
            raise ValueError(f'Sheet named "{ISSUES_SHEET}" not found.')

        issues_sheet_id = issues_sheet['properties']['sheetId']

        # 1. Set dropdown validation on K3:K
        requests = [{
            "setDataValidation": {
                "range": {
                    "sheetId": issues_sheet_id,
                    "startRowIndex": 2,
                    "startColumnIndex": 10,
                    "endColumnIndex": 11
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": name} for name in sheet_names]
                    },
                    "strict": True,
                    "showCustomUi": True
                }
            }
        }]

        execute_with_retries(
            lambda: service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()
        )
        print("✅ Dropdown updated successfully.")

        # 2. Fetch existing K3:K values
        get_range = f"{ISSUES_SHEET}!{DROPDOWN_RANGE}"
        values_res = execute_with_retries(
            lambda: service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=get_range
            ).execute()
        )

        base_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid="
        existing_values = values_res.get("values", [])

        # 3. Rewrite values with HYPERLINK formulas
        updated_values = []
        for row in existing_values:
            val = row[0].strip() if row else ''
            if val in sheet_name_to_gid:
                gid = sheet_name_to_gid[val]
                hyperlink = f'=HYPERLINK("{base_url}{gid}", "{val}")'
                updated_values.append([hyperlink])
            else:
                updated_values.append([val])

        # 4. Update values
        execute_with_retries(
            lambda: service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{ISSUES_SHEET}!K3",
                valueInputOption="USER_ENTERED",
                body={"values": updated_values}
            ).execute()
        )
        print("✅ Hyperlinks added successfully.")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
