from config import sheet_data, credentials_info
from constants import skip_sheets
from google_auth import get_sheet_service
from sheet_utils import (
    get_spreadsheet_id,
    get_sheet_metadata,
)
from retry_utils import execute_with_retries
import sys

TOC_SHEET_NAME = "ToC"
TOC_RANGE_A = f"{TOC_SHEET_NAME}!A2:A"
TOC_RANGE_BK = f"{TOC_SHEET_NAME}!B2:K"

def main():
    spreadsheet_url = sheet_data["spreadsheetUrl"]
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)

    try:
        metadata = get_sheet_metadata(service, spreadsheet_id)
        sheets = metadata.get("sheets", [])
        toc_sheet = next((s for s in sheets if s["properties"]["title"] == TOC_SHEET_NAME), None)

        if not toc_sheet:
            print("❌ ToC sheet not found.")
            sys.exit(1)

        # Clear old values in A2:A and B2:K
        execute_with_retries(
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id, range=TOC_RANGE_A
            )
        )
        execute_with_retries(
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id, range=TOC_RANGE_BK
            )
        )

        toc_rows = []
        sheet_names_seen = set()

        for sheet in sheets:
            name = sheet["properties"]["title"]
            if name in skip_sheets or name == TOC_SHEET_NAME:
                continue

            # Read C4
            c4_range = f"'{name}'!C4"
            c4_result = execute_with_retries(
                service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id, range=c4_range
                )
            )
            c4_value = c4_result.get("values", [[None]])[0][0]

            if not c4_value or c4_value in sheet_names_seen:
                continue

            sheet_names_seen.add(c4_value)
            sheet_id = sheet["properties"]["sheetId"]
            hyperlink = f'=HYPERLINK("{spreadsheet_url}#gid={sheet_id}", "{c4_value}")'

            # Read other fields
            other_cells = ["C5", "C7", "C15", "C18", "C19", "C20", "C21", "C14", "C13", "C6"]
            batch_ranges = [f"'{name}'!{cell}" for cell in other_cells]
            data_result = execute_with_retries(
                service.spreadsheets().values().batchGet(
                    spreadsheetId=spreadsheet_id, ranges=batch_ranges
                )
            )
            other_values = [r.get("values", [[None]])[0][0] if r.get("values") else "" for r in data_result["valueRanges"]]

            toc_rows.append([hyperlink] + other_values)
            print(f"✅ Added entry for: {c4_value}")

        if toc_rows:
            execute_with_retries(
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"{TOC_SHEET_NAME}!A2:K{len(toc_rows)+1}",
                    valueInputOption="USER_ENTERED",
                    body={"values": toc_rows},
                )
            )
            print("✅ ToC sheet updated successfully.")
        else:
            print("⚠️ No valid entries to add.")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
