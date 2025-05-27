from config import sheet_data, credentials_info
from google_auth import get_sheet_service
from constants import skip_sheets
from retry_utils import execute_with_retries
from sheet_utils import get_spreadsheet_id
import sys

def main():
    spreadsheet_url = sheet_data["spreadsheetUrl"]
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)

    TOC_NAME = "ToC"
    print("📄 Spreadsheet ID:", spreadsheet_id)

    try:
        metadata = execute_with_retries(service.spreadsheets().get(spreadsheetId=spreadsheet_id))
        sheets = metadata["sheets"]

        toc_sheet = next((s for s in sheets if s["properties"]["title"] == TOC_NAME), None)
        if not toc_sheet:
            print("❌ ToC sheet not found.")
            return

        print("✅ Clearing existing ToC values...")
        execute_with_retries(service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id, range=f"{TOC_NAME}!A2:K"
        ))

        rows = []
        seen = set()

        for sheet in sheets:
            title = sheet["properties"]["title"]
            if title in skip_sheets or title == TOC_NAME:
                continue

            c4_range = f"'{title}'!C4"
            result = execute_with_retries(service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=c4_range
            ))
            c4_value = result.get("values", [[None]])[0][0]

            if not c4_value or c4_value in seen:
                continue

            seen.add(c4_value)
            sheet_id = sheet["properties"]["sheetId"]
            hyperlink = f'=HYPERLINK("{spreadsheet_url}#gid={sheet_id}", "{c4_value}")'
            dummy_data = [f"Dummy {i}" for i in range(1, 11)]
            rows.append([hyperlink] + dummy_data)

            print(f"✅ Added row for {c4_value}")

        if rows:
            print("📥 Writing new rows...")
            execute_with_retries(service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{TOC_NAME}!A2:K{len(rows)+1}",
                valueInputOption="USER_ENTERED",
                body={"values": rows},
            ))
            print("🎉 Done updating ToC.")
        else:
            print("⚠️ No valid rows found.")

    except Exception as e:
        print("❌ Error:", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
