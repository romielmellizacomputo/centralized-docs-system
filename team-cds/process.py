from googleapiclient.discovery import build
from common import authenticate, get_sheet_titles, get_all_team_cds_sheet_ids, get_selected_milestones
from sheets_utils import get_all_data, clear_target_sheet, insert_data, update_timestamp
from config import CONFIG, G_MILESTONES, UTILS_SHEET_ID  # Updated import

def process_data_type(data_type):
    try:
        creds = authenticate()
        sheets = build('sheets', 'v4', credentials=creds)

        # Validate the UTILS_SHEET_ID and fetch utility sheets once
        util_sheet_titles = get_sheet_titles(sheets, UTILS_SHEET_ID)

        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS sheet")
            return

        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")
                sheet_titles = get_sheet_titles(sheets, sheet_id)
                milestone_sheet = G_MILESTONES
                target_sheet = CONFIG[data_type]["sheet_name"]

                if milestone_sheet not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{milestone_sheet}' sheet")
                    continue
                if target_sheet not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{target_sheet}' sheet")
                    continue

                milestones = get_selected_milestones(sheets, sheet_id, milestone_sheet)

                # Updated to use UTILS_SHEET_ID and proper tab name
                all_data = get_all_data(
                    sheets,
                    data_type,
                    UTILS_SHEET_ID,
                    utils_range="Utils Sheet!B2:B10"  # <-- Replace with exact sheet tab name if different
                )

                label_index = CONFIG[data_type]["label_index"]

                if "filter_labels" in CONFIG[data_type]:
                    labels = CONFIG[data_type]["filter_labels"]
                    filtered = [row for row in all_data if len(row) > label_index and row[label_index] in labels]
                else:
                    filtered = [row for row in all_data if len(row) > label_index and row[label_index] in milestones]

                processed = [row for row in filtered]

                clear_target_sheet(sheets, sheet_id, data_type)
                insert_data(sheets, sheet_id, data_type, processed)
                update_timestamp(sheets, sheet_id)

                print(f"✅ Finished: {sheet_id}")

            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")
    except Exception as e:
        print(f"❌ Main failure: {str(e)}")
