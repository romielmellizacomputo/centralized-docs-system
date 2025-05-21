from googleapiclient.discovery import build
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    NTC_SHEET,
    DASHBOARD_SHEET,
    CENTRAL_ISSUE_SHEET_ID,
    ALL_NTC,
    generate_timestamp_string,
)
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones,
)

def get_all_ntc(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=CENTRAL_ISSUE_SHEET_ID,
        range=ALL_NTC
    ).execute()

    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {ALL_NTC}")
    return values

def clear_ntc(sheets, sheet_id):
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4:U"
    ).execute()

def pad_row_to_u(row):
    full_length = 14
    return row + [''] * (full_length - len(row))

def insert_data_to_ntc(sheets, sheet_id, data):
    padded_data = [pad_row_to_u(row[:14]) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {NTC_SHEET}!C4")

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4",
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    formatted = generate_timestamp_string()

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{DASHBOARD_SHEET}!W6",
        valueInputOption='RAW',
        body={'values': [[formatted]]}
    ).execute()

def main():
    try:
        auth = authenticate()
        sheets = build('sheets', 'v4', credentials=auth)

        get_sheet_titles(sheets, UTILS_SHEET_ID)
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)

        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return

        required_labels = [
            'needs test case',
            'needs test scenario',
            'test case needs update',
        ]

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

                milestones, ntc_data = get_selected_milestones(sheets, sheet_id, G_MILESTONES), get_all_ntc(sheets)
                normalized_milestones = [m.lower().strip() for m in milestones]

                filtered = []
                for i, row in enumerate(ntc_data):
                    milestone_raw = row[6] if len(row) > 6 else ''
                    milestone = milestone_raw.lower().strip()

                    labels_raw = row[5] if len(row) > 5 else ''
                    labels = [label.strip().lower() for label in labels_raw.split(',')]

                    matches_milestone = milestone in normalized_milestones
                    has_relevant_label = any(label in required_labels for label in labels)

                    if matches_milestone and has_relevant_label:
                        print(f"✅ Row {i} MATCHES — Milestone: '{milestone_raw}', Labels: '{labels_raw}'")
                        filtered.append(row)
                    else:
                        reasons = []
                        if not matches_milestone:
                            reasons.append(f"milestone '{milestone_raw}' not matched")
                        if not has_relevant_label:
                            reasons.append(f"labels '{labels_raw}' missing relevant tags")
                        print(f"❌ Row {i} skipped — {', '.join(reasons)}")

                if not filtered:
                    print(f"ℹ️ No matching data found for {sheet_id}, skipping clear & insert.")
                    continue

                processed_data = [row[:21] for row in filtered]

                clear_ntc(sheets, sheet_id)
                insert_data_to_ntc(sheets, sheet_id, processed_data)
                update_timestamp(sheets, sheet_id)

                print(f"✅ Finished: {sheet_id}")

            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")

    except Exception as e:
        print(f"❌ Main failure: {str(e)}")

if __name__ == "__main__":
    main()
