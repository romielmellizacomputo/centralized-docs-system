import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    NTC_SHEET,
    DASHBOARD_SHEET,
    SHEET_SYNC_SID,  # Updated to use SHEET_SYNC_SID
    generate_timestamp_string,
)
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones,
)

def get_all_ntc(sheets):
    """Get all issues from ALL ISSUES sheet (SHEET_SYNC_SID) for NTC filtering"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL ISSUES!C4:T"
    ).execute()

    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range ALL ISSUES!C4:T")
    return values

def clear_ntc(sheets, sheet_id):
    """Clear existing data in NTC sheet"""
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4:N"
    ).execute()

def pad_row_to_n(row):
    """Pad row to 12 columns (C to N = 12 columns)"""
    full_length = 12
    return row + [''] * (full_length - len(row))

def insert_data_to_ntc(sheets, sheet_id, data):
    """Insert processed data to NTC sheet"""
    padded_data = [pad_row_to_n(row[:12]) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {NTC_SHEET}!C4")

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{NTC_SHEET}!C4",
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    """Update timestamp in Dashboard sheet"""
    formatted = generate_timestamp_string()

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{DASHBOARD_SHEET}!W6",
        valueInputOption='RAW',
        body={'values': [[formatted]]}
    ).execute()

def debug_ntc_filtering(ntc_data, milestones):
    """Debug function to show NTC filtering details"""
    print("🔍 DEBUG: NTC Filtering Analysis...")
    print(f"🔍 Column mapping from ALL ISSUES C4:T:")
    print(f"   - Column H (index 5): Labels")
    print(f"   - Column I (index 6): Milestone")
    
    if not ntc_data:
        print("🔍 No source data found!")
        return
    
    # Show first few rows
    print(f"\n🔍 First 3 rows sample:")
    for i, row in enumerate(ntc_data[:3], 1):
        labels = row[5] if len(row) > 5 else 'N/A'
        milestone = row[6] if len(row) > 6 else 'N/A'
        print(f"  Row {i}: Labels='{labels}' | Milestone='{milestone}'")
    
    # Get unique milestones
    source_milestones = set()
    for row in ntc_data:
        if len(row) > 6 and row[6] and str(row[6]).strip():
            source_milestones.add(str(row[6]).strip().lower())
    
    print(f"\n🔍 Found {len(source_milestones)} unique milestones")
    print(f"🔍 First 10 source milestones: {list(source_milestones)[:10]}")
    
    # Check matches
    normalized_milestones = set(m.lower().strip() for m in milestones)
    exact_matches = source_milestones.intersection(normalized_milestones)
    print(f"🔍 Milestone matches: {len(exact_matches)} - {list(exact_matches)[:5]}")
    print(f"🔍 Target milestones (first 5): {milestones[:5]}")

def filter_ntc_data(ntc_data, milestones, required_labels):
    """Filter NTC data by milestones and labels"""
    if not ntc_data:
        return []
    
    debug_ntc_filtering(ntc_data, milestones)
    
    filtered = []
    normalized_milestones = [m.lower().strip() for m in milestones]
    
    print(f"\n📋 Required labels: {required_labels}")
    print(f"📋 Filtering {len(ntc_data)} rows...\n")

    for i, row in enumerate(ntc_data, 1):
        # Column I (index 6) = Milestone from ALL ISSUES C4:T
        milestone_raw = row[6] if len(row) > 6 else ''
        milestone = milestone_raw.lower().strip()

        # Column H (index 5) = Labels from ALL ISSUES C4:T
        labels_raw = row[5] if len(row) > 5 else ''
        labels = [label.strip().lower() for label in labels_raw.split(',')]

        matches_milestone = milestone in normalized_milestones
        has_relevant_label = any(label in required_labels for label in labels)

        if matches_milestone and has_relevant_label:
            print(f"✅ Row {i} MATCHES — Milestone: '{milestone_raw}', Labels: '{labels_raw}'")
            filtered.append(row)
        elif i <= 10:  # Show first 10 non-matches for debugging
            reasons = []
            if not matches_milestone:
                reasons.append(f"milestone '{milestone_raw}' not matched")
            if not has_relevant_label:
                reasons.append(f"labels '{labels_raw}' missing relevant tags")
            print(f"❌ Row {i} skipped — {', '.join(reasons)}")

    print(f"\n📊 Filtered {len(filtered)} rows from {len(ntc_data)} total rows")
    return filtered

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

                # Get milestones and NTC data
                print(f"📋 Getting milestones from {sheet_id} - {G_MILESTONES}")
                milestones = get_selected_milestones(sheets, sheet_id, G_MILESTONES)
                print(f"📋 Found {len(milestones)} milestones")
                
                if not milestones:
                    print(f"⚠️ No milestones found for {sheet_id}, skipping...")
                    continue
                
                # Get all issues from SHEET_SYNC_SID - ALL ISSUES sheet
                print(f"📋 Getting issues from {SHEET_SYNC_SID} - ALL ISSUES!C4:T")
                ntc_data = get_all_ntc(sheets)
                print(f"📋 Found {len(ntc_data)} total rows")

                # Filter by milestones and required labels
                filtered = filter_ntc_data(ntc_data, milestones, required_labels)

                if not filtered:
                    print(f"⚠️ No matching NTC data found for {sheet_id}")
                    # Still clear the sheet and update timestamp
                    clear_ntc(sheets, sheet_id)
                    update_timestamp(sheets, sheet_id)
                    continue

                print(f"📊 Processing {len(filtered)} filtered NTC rows")

                # Clear existing data and insert new data
                clear_ntc(sheets, sheet_id)
                insert_data_to_ntc(sheets, sheet_id, filtered)
                update_timestamp(sheets, sheet_id)

                print(f"✅ Finished: {sheet_id}")

            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")
                import traceback
                traceback.print_exc()

    except Exception as e:
        print(f"❌ Main failure: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
