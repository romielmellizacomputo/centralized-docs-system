import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_MR_SHEET,
    DASHBOARD_SHEET,
    SHEET_SYNC_SID,  # Updated to use SHEET_SYNC_SID
    generate_timestamp_string
)
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones,
)

def get_all_mr(sheets):
    """Get all MRs from ALL MRs sheet (SHEET_SYNC_SID)"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL MRs!C4:S"
    ).execute()
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range ALL MRs!C4:S")
    return values

def clear_gmr(sheets, sheet_id):
    """Clear existing data in G-MR sheet"""
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f"{G_MR_SHEET}!C4:S"
    ).execute()

def pad_row_to_s(row):
    """Pad row to 17 columns (C to S = 17 columns)"""
    full_length = 17
    return row + [''] * (full_length - len(row))

def insert_data_to_gmr(sheets, sheet_id, data):
    """Insert processed data to G-MR sheet"""
    padded_data = [pad_row_to_s(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_MR_SHEET}!C4")
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{G_MR_SHEET}!C4",
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

def debug_milestone_matching(mr_data, milestones):
    """Debug function to help understand milestone matching"""
    print("🔍 DEBUG: Analyzing milestone data in column J (index 7)...")
    
    if not mr_data:
        print("🔍 No source data found!")
        return
    
    # Get unique milestones from source data (column J, index 7)
    source_milestones = set()
    for row in mr_data:
        if len(row) > 7 and row[7] and str(row[7]).strip():  # Column J is index 7
            source_milestones.add(str(row[7]).strip())
    
    print(f"🔍 Found {len(source_milestones)} unique milestones in column J")
    print(f"🔍 First 10 source milestones: {list(source_milestones)[:10]}")
    
    # Check for exact matches
    milestone_set = set(milestones)
    exact_matches = source_milestones.intersection(milestone_set)
    print(f"🔍 Exact matches: {len(exact_matches)} - {list(exact_matches)[:5]}")
    
    # Show target milestones for comparison
    print(f"🔍 Target milestones (first 5): {milestones[:5]}")

def filter_mrs_by_milestones(mr_data, milestones):
    """Filter MRs by milestones using column J (index 7)"""
    if not mr_data:
        return []
    
    # Debug milestone matching
    debug_milestone_matching(mr_data, milestones)
    
    filtered = []
    milestone_set = set(milestones)
    milestone_col_idx = 7  # Column J is index 7 (C=0, D=1, E=2, F=3, G=4, H=5, I=6, J=7)
    
    # Process data rows (no header in C4:S range)
    for i, row in enumerate(mr_data, 1):
        if len(row) <= milestone_col_idx:  # Row doesn't have enough columns
            continue
            
        milestone = str(row[milestone_col_idx]).strip() if row[milestone_col_idx] else ""
        
        if milestone in milestone_set:
            filtered.append(row)
            if len(filtered) <= 5:  # Show first 5 matches for debugging
                print(f"✅ Match found at row {i}: milestone='{milestone}'")
    
    print(f"📊 Filtered {len(filtered)} MRs from {len(mr_data)} total MRs using column J")
    return filtered

def main():
    try:
        creds = authenticate()
        sheets = build('sheets', 'v4', credentials=creds)
        get_sheet_titles(sheets, UTILS_SHEET_ID)
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        
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
                
                if G_MR_SHEET not in sheet_titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MR_SHEET}' sheet")
                    continue
                
                # Get selected milestones for filtering
                print(f"📋 Getting milestones from {sheet_id} - {G_MILESTONES}")
                milestones = get_selected_milestones(sheets, sheet_id, G_MILESTONES)
                print(f"📋 Found {len(milestones)} milestones")
                
                if not milestones:
                    print(f"⚠️ No milestones found for {sheet_id}, skipping...")
                    continue
                
                # Get all MRs from SHEET_SYNC_SID - ALL MRs sheet
                print(f"📋 Getting MRs from {SHEET_SYNC_SID} - ALL MRs!C4:S")
                mr_data = get_all_mr(sheets)
                print(f"📋 Found {len(mr_data)} total rows")
                
                # Filter by milestones using column J (index 7)
                filtered = filter_mrs_by_milestones(mr_data, milestones)
                
                if not filtered:
                    print(f"⚠️ No matching MRs found for {sheet_id}")
                    # Still clear the sheet and update timestamp
                    clear_gmr(sheets, sheet_id)
                    update_timestamp(sheets, sheet_id)
                    continue
                
                print(f"📊 Processing {len(filtered)} filtered MRs")
                
                # Clear existing data and insert new data
                clear_gmr(sheets, sheet_id)
                insert_data_to_gmr(sheets, sheet_id, filtered)
                update_timestamp(sheets, sheet_id)
                
                print(f"✅ Finished: {sheet_id}")
                
            except Exception as err:
                print(f"❌ Error processing {sheet_id}: {str(err)}")
                import traceback
                traceback.print_exc()
                
    except Exception as err:
        print(f"❌ Main failure: {str(err)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
