import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from googleapiclient.discovery import build
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones
)
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_ISSUES_SHEET,
    DASHBOARD_SHEET,
    CBS_ID,  # Updated to use CBS_ID instead of CENTRAL_ISSUE_SHEET_ID
    GITLAB_ISSUES,  # Updated to use GITLAB_ISSUES instead of ALL_ISSUES
    generate_timestamp_string
)

def get_all_issues(sheets):
    """Get all issues from the new GitLab Issues source (CBS_ID)"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=CBS_ID,
        range=GITLAB_ISSUES
    ).execute()
    
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {GITLAB_ISSUES}")
    
    return values

def clear_g_issues(sheets, sheet_id):
    """Clear existing data in G-Issues sheet"""
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f'{G_ISSUES_SHEET}!C4:U'
    ).execute()

def pad_row_to_u(row):
    """Pad row to 19 columns (C to U = 19 columns)"""
    full_length = 19
    return row + [''] * (full_length - len(row))

def insert_data_to_g_issues(sheets, sheet_id, data):
    """Insert processed data to G-Issues sheet"""
    padded_data = [pad_row_to_u(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_ISSUES_SHEET}!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{G_ISSUES_SHEET}!C4',
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    """Update timestamp in Dashboard sheet"""
    timestamp = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{DASHBOARD_SHEET}!W6',
        valueInputOption='RAW',
        body={'values': [[timestamp]]}
    ).execute()

def debug_milestone_matching(issues_data, milestones):
    """Debug function to help understand milestone matching issues"""
    print("🔍 DEBUG: Analyzing milestone matching...")
    
    # Get unique milestones from source data (column F, index 5)
    source_milestones = set()
    for row in issues_data[1:]:  # Skip header row
        if len(row) > 5 and row[5].strip():  # Check if milestone column exists and is not empty
            source_milestones.add(row[5].strip())
    
    print(f"🔍 Found {len(source_milestones)} unique milestones in source data")
    print(f"🔍 First 10 source milestones: {list(source_milestones)[:10]}")
    
    # Check for exact matches
    exact_matches = source_milestones.intersection(set(milestones))
    print(f"🔍 Exact matches: {len(exact_matches)} - {list(exact_matches)[:5]}")
    
    # Check for partial matches (case-insensitive)
    partial_matches = []
    milestone_lower = [m.lower() for m in milestones]
    for source_milestone in source_milestones:
        for target_milestone in milestones:
            if (source_milestone.lower() in target_milestone.lower() or 
                target_milestone.lower() in source_milestone.lower()):
                partial_matches.append((source_milestone, target_milestone))
                break
    
    print(f"🔍 Partial matches found: {len(partial_matches)}")
    if partial_matches:
        print(f"🔍 First 5 partial matches: {partial_matches[:5]}")

def filter_issues_by_milestones(issues_data, milestones):
    """Filter issues by milestones with improved matching logic"""
    if not issues_data:
        return []
    
    # Debug milestone matching
    debug_milestone_matching(issues_data, milestones)
    
    filtered = []
    milestone_set = set(milestones)
    
    # Skip header row (index 0) and process data rows
    for i, row in enumerate(issues_data[1:], 1):  # Start from row 1, but count from 1
        if len(row) <= 5:  # Row doesn't have enough columns
            continue
            
        milestone = row[5].strip() if row[5] else ""  # Column F (index 5) is milestone
        
        if milestone in milestone_set:
            filtered.append(row)
            if len(filtered) <= 5:  # Show first 5 matches for debugging
                print(f"✅ Match found at row {i}: milestone='{milestone}'")
    
    print(f"📊 Filtered {len(filtered)} issues from {len(issues_data)-1} total issues")
    return filtered

def main():
    try:
        credentials = authenticate()
        sheets = build('sheets', 'v4', credentials=credentials)
        
        # Get sheet titles for validation
        get_sheet_titles(sheets, UTILS_SHEET_ID)
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return
        
        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")
                
                # Check if required sheets exist
                titles = get_sheet_titles(sheets, sheet_id)
                
                if G_MILESTONES not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MILESTONES}' sheet")
                    continue
                
                if G_ISSUES_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_ISSUES_SHEET}' sheet")
                    continue
                
                # Get selected milestones for filtering
                print(f"📋 Getting milestones from {sheet_id} - {G_MILESTONES}")
                milestones = get_selected_milestones(sheets, sheet_id, G_MILESTONES)
                print(f"📋 Found {len(milestones)} milestones")
                
                if not milestones:
                    print(f"⚠️ No milestones found for {sheet_id}, skipping...")
                    continue
                
                # Get all issues from the new GitLab source
                print(f"📋 Getting issues from {CBS_ID} - {GITLAB_ISSUES}")
                issues_data = get_all_issues(sheets)
                print(f"📋 Found {len(issues_data)} total rows (including header)")
                
                # Filter by milestones with improved logic
                filtered = filter_issues_by_milestones(issues_data, milestones)
                
                if not filtered:
                    print(f"⚠️ No matching issues found for {sheet_id}")
                    # Still clear the sheet and update timestamp
                    clear_g_issues(sheets, sheet_id)
                    update_timestamp(sheets, sheet_id)
                    continue
                
                # Map source columns to target columns
                processed = []
                for row in filtered:
                    # Ensure row has enough columns (pad to at least 22 columns)
                    padded_row = row + [''] * (22 - len(row))
                    
                    # Map columns from source to target format
                    mapped_row = [
                        padded_row[1],   # C: ID (from B)
                        '',              # D: IID (skip - not in source)
                        padded_row[2],   # E: Issue Title (from C)
                        padded_row[3],   # F: Issue Author (from D)
                        '',              # G: Assignee (skip - not in source)
                        padded_row[4],   # H: Labels (from E)
                        padded_row[5],   # I: Milestone (from F)
                        padded_row[6],   # J: Status (from G)
                        padded_row[7],   # K: Created At (from H)
                        padded_row[8],   # L: Closed At (from I)
                        padded_row[10] if len(padded_row) > 10 else '',  # M: Closed By (from K)
                        padded_row[9],   # N: Project (from J)
                        padded_row[17] if len(padded_row) > 17 else '',  # O: Reviewer (from R)
                        padded_row[19] if len(padded_row) > 19 else '',  # P: Reopened? (from T)
                        '',              # Q: Reopened By (skip - not in source)
                        '',              # R: Date Reopened (skip - not in source)
                        '',              # S: Local Status (skip - not in source)
                        '',              # T: Status Date (skip - not in source)
                        ''               # U: Duration/Status (skip - not in source)
                    ]
                    processed.append(mapped_row)
                
                print(f"📊 Processing {len(processed)} filtered issues")
                
                # Clear existing data and insert new data
                clear_g_issues(sheets, sheet_id)
                insert_data_to_g_issues(sheets, sheet_id, processed)
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

if __name__ == '__main__':
    main()
