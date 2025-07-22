import sys
import os
from datetime import datetime
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

def preserve_hyperlink_format(value):
    """Preserve hyperlink format if the value contains a hyperlink formula"""
    if not value or not isinstance(value, str):
        return value
    
    # Check if it's already a hyperlink formula
    if value.strip().startswith('=HYPERLINK('):
        return value
    
    # If it looks like a URL but isn't formatted as a hyperlink, leave it as is
    # Google Sheets will auto-detect URLs when using USER_ENTERED
    return value

def pad_row_to_u(row):
    """Pad row to 19 columns (C to U = 19 columns)"""
    full_length = 19
    return row + [''] * (full_length - len(row))
    """Pad row to 19 columns (C to U = 19 columns)"""
    full_length = 19
    return row + [''] * (full_length - len(row))

def parse_date_safely(date_str):
    """Safely parse date string and return datetime object or None"""
    if not date_str or not isinstance(date_str, str):
        return None
    
    date_str = date_str.strip()
    if not date_str:
        return None
    
    # Try common date formats
    date_formats = [
        '%Y-%m-%d %H:%M:%S',  # 2024-01-15 10:30:00
        '%Y-%m-%d',           # 2024-01-15
        '%m/%d/%Y %H:%M:%S',  # 01/15/2024 10:30:00
        '%m/%d/%Y',           # 01/15/2024
        '%d/%m/%Y %H:%M:%S',  # 15/01/2024 10:30:00
        '%d/%m/%Y',           # 15/01/2024
        '%Y-%m-%dT%H:%M:%S',  # ISO format without timezone
        '%Y-%m-%dT%H:%M:%SZ', # ISO format with Z
    ]
    
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    print(f"⚠️ Could not parse date: '{date_str}'")
    return None

def sort_issues_by_date(filtered_issues):
    """Sort issues by creation date (column H, index 6) - most recent first"""
    print(f"📅 Sorting {len(filtered_issues)} issues by creation date...")
    
    def get_sort_key(row):
        # Column H is index 6 (B=0, C=1, D=2, E=3, F=4, G=5, H=6)
        created_date_str = row[6] if len(row) > 6 else ''
        parsed_date = parse_date_safely(created_date_str)
        
        # Return parsed date if available, otherwise use epoch (very old date)
        # This ensures unparseable dates go to the bottom
        if parsed_date:
            return parsed_date
        else:
            return datetime(1970, 1, 1)  # Epoch time for unparseable dates
    
    try:
        # Sort in descending order (most recent first)
        sorted_issues = sorted(filtered_issues, key=get_sort_key, reverse=True)
        
        # Debug: show first few dates
        print("📅 First 3 sorted dates:")
        for i, row in enumerate(sorted_issues[:3]):
            date_str = row[6] if len(row) > 6 else 'N/A'
            parsed = parse_date_safely(date_str)
            print(f"  {i+1}. {date_str} -> {parsed}")
        
        print(f"✅ Successfully sorted {len(sorted_issues)} issues by date")
        return sorted_issues
        
    except Exception as e:
        print(f"❌ Error sorting issues by date: {str(e)}")
        print("📄 Returning unsorted list...")
        return filtered_issues

def insert_data_to_g_issues(sheets, sheet_id, data):
    """Insert processed data to G-Issues sheet"""
    padded_data = [pad_row_to_u(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_ISSUES_SHEET}!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{G_ISSUES_SHEET}!C4',
        valueInputOption='USER_ENTERED',  # Changed from 'RAW' to preserve hyperlinks
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    """Update timestamp in Dashboard sheet"""
    timestamp = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{DASHBOARD_SHEET}!W6',
        valueInputOption='USER_ENTERED',  # Changed from 'RAW' to be consistent
        body={'values': [[timestamp]]}
    ).execute()

def debug_milestone_matching(issues_data, milestones):
    """Debug function to help understand milestone matching issues"""
    print("🔍 DEBUG: Analyzing milestone data in column F (index 4)...")
    
    if not issues_data:
        print("🔍 No source data found!")
        return
    
    # Show header row
    header = issues_data[0] if issues_data else []
    print(f"🔍 Header row: B={header[0] if len(header) > 0 else 'N/A'}, C={header[1] if len(header) > 1 else 'N/A'}, D={header[2] if len(header) > 2 else 'N/A'}, E={header[3] if len(header) > 3 else 'N/A'}, F={header[4] if len(header) > 4 else 'N/A'}, G={header[5] if len(header) > 5 else 'N/A'}")
    
    # Get unique milestones from source data (column F, index 4)
    source_milestones = set()
    for row in issues_data[1:]:  # Skip header row
        if len(row) > 4 and row[4] and str(row[4]).strip():  # Column F is index 4
            source_milestones.add(str(row[4]).strip())
    
    print(f"🔍 Found {len(source_milestones)} unique milestones in column F")
    print(f"🔍 First 10 source milestones: {list(source_milestones)[:10]}")
    
    # Check for exact matches
    exact_matches = source_milestones.intersection(set(milestones))
    print(f"🔍 Exact matches: {len(exact_matches)} - {list(exact_matches)[:5]}")
    
    # Show target milestones for comparison
    print(f"🔍 Target milestones (first 5): {milestones[:5]}")

def filter_issues_by_milestones(issues_data, milestones):
    """Filter issues by milestones using column F (index 4)"""
    if not issues_data:
        return []
    
    # Debug milestone matching
    debug_milestone_matching(issues_data, milestones)
    
    filtered = []
    milestone_set = set(milestones)
    milestone_col_idx = 4  # Column F is index 4 (B=0, C=1, D=2, E=3, F=4)
    
    # Skip header row and process data rows
    for i, row in enumerate(issues_data[1:], 1):
        if len(row) <= milestone_col_idx:  # Row doesn't have enough columns
            continue
            
        milestone = str(row[milestone_col_idx]).strip() if row[milestone_col_idx] else ""  # Column F (index 4)
        
        if milestone in milestone_set:
            filtered.append(row)
            if len(filtered) <= 5:  # Show first 5 matches for debugging
                print(f"✅ Match found at row {i}: milestone='{milestone}'")
    
    print(f"📊 Filtered {len(filtered)} issues from {len(issues_data)-1} total issues using column F")
    return filtered, milestone_col_idx

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
                
                # Filter by milestones using correct column F (index 4)
                filtered_result = filter_issues_by_milestones(issues_data, milestones)
                
                if isinstance(filtered_result, tuple):
                    filtered, milestone_col_idx = filtered_result
                else:
                    print("❌ Error in filtering function, skipping...")
                    continue
                
                if not filtered:
                    print(f"⚠️ No matching issues found for {sheet_id}")
                    # Still clear the sheet and update timestamp
                    clear_g_issues(sheets, sheet_id)
                    update_timestamp(sheets, sheet_id)
                    continue
                
                # Sort by creation date (most recent first) before processing
                sorted_filtered = sort_issues_by_date(filtered)
                
                # Map source columns to target columns
                processed = []
                for row in sorted_filtered:
                    # Ensure row has enough columns (pad to at least 22 columns)
                    padded_row = row + [''] * (22 - len(row))
                    
                    # Map columns from source to target format with hyperlink preservation
                    # Column mapping: B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7, J=8, K=9, L=10, etc.
                    mapped_row = [
                        preserve_hyperlink_format(padded_row[0] if len(padded_row) > 0 else ''),   # C: ID (from B, index 0)
                        '',                                                                          # D: IID (skip - not in source)
                        preserve_hyperlink_format(padded_row[1] if len(padded_row) > 1 else ''),   # E: Issue Title (from C, index 1)
                        preserve_hyperlink_format(padded_row[2] if len(padded_row) > 2 else ''),   # F: Issue Author (from D, index 2)
                        '',                                                                          # G: Assignee (skip - not in source)
                        preserve_hyperlink_format(padded_row[3] if len(padded_row) > 3 else ''),   # H: Labels (from E, index 3)
                        preserve_hyperlink_format(padded_row[4] if len(padded_row) > 4 else ''),   # I: Milestone (from F, index 4)
                        preserve_hyperlink_format(padded_row[5] if len(padded_row) > 5 else ''),   # J: Status (from G, index 5)
                        preserve_hyperlink_format(padded_row[6] if len(padded_row) > 6 else ''),   # K: Created At (from H, index 6)
                        preserve_hyperlink_format(padded_row[7] if len(padded_row) > 7 else ''),   # L: Closed At (from I, index 7)
                        preserve_hyperlink_format(padded_row[9] if len(padded_row) > 9 else ''),   # M: Closed By (from K, index 9)
                        preserve_hyperlink_format(padded_row[8] if len(padded_row) > 8 else ''),   # N: Project (from J, index 8)
                        preserve_hyperlink_format(padded_row[16] if len(padded_row) > 16 else ''), # O: Reviewer (from R, index 16)
                        preserve_hyperlink_format(padded_row[18] if len(padded_row) > 18 else ''), # P: Reopened? (from T, index 18)
                        '',                                                                          # Q: Reopened By (skip - not in source)
                        '',                                                                          # R: Date Reopened (skip - not in source)
                        '',                                                                          # S: Local Status (skip - not in source)
                        '',                                                                          # T: Status Date (skip - not in source)
                        ''                                                                           # U: Duration/Status (skip - not in source)
                    ]
                    processed.append(mapped_row)
                
                print(f"📊 Processing {len(processed)} filtered and sorted issues")
                
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
