import sys
import os
from datetime import datetime
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from googleapiclient.discovery import build
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids
)
from constants import (
    UTILS_SHEET_ID,
    G_TC_SHEET,
    DASHBOARD_SHEET,
    SHEET_SYNC_SID,
    generate_timestamp_string
)

def get_all_test_cases(sheets):
    """Get all test cases from ALL ISSUES sheet (SHEET_SYNC_SID)"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL ISSUES!C4:N"
    ).execute()
    
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range ALL ISSUES!C4:N")
    
    return values

def clear_g_tc(sheets, sheet_id):
    """Clear existing data in G-TC sheet"""
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f'{G_TC_SHEET}!C4:N'
    ).execute()

def pad_row_to_n(row):
    """Pad row to 12 columns (C to N = 12 columns)"""
    full_length = 12
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

def sort_test_cases_by_date(filtered_test_cases):
    """Sort test cases by creation date (column K/index 8) - most recent first"""
    print(f"📅 Sorting {len(filtered_test_cases)} test cases by creation date...")
    
    def get_sort_key(row):
        # Column K is index 8 (C=0, D=1, E=2, F=3, G=4, H=5, I=6, J=7, K=8)
        created_date_str = row[8] if len(row) > 8 else ''
        parsed_date = parse_date_safely(created_date_str)
        
        # Return parsed date if available, otherwise use epoch (very old date)
        # This ensures unparseable dates go to the bottom
        if parsed_date:
            return parsed_date
        else:
            return datetime(1970, 1, 1)  # Epoch time for unparseable dates
    
    try:
        # Sort in descending order (most recent first)
        sorted_test_cases = sorted(filtered_test_cases, key=get_sort_key, reverse=True)
        
        # Debug: show first few dates
        print("📅 First 3 sorted dates:")
        for i, row in enumerate(sorted_test_cases[:3]):
            date_str = row[8] if len(row) > 8 else 'N/A'
            parsed = parse_date_safely(date_str)
            print(f"  {i+1}. {date_str} -> {parsed}")
        
        print(f"✅ Successfully sorted {len(sorted_test_cases)} test cases by date")
        return sorted_test_cases
        
    except Exception as e:
        print(f"❌ Error sorting test cases by date: {str(e)}")
        print("📄 Returning unsorted list...")
        return filtered_test_cases

def insert_data_to_g_tc(sheets, sheet_id, data):
    """Insert processed data to G-TC sheet"""
    padded_data = [pad_row_to_n(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_TC_SHEET}!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{G_TC_SHEET}!C4',
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

def debug_qa_team_filtering(test_cases_data):
    """Debug function to show QA TEAM filtering details"""
    print("🔍 DEBUG: Analyzing milestone data in column I (index 6)...")
    
    if not test_cases_data:
        print("🔍 No source data found!")
        return
    
    # Get unique milestones from source data (column I, index 6)
    source_milestones = set()
    qa_team_count = 0
    
    for row in test_cases_data:
        if len(row) > 6 and row[6] and str(row[6]).strip():  # Column I is index 6
            milestone = str(row[6]).strip()
            source_milestones.add(milestone)
            if milestone == "QA TEAM":
                qa_team_count += 1
    
    print(f"🔍 Found {len(source_milestones)} unique milestones in column I")
    print(f"🔍 First 10 source milestones: {list(source_milestones)[:10]}")
    print(f"🔍 Rows with 'QA TEAM' milestone: {qa_team_count}")

def filter_test_cases_by_qa_team(test_cases_data):
    """Filter test cases by 'QA TEAM' milestone in column I (index 6)"""
    if not test_cases_data:
        return []
    
    # Debug QA TEAM filtering
    debug_qa_team_filtering(test_cases_data)
    
    filtered = []
    milestone_col_idx = 6  # Column I is index 6 (C=0, D=1, E=2, F=3, G=4, H=5, I=6)
    
    # Process data rows (no header in C4:M range)
    for i, row in enumerate(test_cases_data, 1):
        if len(row) <= milestone_col_idx:  # Row doesn't have enough columns
            continue
            
        milestone = str(row[milestone_col_idx]).strip() if row[milestone_col_idx] else ""
        
        if milestone == "QA TEAM":
            filtered.append(row)
            if len(filtered) <= 5:  # Show first 5 matches for debugging
                print(f"✅ Match found at row {i}: milestone='{milestone}'")
    
    print(f"📊 Filtered {len(filtered)} test cases from {len(test_cases_data)} total rows using 'QA TEAM' milestone")
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
                
                if G_TC_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_TC_SHEET}' sheet")
                    continue
                
                # Get all test cases from SHEET_SYNC_SID - ALL ISSUES sheet
                print(f"📋 Getting test cases from {SHEET_SYNC_SID} - ALL ISSUES!C4:N")
                test_cases_data = get_all_test_cases(sheets)
                print(f"📋 Found {len(test_cases_data)} total rows")
                
                # Filter by 'QA TEAM' milestone in column I (index 6)
                filtered = filter_test_cases_by_qa_team(test_cases_data)
                
                if not filtered:
                    print(f"⚠️ No test cases with 'QA TEAM' milestone found for {sheet_id}")
                    # Still clear the sheet and update timestamp
                    clear_g_tc(sheets, sheet_id)
                    update_timestamp(sheets, sheet_id)
                    continue
                
                # Sort by creation date (most recent first)
                sorted_filtered = sort_test_cases_by_date(filtered)
                
                print(f"📊 Processing {len(sorted_filtered)} filtered and sorted test cases")
                
                # Clear existing data and insert new data
                clear_g_tc(sheets, sheet_id)
                insert_data_to_g_tc(sheets, sheet_id, sorted_filtered)
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
