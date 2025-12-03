import sys
import os
from datetime import datetime

# Add the parent directory to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from googleapiclient.discovery import build
from common import authenticate
from constants import (
    SHEET_SYNC_SID,
    CBS_ID
)

def get_all_issues(sheets):
    """Get all issues from SHEET_SYNC_SID - ALL ISSUES sheet"""
    print(f"📋 Getting issues from {SHEET_SYNC_SID} - ALL ISSUES!C4:T")
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL ISSUES!C4:T"
    ).execute()
    
    values = result.get('values', [])
    if not values:
        print("⚠️ No data found in range ALL ISSUES!C4:T")
        return []
    
    print(f"📋 Found {len(values)} total rows")
    return values

def clear_cbs_issues(sheets):
    """Clear existing data in CBS_ID - ALL ISSUES sheet"""
    print(f"🧹 Clearing CBS_ID - ALL ISSUES!C4:T")
    sheets.spreadsheets().values().clear(
        spreadsheetId=CBS_ID,
        range='ALL ISSUES!C4:T'
    ).execute()

def pad_row_to_t(row):
    """Pad row to 18 columns (C to T = 18 columns)"""
    full_length = 18
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
    
    return None

def sort_issues_by_date(issues):
    """Sort issues by creation date (column K/index 8) - most recent first"""
    print(f"📅 Sorting {len(issues)} issues by creation date...")
    
    def get_sort_key(row):
        # Column K is index 8 (C=0, D=1, E=2, F=3, G=4, H=5, I=6, J=7, K=8)
        created_date_str = row[8] if len(row) > 8 else ''
        parsed_date = parse_date_safely(created_date_str)
        
        # Return parsed date if available, otherwise use epoch (very old date)
        if parsed_date:
            return parsed_date
        else:
            return datetime(1970, 1, 1)  # Epoch time for unparseable dates
    
    try:
        # Sort in descending order (most recent first)
        sorted_issues = sorted(issues, key=get_sort_key, reverse=True)
        
        # Debug: show first few dates
        print("📅 First 3 sorted dates:")
        for i, row in enumerate(sorted_issues[:3]):
            date_str = row[8] if len(row) > 8 else 'N/A'
            parsed = parse_date_safely(date_str)
            print(f"  {i+1}. {date_str} -> {parsed}")
        
        print(f"✅ Successfully sorted {len(sorted_issues)} issues by date")
        return sorted_issues
        
    except Exception as e:
        print(f"❌ Error sorting issues by date: {str(e)}")
        print("📄 Returning unsorted list...")
        return issues

def insert_data_to_cbs(sheets, data):
    """Insert data to CBS_ID - ALL ISSUES sheet"""
    if not data:
        print("⚠️ No data to insert")
        return
    
    padded_data = [pad_row_to_t(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to CBS_ID - ALL ISSUES!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=CBS_ID,
        range='ALL ISSUES!C4',
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()
    
    print(f"✅ Successfully inserted {len(padded_data)} rows")

def main():
    try:
        print("=" * 60)
        print("🚀 Starting CBS ALL ISSUES Sync")
        print("=" * 60)
        
        credentials = authenticate()
        sheets = build('sheets', 'v4', credentials=credentials)
        
        # Get all issues from source (SHEET_SYNC_SID)
        issues_data = get_all_issues(sheets)
        
        if not issues_data:
            print("⚠️ No issues found, clearing CBS sheet")
            clear_cbs_issues(sheets)
            print("✅ Process completed (no data)")
            return
        
        # Sort by creation date (most recent first)
        sorted_issues = sort_issues_by_date(issues_data)
        
        print(f"📊 Processing {len(sorted_issues)} issues")
        
        # Clear existing data and insert new data
        clear_cbs_issues(sheets)
        insert_data_to_cbs(sheets, sorted_issues)
        
        print("=" * 60)
        print("✅ CBS ALL ISSUES Sync Completed Successfully")
        print(f"📊 Total issues synced: {len(sorted_issues)}")
        print("=" * 60)
        
    except Exception as e:
        print("=" * 60)
        print(f"❌ Process failed: {str(e)}")
        print("=" * 60)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
