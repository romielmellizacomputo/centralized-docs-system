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
    CBS_SID
)

def get_all_mr(sheets):
    """Get all MRs from SHEET_SYNC_SID - ALL MRs sheet"""
    print(f"📋 Getting MRs from {SHEET_SYNC_SID} - ALL MRs!C4:S")
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL MRs!C4:S"
    ).execute()
    
    values = result.get('values', [])
    if not values:
        print("⚠️ No data found in range ALL MRs!C4:S")
        return []
    
    print(f"📋 Found {len(values)} total rows")
    return values

def clear_cbs_mrs(sheets):
    """Clear existing data in CBS_SID - ALL MRs sheet"""
    print(f"🧹 Clearing CBS_SID - ALL MRs!C4:S")
    sheets.spreadsheets().values().clear(
        spreadsheetId=CBS_SID,
        range='ALL MRs!C4:S'
    ).execute()

def pad_row_to_s(row):
    """Pad row to 17 columns (C to S = 17 columns)"""
    full_length = 17
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

def sort_mrs_by_date(mrs):
    """Sort MRs by creation date (column K/index 8) - most recent first"""
    print(f"📅 Sorting {len(mrs)} MRs by creation date...")
    
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
        sorted_mrs = sorted(mrs, key=get_sort_key, reverse=True)
        
        # Debug: show first few dates
        print("📅 First 3 sorted dates:")
        for i, row in enumerate(sorted_mrs[:3]):
            date_str = row[8] if len(row) > 8 else 'N/A'
            parsed = parse_date_safely(date_str)
            print(f"  {i+1}. {date_str} -> {parsed}")
        
        print(f"✅ Successfully sorted {len(sorted_mrs)} MRs by date")
        return sorted_mrs
        
    except Exception as e:
        print(f"❌ Error sorting MRs by date: {str(e)}")
        print("📄 Returning unsorted list...")
        return mrs

def insert_data_to_cbs(sheets, data):
    """Insert data to CBS_SID - ALL MRs sheet"""
    if not data:
        print("⚠️ No data to insert")
        return
    
    padded_data = [pad_row_to_s(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to CBS_SID - ALL MRs!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=CBS_SID,
        range='ALL MRs!C4',
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()
    
    print(f"✅ Successfully inserted {len(padded_data)} rows")

def main():
    try:
        print("=" * 60)
        print("🚀 Starting CBS ALL MRs Sync")
        print("=" * 60)
        
        credentials = authenticate()
        sheets = build('sheets', 'v4', credentials=credentials)
        
        # Get all MRs from source (SHEET_SYNC_SID)
        mr_data = get_all_mr(sheets)
        
        if not mr_data:
            print("⚠️ No MRs found, clearing CBS sheet")
            clear_cbs_mrs(sheets)
            print("✅ Process completed (no data)")
            return
        
        # Sort by creation date (most recent first)
        sorted_mrs = sort_mrs_by_date(mr_data)
        
        print(f"📊 Processing {len(sorted_mrs)} MRs")
        
        # Clear existing data and insert new data
        clear_cbs_mrs(sheets)
        insert_data_to_cbs(sheets, sorted_mrs)
        
        print("=" * 60)
        print("✅ CBS ALL MRs Sync Completed Successfully")
        print(f"📊 Total MRs synced: {len(sorted_mrs)}")
        print("=" * 60)
        
    except Exception as e:
        print("=" * 60)
        print(f"❌ Process failed: {str(e)}")
        print("=" * 60)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
