import sys
import os
from datetime import datetime
from zoneinfo import ZoneInfo

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

def get_all_issues_with_formulas(sheets):
    """Get all issues from SHEET_SYNC_SID - ALL ISSUES sheet with formulas preserved"""
    print(f"📋 Getting issues with formulas from {SHEET_SYNC_SID} - ALL ISSUES!C4:T")
    
    # Get the spreadsheet data including formulas
    result = sheets.spreadsheets().get(
        spreadsheetId=SHEET_SYNC_SID,
        ranges=['ALL ISSUES!C4:T'],
        includeGridData=True
    ).execute()
    
    if 'sheets' not in result or not result['sheets']:
        print("⚠️ No sheets found")
        return []
    
    sheet_data = result['sheets'][0]
    if 'data' not in sheet_data or not sheet_data['data']:
        print("⚠️ No data found")
        return []
    
    grid_data = sheet_data['data'][0]
    if 'rowData' not in grid_data:
        print("⚠️ No row data found")
        return []
    
    rows = []
    for row_data in grid_data['rowData']:
        if 'values' not in row_data:
            continue
        
        row = []
        for i, cell in enumerate(row_data['values']):
            # Column E is index 2 (C=0, D=1, E=2)
            if i == 2:  # Column E - check for hyperlink
                if 'hyperlink' in cell:
                    # Create HYPERLINK formula
                    url = cell['hyperlink']
                    display_text = cell.get('formattedValue', url)
                    # Escape quotes in display text
                    display_text = display_text.replace('"', '""')
                    row.append(f'=HYPERLINK("{url}","{display_text}")')
                elif 'userEnteredValue' in cell:
                    # Check if it's already a formula
                    user_value = cell['userEnteredValue']
                    if 'formulaValue' in user_value:
                        row.append(user_value['formulaValue'])
                    else:
                        row.append(cell.get('formattedValue', ''))
                else:
                    row.append(cell.get('formattedValue', ''))
            else:
                # For other columns, just get the formatted value
                row.append(cell.get('formattedValue', ''))
        
        rows.append(row)
    
    print(f"📋 Found {len(rows)} total rows")
    return rows

def clear_cbs_issues(sheets):
    """Clear existing data in CBS_ID - ALL ISSUES sheet"""
    print(f"🧹 Clearing CBS_ID - ALL ISSUES!C4:T")
    sheets.spreadsheets().values().clear(
        spreadsheetId=CBS_ID,
        range='ALL ISSUES!C4:T'
    ).execute()

def update_timestamp(sheets):
    """Update timestamp in F2 with Kampala and Philippines time"""
    print(f"⏰ Updating timestamp in CBS_ID - ALL ISSUES!F2")
    
    # Get current time in both timezones
    kampala_tz = ZoneInfo('Africa/Kampala')  # Kampala/Uganda timezone (EAT - UTC+3)
    ph_tz = ZoneInfo('Asia/Manila')          # Philippines timezone (PHT - UTC+8)
    
    now_kampala = datetime.now(kampala_tz)
    now_ph = datetime.now(ph_tz)
    
    # Check if dates are the same
    kampala_date = now_kampala.strftime('%B %d, %Y')
    ph_date = now_ph.strftime('%B %d, %Y')
    
    if kampala_date == ph_date:
        # Same date: show date once
        # Format: "Sync on December 03, 2025, 06:29:04 AM (Kampala) / 11:29:04 AM (PH)"
        timestamp = (
            f"Sync on {kampala_date}, "
            f"{now_kampala.strftime('%I:%M:%S %p')} (EAT) / "
            f"{now_ph.strftime('%I:%M:%S %p')} (PHT)"
        )
    else:
        # Different dates: show both dates
        # Format: "Sync on December 02, 2025, 11:30:00 PM (Kampala) / December 03, 2025, 04:30:00 AM (PH)"
        timestamp = (
            f"Sync on {now_kampala.strftime('%B %d, %Y, %I:%M:%S %p')} (EAT) / "
            f"{now_ph.strftime('%B %d, %Y, %I:%M:%S %p')} (PHT)"
        )
    
    print(f"📝 Timestamp: {timestamp}")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=CBS_ID,
        range='ALL ISSUES!F2',
        valueInputOption='RAW',
        body={'values': [[timestamp]]}
    ).execute()
    
    print(f"✅ Timestamp updated in F2")

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
    """Clear existing data and insert new data to CBS_ID - ALL ISSUES sheet"""
    if not data:
        print("⚠️ No data to insert")
        return
    
    # Clear the target sheet first (right before inserting)
    print(f"🧹 Clearing CBS_ID - ALL ISSUES!C4:T before inserting new data")
    clear_cbs_issues(sheets)
    
    # Prepare and insert data
    padded_data = [pad_row_to_t(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to CBS_ID - ALL ISSUES!C4")
    
    # Use USER_ENTERED to preserve formulas (including HYPERLINK formulas)
    sheets.spreadsheets().values().update(
        spreadsheetId=CBS_ID,
        range='ALL ISSUES!C4',
        valueInputOption='USER_ENTERED',  # Changed from RAW to USER_ENTERED
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
        
        # Get all issues from source (SHEET_SYNC_SID) with formulas preserved
        issues_data = get_all_issues_with_formulas(sheets)
        
        if not issues_data:
            print("⚠️ No issues found in source")
            # Clear CBS sheet only if there's no data
            print(f"🧹 Clearing CBS_ID - ALL ISSUES!C4:T")
            clear_cbs_issues(sheets)
            update_timestamp(sheets)  # Update timestamp even when no data
            print("✅ Process completed (no data)")
            return
        
        # Sort by creation date (most recent first)
        sorted_issues = sort_issues_by_date(issues_data)
        
        print(f"📊 Processing {len(sorted_issues)} issues")
        
        # Insert data (this will clear the target sheet first, then insert)
        insert_data_to_cbs(sheets, sorted_issues)
        
        # Update timestamp after successful sync
        update_timestamp(sheets)
        
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
