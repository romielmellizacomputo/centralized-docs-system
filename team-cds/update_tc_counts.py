import sys
import os
import re
import json
import time
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
    DASHBOARD_SHEET,
    generate_timestamp_string
)

# Configuration
TC_REVIEW_SHEET = 'TC Review'
URL_COLUMN = 'I'  # Column I contains the Google Sheets URLs
TOTAL_CASES_COLUMN = 'K'  # Column K will contain the total test case count

# Sheets to skip when counting test cases
SHEETS_TO_SKIP = ["HELP", "ToC", "Issues", "Roster"]

# File to store last processed state
STATE_FILE = 'tc_review_state.json'
ROWS_TO_PROCESS_PER_RUN = 20  # Process only 20 URLs per run

# Rate limiting - More aggressive to stay under 60 requests/minute
REQUEST_DELAY = 2.0  # 2 seconds between requests (max 30 spreadsheets/minute)
# With 2 API calls per spreadsheet, that's 60 requests/minute = at the limit

def load_last_processed_state(sheet_id):
    """Load the last processed row for a specific sheet"""
    if not os.path.exists(STATE_FILE):
        return 2  # Start from row 2 (skip header)
    
    try:
        with open(STATE_FILE, 'r') as f:
            state = json.load(f)
            return state.get(sheet_id, 2)
    except Exception as e:
        print(f"⚠️ Error loading state: {e}")
        return 2

def save_last_processed_state(sheet_id, last_row):
    """Save the last processed row for a specific sheet"""
    state = {}
    
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)
        except:
            pass
    
    state[sheet_id] = last_row
    
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)
    except Exception as e:
        print(f"⚠️ Error saving state: {e}")

def is_valid_google_sheets_url(url):
    """Validate if the URL is a valid Google Sheets URL"""
    if not url or not isinstance(url, str):
        return False
    
    url = url.strip()
    pattern = r'^https://docs\.google\.com/spreadsheets/d/[a-zA-Z0-9-_]+.*'
    return bool(re.match(pattern, url))

def extract_spreadsheet_id(url):
    """Extract spreadsheet ID from Google Sheets URL"""
    if not url:
        return None
    
    # Pattern: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/...
    pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url)
    
    if match:
        return match.group(1)
    return None

def get_tc_review_urls(sheets, sheet_id):
    """Get all URLs from TC Review sheet column I"""
    print(f"📋 Fetching URLs from TC Review sheet in {sheet_id}")
    
    # Get data from column I starting at row 2
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TC_REVIEW_SHEET}'!I2:I",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    
    values = result.get('values', [])
    if not values:
        print("⚠️ No URLs found in TC Review sheet")
        return []
    
    # Flatten the list and get actual URLs
    urls = []
    for idx, row in enumerate(values, start=2):
        url = row[0] if row and row[0] else ''
        urls.append({'row': idx, 'url': str(url).strip() if url else ''})
    
    print(f"✅ Found {len(urls)} rows with potential URLs")
    return urls

def count_test_cases_in_sheet_optimized(sheets, spreadsheet_id):
    """
    OPTIMIZED: Count test cases with a single batch request instead of multiple API calls
    Returns: (total_sheets_processed, test_case_counts_dict)
    """
    try:
        # OPTIMIZATION 1: Get spreadsheet metadata and all C5 values in ONE request
        spreadsheet = sheets.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False  # Don't fetch all cell data, just metadata
        ).execute()
        
        all_sheets = spreadsheet.get('sheets', [])
        
        # Filter out sheets to skip
        sheets_to_process = [
            sheet_info for sheet_info in all_sheets
            if sheet_info['properties']['title'] not in SHEETS_TO_SKIP
        ]
        
        if not sheets_to_process:
            return 0, {}
        
        total_sheets_processed = len(sheets_to_process)
        
        # OPTIMIZATION 2: Build batch ranges for all C5 cells at once
        ranges = [f"'{sheet_info['properties']['title']}'!C5" for sheet_info in sheets_to_process]
        
        # OPTIMIZATION 3: Single batchGet request for all C5 values
        result = sheets.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=ranges,
            valueRenderOption='UNFORMATTED_VALUE'
        ).execute()
        
        value_ranges = result.get('valueRanges', [])
        
        # Count test case categories
        test_case_counts = {}
        
        for value_range in value_ranges:
            values = value_range.get('values', [])
            c5_value = str(values[0][0]).strip() if values and values[0] and len(values[0]) > 0 and values[0][0] else ''
            
            # Count only if C5 has a value
            if c5_value:
                test_case_counts[c5_value] = test_case_counts.get(c5_value, 0) + 1
        
        return total_sheets_processed, test_case_counts
        
    except Exception as e:
        print(f"❌ Error accessing spreadsheet {spreadsheet_id}: {e}")
        return 0, {}

def create_note_text(test_case_counts):
    """Create note text based on test case counts"""
    if not test_case_counts or len(test_case_counts) == 0:
        return "Test Case does not match the current Test Case Template. Please check and make sure to use updated test case template."
    
    note_lines = ["Test Case Counts:"]
    for category, count in sorted(test_case_counts.items()):
        note_lines.append(f"* {count} {category}")
    
    return '\n'.join(note_lines)

def update_tc_review_counts(sheets, sheet_id):
    """Update test case counts in TC Review sheet"""
    print(f"\n🔄 Processing TC Review counts for {sheet_id}")
    
    # Load last processed row
    last_processed_row = load_last_processed_state(sheet_id)
    print(f"📍 Starting from row {last_processed_row}")
    
    # Get all URLs
    url_data = get_tc_review_urls(sheets, sheet_id)
    
    if not url_data:
        print("⚠️ No URLs to process")
        return
    
    # Calculate end row
    end_row = min(last_processed_row + ROWS_TO_PROCESS_PER_RUN, len(url_data) + 1)
    rows_to_process = url_data[last_processed_row - 2:end_row - 2]  # Adjust for 0-indexing
    
    print(f"📊 Processing rows {last_processed_row} to {end_row - 1} ({len(rows_to_process)} rows)")
    print(f"⏱️ Estimated time: ~{len(rows_to_process) * REQUEST_DELAY / 60:.1f} minutes")
    
    updates = []
    processed_count = 0
    skipped_count = 0
    api_calls_made = 0
    
    for idx, url_info in enumerate(rows_to_process, start=1):
        row_num = url_info['row']
        url = url_info['url']
        
        print(f"\n[{idx}/{len(rows_to_process)}] Row {row_num}: Processing...")
        
        # Skip empty or invalid URLs
        if not is_valid_google_sheets_url(url):
            print(f"Row {row_num}: Invalid or missing URL - skipping")
            skipped_count += 1
            continue
        
        # Extract spreadsheet ID
        spreadsheet_id = extract_spreadsheet_id(url)
        if not spreadsheet_id:
            print(f"Row {row_num}: Could not extract spreadsheet ID from URL")
            skipped_count += 1
            continue
        
        print(f"Row {row_num}: Spreadsheet ID = {spreadsheet_id}")
        
        # Count test cases with optimized method
        try:
            # RATE LIMITING: Add delay BEFORE making the request
            if processed_count > 0:
                print(f"⏳ Waiting {REQUEST_DELAY}s to respect rate limits...")
                time.sleep(REQUEST_DELAY)
            
            total_sheets, test_case_counts = count_test_cases_in_sheet_optimized(sheets, spreadsheet_id)
            api_calls_made += 2  # Each spreadsheet requires 2 API calls
            
            print(f"Row {row_num}: Found {total_sheets} test case sheets")
            if test_case_counts:
                print(f"Row {row_num}: Categories: {test_case_counts}")
            
            # Update K column with total count
            updates.append({
                'range': f"'{TC_REVIEW_SHEET}'!K{row_num}",
                'values': [[total_sheets]]
            })
            
            print(f"Row {row_num}: ✅ Will update with count = {total_sheets}")
            print(f"📊 API calls so far: {api_calls_made}/60")
            processed_count += 1
            
        except Exception as e:
            error_msg = str(e)
            if '429' in error_msg or 'Quota exceeded' in error_msg:
                print(f"Row {row_num}: ⚠️ Rate limit hit after {api_calls_made} API calls")
                # Save current progress before stopping
                save_last_processed_state(sheet_id, row_num)
                print(f"💾 Saved progress at row {row_num}")
                print(f"⏸️ Please wait 60 seconds and run the script again to continue")
                
                # Apply updates collected so far
                if updates:
                    print(f"\n📤 Applying {len(updates)} updates before stopping...")
                    try:
                        batch_update_body = {
                            'valueInputOption': 'RAW',
                            'data': updates
                        }
                        sheets.spreadsheets().values().batchUpdate(
                            spreadsheetId=sheet_id,
                            body=batch_update_body
                        ).execute()
                        print(f"✅ Successfully applied updates")
                    except Exception as update_error:
                        print(f"❌ Error applying updates: {update_error}")
                
                return  # Exit function early
            else:
                print(f"Row {row_num}: ❌ Error processing - {e}")
                skipped_count += 1
                continue
    
    # Batch update all changes
    if updates:
        print(f"\n📤 Applying {len(updates)} updates...")
        try:
            batch_update_body = {
                'valueInputOption': 'RAW',
                'data': updates
            }
            
            sheets.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=batch_update_body
            ).execute()
            
            print(f"✅ Successfully applied all updates")
        except Exception as e:
            print(f"❌ Error applying updates: {e}")
    
    # Update last processed row
    if processed_count > 0:
        next_row = end_row
        if next_row > len(url_data) + 1:
            next_row = 2  # Reset to beginning
            print("🔄 All rows processed. Resetting to start.")
        
        save_last_processed_state(sheet_id, next_row)
        print(f"📍 Next run will start from row {next_row}")
    
    print(f"\n📊 Summary: Processed {processed_count} URLs, Skipped {skipped_count} URLs")
    print(f"📈 Total API calls made: {api_calls_made}")

def update_timestamp(sheets, sheet_id):
    """Update timestamp in Dashboard sheet"""
    timestamp = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{DASHBOARD_SHEET}!W6',
        valueInputOption='RAW',
        body={'values': [[timestamp]]}
    ).execute()
    print(f"🕐 Updated timestamp in Dashboard: {timestamp}")

def main():
    try:
        credentials = authenticate()
        sheets = build('sheets', 'v4', credentials=credentials)
        
        # Get all Team CDS sheet IDs from UTILS
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return
        
        # Process each sheet
        for sheet_id in sheet_ids:
            try:
                print(f"\n{'='*60}")
                print(f"🔄 Processing: {sheet_id}")
                print(f"{'='*60}")
                
                # Verify TC Review sheet exists
                titles = get_sheet_titles(sheets, sheet_id)
                
                if TC_REVIEW_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{TC_REVIEW_SHEET}' sheet")
                    continue
                
                # Update TC Review test case counts
                update_tc_review_counts(sheets, sheet_id)
                
                # Update timestamp in Dashboard
                update_timestamp(sheets, sheet_id)
                
                print(f"\n✅ Finished: {sheet_id}")
                
            except Exception as e:
                error_msg = str(e)
                if '429' in error_msg or 'Quota exceeded' in error_msg:
                    print(f"⚠️ Rate limit reached for {sheet_id}")
                    print(f"💾 Progress has been saved")
                    print(f"⏸️ Please wait a few minutes and run the script again")
                    break  # Stop processing other sheets
                else:
                    print(f"❌ Error processing {sheet_id}: {error_msg}")
                    import traceback
                    traceback.print_exc()
        
        print("\n" + "="*60)
        print("✅ Script completed")
        print("💡 If rate limited, run again in a few minutes to continue")
        print("="*60)
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
