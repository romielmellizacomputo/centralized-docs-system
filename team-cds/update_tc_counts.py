import sys
import os
import re
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

# Rate limiting - Stay well under 60 requests/minute
BATCH_SIZE = 25  # Process 25 spreadsheets per batch
REQUEST_DELAY = 1.2  # 1.2 seconds between requests
BATCH_COOLDOWN = 65  # Wait 65 seconds between batches to reset quota

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
    OPTIMIZED: Count test cases with a single batch request
    Returns: (total_sheets_processed, test_case_counts_dict)
    """
    try:
        # Get spreadsheet metadata (1 API call)
        spreadsheet = sheets.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False
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
        
        # Build batch ranges for all C5 cells
        ranges = [f"'{sheet_info['properties']['title']}'!C5" for sheet_info in sheets_to_process]
        
        # Single batchGet request for all C5 values (1 API call)
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
            
            if c5_value:
                test_case_counts[c5_value] = test_case_counts.get(c5_value, 0) + 1
        
        return total_sheets_processed, test_case_counts
        
    except Exception as e:
        print(f"❌ Error accessing spreadsheet {spreadsheet_id}: {e}")
        return 0, {}

def process_batch(sheets, sheet_id, batch_urls, batch_num, total_batches):
    """Process a batch of URLs with rate limiting"""
    print(f"\n{'='*60}")
    print(f"📦 Processing Batch {batch_num}/{total_batches}")
    print(f"{'='*60}")
    
    updates = []
    processed = 0
    skipped = 0
    api_calls = 0
    
    for idx, url_info in enumerate(batch_urls, start=1):
        row_num = url_info['row']
        url = url_info['url']
        
        print(f"\n[{idx}/{len(batch_urls)}] Row {row_num}: Processing...")
        
        # Skip invalid URLs
        if not is_valid_google_sheets_url(url):
            print(f"Row {row_num}: Invalid URL - skipping")
            skipped += 1
            continue
        
        # Extract spreadsheet ID
        spreadsheet_id = extract_spreadsheet_id(url)
        if not spreadsheet_id:
            print(f"Row {row_num}: Could not extract spreadsheet ID")
            skipped += 1
            continue
        
        print(f"Row {row_num}: Spreadsheet ID = {spreadsheet_id}")
        
        # Add delay before processing (except for first item)
        if processed > 0:
            print(f"⏳ Waiting {REQUEST_DELAY}s...")
            time.sleep(REQUEST_DELAY)
        
        # Count test cases
        try:
            total_sheets, test_case_counts = count_test_cases_in_sheet_optimized(sheets, spreadsheet_id)
            api_calls += 2  # 2 API calls per spreadsheet
            
            print(f"Row {row_num}: Found {total_sheets} test case sheets")
            if test_case_counts:
                counts_str = ', '.join([f"{count} {cat}" for cat, count in test_case_counts.items()])
                print(f"Row {row_num}: {counts_str}")
            
            # Queue update
            updates.append({
                'range': f"'{TC_REVIEW_SHEET}'!K{row_num}",
                'values': [[total_sheets]]
            })
            
            print(f"Row {row_num}: ✅ Count = {total_sheets}")
            print(f"📊 Progress: {api_calls} API calls in this batch")
            processed += 1
            
        except Exception as e:
            error_msg = str(e)
            if '429' in error_msg or 'Quota exceeded' in error_msg:
                print(f"Row {row_num}: ⚠️ Rate limit hit - waiting and retrying...")
                print(f"⏸️ Sleeping 70 seconds to reset quota...")
                time.sleep(70)
                
                # Retry the same URL
                try:
                    total_sheets, test_case_counts = count_test_cases_in_sheet_optimized(sheets, spreadsheet_id)
                    api_calls += 2
                    
                    updates.append({
                        'range': f"'{TC_REVIEW_SHEET}'!K{row_num}",
                        'values': [[total_sheets]]
                    })
                    print(f"Row {row_num}: ✅ Retry successful - Count = {total_sheets}")
                    processed += 1
                except Exception as retry_error:
                    print(f"Row {row_num}: ❌ Retry failed - {retry_error}")
                    skipped += 1
            else:
                print(f"Row {row_num}: ❌ Error - {e}")
                skipped += 1
    
    # Apply all updates for this batch
    if updates:
        print(f"\n📤 Applying {len(updates)} updates from batch {batch_num}...")
        try:
            batch_update_body = {
                'valueInputOption': 'RAW',
                'data': updates
            }
            sheets.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=batch_update_body
            ).execute()
            print(f"✅ Batch {batch_num} updates applied successfully")
        except Exception as e:
            print(f"❌ Error applying batch {batch_num} updates: {e}")
    
    print(f"\n📊 Batch {batch_num} Summary:")
    print(f"   ✅ Processed: {processed}")
    print(f"   ⏭️  Skipped: {skipped}")
    print(f"   📈 API calls: {api_calls}")
    
    return processed, skipped

def update_tc_review_counts(sheets, sheet_id):
    """Update test case counts in TC Review sheet - Process ALL rows"""
    print(f"\n🔄 Processing TC Review counts for {sheet_id}")
    
    # Get all URLs
    url_data = get_tc_review_urls(sheets, sheet_id)
    
    if not url_data:
        print("⚠️ No URLs to process")
        return
    
    # Split into batches
    batches = []
    for i in range(0, len(url_data), BATCH_SIZE):
        batches.append(url_data[i:i + BATCH_SIZE])
    
    total_batches = len(batches)
    print(f"\n📦 Total URLs: {len(url_data)}")
    print(f"📦 Batches: {total_batches} (max {BATCH_SIZE} URLs per batch)")
    print(f"⏱️  Estimated time: ~{(total_batches * BATCH_SIZE * REQUEST_DELAY + (total_batches - 1) * BATCH_COOLDOWN) / 60:.1f} minutes")
    
    total_processed = 0
    total_skipped = 0
    
    # Process each batch
    for batch_num, batch_urls in enumerate(batches, start=1):
        processed, skipped = process_batch(sheets, sheet_id, batch_urls, batch_num, total_batches)
        total_processed += processed
        total_skipped += skipped
        
        # Wait between batches (except after the last batch)
        if batch_num < total_batches:
            print(f"\n⏸️  Batch {batch_num} complete. Cooling down for {BATCH_COOLDOWN}s to reset API quota...")
            print(f"📊 Overall progress: {batch_num}/{total_batches} batches ({total_processed + total_skipped}/{len(url_data)} URLs)")
            time.sleep(BATCH_COOLDOWN)
    
    print(f"\n{'='*60}")
    print(f"🎉 ALL BATCHES COMPLETE")
    print(f"{'='*60}")
    print(f"✅ Total processed: {total_processed}/{len(url_data)}")
    print(f"⏭️  Total skipped: {total_skipped}/{len(url_data)}")

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
                print(f"\n{'#'*60}")
                print(f"# Processing Sheet: {sheet_id}")
                print(f"{'#'*60}")
                
                # Verify TC Review sheet exists
                titles = get_sheet_titles(sheets, sheet_id)
                
                if TC_REVIEW_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{TC_REVIEW_SHEET}' sheet")
                    continue
                
                # Update TC Review test case counts (processes ALL rows)
                update_tc_review_counts(sheets, sheet_id)
                
                # Update timestamp in Dashboard
                update_timestamp(sheets, sheet_id)
                
                print(f"\n✅ Finished: {sheet_id}")
                
            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")
                import traceback
                traceback.print_exc()
        
        print("\n" + "="*60)
        print("✅ Script completed successfully - All sheets processed")
        print("="*60)
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
