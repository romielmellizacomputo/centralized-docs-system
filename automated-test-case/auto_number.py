from config import sheet_data, credentials_info
from constants import SCOPES, skip_sheets
from google_auth import get_sheet_service
from sheet_utils import (
    get_spreadsheet_id,
    process_sheet,
    get_sheet_metadata,
)
from retry_utils import execute_with_retries
import sys
import time
from datetime import datetime

# Rate limiting configuration
COOLDOWN_SECONDS = 65  # Cooldown between sheets (65s to stay under 60 req/min limit)
SHOW_COUNTDOWN = True  # Set to False to disable countdown display

def format_time_remaining(seconds):
    """Format seconds into a readable time string"""
    mins, secs = divmod(seconds, 60)
    if mins > 0:
        return f"{mins}m {secs}s"
    return f"{secs}s"

def cooldown_with_progress(seconds, sheet_name=None):
    """
    Pause execution with a progress indicator
    
    Args:
        seconds: Number of seconds to wait
        sheet_name: Optional name of the next sheet to process
    """
    if not SHOW_COUNTDOWN:
        print(f"⏳ Cooling down for {seconds} seconds...")
        time.sleep(seconds)
        return
    
    start_time = time.time()
    end_time = start_time + seconds
    
    next_sheet_msg = f" (next: '{sheet_name}')" if sheet_name else ""
    print(f"⏳ Cooling down for {seconds} seconds to avoid rate limits{next_sheet_msg}")
    
    # Update every 5 seconds
    update_interval = 5
    last_update = 0
    
    while True:
        elapsed = time.time() - start_time
        remaining = max(0, seconds - elapsed)
        
        if remaining <= 0:
            break
        
        # Only update display every update_interval seconds
        if elapsed - last_update >= update_interval or remaining < update_interval:
            progress = (elapsed / seconds) * 100
            bar_length = 30
            filled = int(bar_length * progress / 100)
            bar = '█' * filled + '░' * (bar_length - filled)
            
            print(f"   [{bar}] {progress:.0f}% - {format_time_remaining(int(remaining))} remaining", 
                  end='\r', flush=True)
            last_update = elapsed
        
        # Sleep for a short interval
        time.sleep(0.5)
    
    # Clear the progress line
    print(f"   ✅ Cooldown complete!{' ' * 60}")

def main():
    print("=" * 70)
    print("  📋 Google Sheets Auto-Numbering & Merging System")
    print("=" * 70)
    print(f"⏰ Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    spreadsheet_url = sheet_data['spreadsheetUrl']
    spreadsheet_id = get_spreadsheet_id(spreadsheet_url)
    service = get_sheet_service(credentials_info)
    
    try:
        print("📥 Fetching spreadsheet metadata...")
        metadata = get_sheet_metadata(service, spreadsheet_id)
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]
        
        # Filter out sheets to skip
        sheets_to_process = [name for name in sheet_names if name not in skip_sheets]
        total_sheets = len(sheets_to_process)
        
        print(f"📊 Total sheets in spreadsheet: {len(sheet_names)}")
        print(f"⏭️  Skipping sheets: {', '.join(skip_sheets) if skip_sheets else 'None'}")
        print(f"✅ Sheets to process: {total_sheets}")
        print(f"⚙️  Cooldown between sheets: {COOLDOWN_SECONDS} seconds")
        print("\n" + "-" * 70 + "\n")
        
        successful = 0
        failed = 0
        start_time = time.time()
        
        for idx, name in enumerate(sheets_to_process, 1):
            sheet_start_time = time.time()
            print(f"🔄 [{idx}/{total_sheets}] Processing: '{name}'")
            
            try:
                process_sheet(service, spreadsheet_id, sheets, name)
                successful += 1
                sheet_duration = time.time() - sheet_start_time
                print(f"   ⏱️  Completed in {sheet_duration:.1f}s")
                
                # Add cooldown after each sheet except the last one
                if idx < total_sheets:
                    next_sheet = sheets_to_process[idx] if idx < len(sheets_to_process) else None
                    cooldown_with_progress(COOLDOWN_SECONDS, next_sheet)
                    print()  # Add spacing between sheets
                    
            except Exception as sheet_error:
                failed += 1
                print(f"   ❌ Error: {sheet_error}")
                
                # Still cooldown before next sheet to avoid cascading rate limit errors
                if idx < total_sheets:
                    print(f"   ⏳ Cooling down before attempting next sheet...")
                    cooldown_with_progress(COOLDOWN_SECONDS)
                    print()
        
        # Summary
        total_duration = time.time() - start_time
        print("\n" + "=" * 70)
        print("  📊 PROCESSING SUMMARY")
        print("=" * 70)
        print(f"✅ Successful: {successful}/{total_sheets}")
        if failed > 0:
            print(f"❌ Failed: {failed}/{total_sheets}")
        print(f"⏱️  Total time: {format_time_remaining(int(total_duration))}")
        print(f"⏰ Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)
        
        if failed > 0:
            sys.exit(1)
        
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
