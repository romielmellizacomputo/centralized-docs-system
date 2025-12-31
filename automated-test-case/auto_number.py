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
REQUESTS_PER_SHEET_ESTIMATE = 5  # Estimated API calls per sheet (adjust based on observation)
MAX_REQUESTS_PER_MINUTE = 60  # Google Sheets API limit
SAFETY_MARGIN = 0.9  # Use 90% of the limit to be safe
COOLDOWN_SECONDS = max(1, int((REQUESTS_PER_SHEET_ESTIMATE * 60) / (MAX_REQUESTS_PER_MINUTE * SAFETY_MARGIN)))
SHOW_COUNTDOWN = True  # Set to False to disable countdown display

# Dynamic rate limiting tracker
class RateLimiter:
    def __init__(self, max_requests_per_minute=60):
        self.max_requests = max_requests_per_minute * SAFETY_MARGIN
        self.request_times = []
        
    def add_request(self, count=1):
        """Record API request(s)"""
        current_time = time.time()
        self.request_times.extend([current_time] * count)
        # Clean up old requests (older than 60 seconds)
        self.request_times = [t for t in self.request_times if current_time - t < 60]
    
    def get_required_wait(self):
        """Calculate how long to wait before next request"""
        current_time = time.time()
        # Remove requests older than 60 seconds
        self.request_times = [t for t in self.request_times if current_time - t < 60]
        
        if len(self.request_times) < self.max_requests:
            return 0
        
        # Wait until oldest request is 60 seconds old
        oldest_request = min(self.request_times)
        wait_time = 60 - (current_time - oldest_request)
        return max(0, wait_time + 1)  # Add 1 second buffer
    
    def get_current_rate(self):
        """Get current requests per minute"""
        current_time = time.time()
        self.request_times = [t for t in self.request_times if current_time - t < 60]
        return len(self.request_times)

rate_limiter = RateLimiter(MAX_REQUESTS_PER_MINUTE)

def format_time_remaining(seconds):
    """Format seconds into a readable time string"""
    mins, secs = divmod(int(seconds), 60)
    if mins > 0:
        return f"{mins}m {secs}s"
    return f"{secs}s"

def cooldown_with_progress(seconds, sheet_name=None, reason="rate limits"):
    """
    Pause execution with a progress indicator
    
    Args:
        seconds: Number of seconds to wait
        sheet_name: Optional name of the next sheet to process
        reason: Reason for cooldown
    """
    if seconds < 1:
        return
        
    if not SHOW_COUNTDOWN:
        print(f"⏳ Cooling down for {format_time_remaining(seconds)}...")
        time.sleep(seconds)
        return
    
    start_time = time.time()
    
    next_sheet_msg = f" (next: '{sheet_name}')" if sheet_name else ""
    print(f"⏳ Cooling down for {format_time_remaining(seconds)} to avoid {reason}{next_sheet_msg}")
    
    # Update every 2 seconds for shorter waits, 5 seconds for longer
    update_interval = 2 if seconds < 10 else 5
    last_update = 0
    
    while True:
        elapsed = time.time() - start_time
        remaining = max(0, seconds - elapsed)
        
        if remaining <= 0:
            break
        
        # Only update display at intervals
        if elapsed - last_update >= update_interval or remaining < update_interval:
            progress = (elapsed / seconds) * 100
            bar_length = 30
            filled = int(bar_length * progress / 100)
            bar = '█' * filled + '░' * (bar_length - filled)
            
            print(f"   [{bar}] {progress:.0f}% - {format_time_remaining(remaining)} remaining", 
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
        rate_limiter.add_request(1)  # Count metadata request
        
        sheets = metadata.get('sheets', [])
        sheet_names = [s['properties']['title'] for s in sheets]
        
        # Filter out sheets to skip
        sheets_to_process = [name for name in sheet_names if name not in skip_sheets]
        total_sheets = len(sheets_to_process)
        
        print(f"📊 Total sheets in spreadsheet: {len(sheet_names)}")
        print(f"⏭️  Skipping sheets: {', '.join(skip_sheets) if skip_sheets else 'None'}")
        print(f"✅ Sheets to process: {total_sheets}")
        print(f"⚙️  Estimated requests per sheet: {REQUESTS_PER_SHEET_ESTIMATE}")
        print(f"⚙️  Base cooldown: {COOLDOWN_SECONDS}s (dynamically adjusted)")
        print("\n" + "-" * 70 + "\n")
        
        successful = 0
        failed = 0
        start_time = time.time()
        
        for idx, name in enumerate(sheets_to_process, 1):
            # Check if we need to wait before processing
            wait_time = rate_limiter.get_required_wait()
            if wait_time > 0:
                current_rate = rate_limiter.get_current_rate()
                print(f"⚠️  Rate limit approaching ({current_rate} requests in last 60s)")
                cooldown_with_progress(wait_time, name, "rate limit")
                print()
            
            sheet_start_time = time.time()
            current_rate = rate_limiter.get_current_rate()
            print(f"🔄 [{idx}/{total_sheets}] Processing: '{name}' (rate: {current_rate} req/min)")
            
            try:
                process_sheet(service, spreadsheet_id, sheets, name)
                rate_limiter.add_request(REQUESTS_PER_SHEET_ESTIMATE)  # Track estimated requests
                
                successful += 1
                sheet_duration = time.time() - sheet_start_time
                print(f"   ⏱️  Completed in {sheet_duration:.1f}s")
                
                # Add small cooldown between sheets (only if not already waiting for rate limit)
                if idx < total_sheets:
                    next_sheet = sheets_to_process[idx] if idx < len(sheets_to_process) else None
                    
                    # Check if we're approaching rate limit
                    projected_rate = rate_limiter.get_current_rate() + REQUESTS_PER_SHEET_ESTIMATE
                    if projected_rate > rate_limiter.max_requests * 0.8:  # 80% threshold
                        wait_time = rate_limiter.get_required_wait()
                        if wait_time > 0:
                            cooldown_with_progress(wait_time, next_sheet, "rate limit prevention")
                        else:
                            cooldown_with_progress(COOLDOWN_SECONDS, next_sheet, "safety buffer")
                    else:
                        # Just a brief pause for processing
                        time.sleep(1)
                    print()
                    
            except Exception as sheet_error:
                failed += 1
                error_msg = str(sheet_error)
                print(f"   ❌ Error: {error_msg}")
                
                # If rate limit error, wait longer
                if "quota" in error_msg.lower() or "rate" in error_msg.lower():
                    print(f"   ⚠️  Rate limit hit! Backing off...")
                    cooldown_with_progress(65, None, "rate limit recovery")
                    rate_limiter.request_times = []  # Reset tracker
                elif idx < total_sheets:
                    # Regular cooldown for other errors
                    cooldown_with_progress(COOLDOWN_SECONDS, None, "error recovery")
                print()
        
        # Summary
        total_duration = time.time() - start_time
        avg_time_per_sheet = total_duration / total_sheets if total_sheets > 0 else 0
        
        print("\n" + "=" * 70)
        print("  📊 PROCESSING SUMMARY")
        print("=" * 70)
        print(f"✅ Successful: {successful}/{total_sheets}")
        if failed > 0:
            print(f"❌ Failed: {failed}/{total_sheets}")
        print(f"⏱️  Total time: {format_time_remaining(total_duration)}")
        print(f"⏱️  Average per sheet: {avg_time_per_sheet:.1f}s")
        print(f"📊 Total API requests (estimated): {rate_limiter.get_current_rate()}")
        print(f"⏰ Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)
        
        if failed > 0:
            sys.exit(1)
        
    except Exception as e:
        print(f"\n❌ FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
