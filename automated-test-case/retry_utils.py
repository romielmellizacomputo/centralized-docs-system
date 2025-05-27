import time
from googleapiclient.errors import HttpError

def update_values_with_retry(service, spreadsheet_id, range_, values):
    attempts = 0
    while True:
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_,
                valueInputOption='USER_ENTERED',
                body={'values': values}
            ).execute()
            return
        except HttpError as e:
            if e.resp.status == 429:
                attempts += 1
                wait_time = (2 ** attempts)
                print(f"Quota exceeded. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise

def execute_with_retries(request_fn, max_retries=5, base_delay=5):
    for attempt in range(max_retries):
        try:
            return request_fn()
        except HttpError as e:
            if e.resp.status in [429, 500, 503]:
                wait_time = base_delay * (2 ** attempt)
                print(f"⚠️ Rate limited or server error. Retrying in {wait_time} seconds... (Attempt {attempt + 1})")
                time.sleep(wait_time)
            else:
                print("❌ Non-retryable error encountered:")
                raise
    raise Exception("❌ Exceeded maximum retries due to rate limits or server errors.")
