import os
import sys
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from common import authenticate
from constants import UTILS_SHEET_ID

if not UTILS_SHEET_ID:
    print("❌ UTILS_SHEET_ID is not set. Please set LEADS_CDS_SID environment variable.")
    sys.exit(1)

def get_sheet_ids(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=UTILS_SHEET_ID,
        range="UTILS!B2:B"
    ).execute()
    values = result.get("values", [])
    sheet_ids = [row[0].strip() for row in values if row and row[0].strip()]
    print(f"🔗 Found {len(sheet_ids)} valid sheet IDs in UTILS sheet")
    return sheet_ids

def days_since(date_str):
    try:
        # Try your original format first
        return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%m/%d/%Y")).days
    except ValueError:
        try:
            # Try this new format for strings like "Mon, Mar 17, 2025"
            return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%a, %b %d, %Y")).days
        except Exception as e:
            print(f"⚠️ Could not parse date '{date_str}': {e}")
            return None

def should_send_reminder(row):
    # Defensive: fill empty fields with empty string
    row += [""] * (17 - len(row))  # ensure at least 17 elements

    assigned_date = row[2].strip()
    task_name = row[3].strip()
    assignee = row[4].strip()
    task_type = row[5].strip().lower()
    status = row[6].strip().lower()
    estimate = row[8].strip()
    output_url = row[15].strip()
    test_case_link = row[16].strip()

    print(f"Checking task: '{task_name}', Type: '{task_type}', Status: '{status}', Assigned Date: '{assigned_date}'")

    days = days_since(assigned_date)
    print(f"Days since assigned: {days}")

    if days is None or days < 3:
        print("  => No reminder: Task is too recent or invalid date.")
        return False, None

    if task_type not in ["sprint deliverable", "parking lot task"]:
        print("  => No reminder: Task type not applicable.")
        return False, None

    missing_items = []

    if not estimate:
        missing_items.append("estimation")

    if status == "done":
        if not output_url:
            missing_items.append("output reference")
        if not test_case_link:
            missing_items.append("test case link")

    if missing_items:
        print(f"  => Reminder needed for missing: {missing_items}")
        return True, {
            "assignee": assignee,
            "task": task_name,
            "days": days,
            "missing": missing_items
        }

    print("  => No reminder: All required fields are present.")
    return False, None

def send_email(assignee, task, days, missing):
    recipient = "romiel@bposeats.com"
    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    subject = f"Task Reminder: {task}"

    body = f"Hey, {assignee}, you have a pending task: '{task}' for {days} days.\n\n"
    if "estimation" in missing:
        body += "- You missed declaring estimation on your task.\n"
    if "output reference" in missing:
        body += "- You missed declaring your output reference.\n"
    if "test case link" in missing:
        body += "- You missed declaring your test case sheet link.\n"

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    print(f"Attempting to send email from {sender} to {recipient} regarding task '{task}'")
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_password)
            server.sendmail(sender, recipient, msg.as_string())
            print(f"📧 Email sent to {recipient} for task '{task}'")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

def process_sheet(sheet_service, sheet_id):
    try:
        result = sheet_service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range="Test Cases!A2:Q"
        ).execute()
        rows = result.get("values", [])
        print(f"Fetched {len(rows)} rows from sheet ID: {sheet_id}")

        for row in rows:
            send_flag, info = should_send_reminder(row)
            if send_flag and info:
                send_email(
                    info["assignee"],
                    info["task"],
                    info["days"],
                    info["missing"]
                )
            else:
                print("No email triggered for this row.")
    except Exception as e:
        print(f"❌ Error processing sheet {sheet_id}: {str(e)}")

def main():
    credentials = authenticate()
    sheets = build("sheets", "v4", credentials=credentials)
    sheet_ids = get_sheet_ids(sheets)

    for sheet_id in sheet_ids:
        print(f"📄 Processing sheet: {sheet_id}")
        process_sheet(sheets, sheet_id)

if __name__ == "__main__":
    main()
