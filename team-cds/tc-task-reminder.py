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


def get_sheet_urls(sheets):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=UTILS_SHEET_ID,
        range="B2:B"
    ).execute()
    values = result.get("values", [])
    return [row[0] for row in values if row]


def extract_sheet_id_from_url(url):
    try:
        return url.split("/d/")[1].split("/")[0]
    except IndexError:
        return None


def days_since(date_str):
    try:
        assigned_date = datetime.datetime.strptime(date_str, "%m/%d/%Y")
        return (datetime.datetime.now() - assigned_date).days
    except:
        return None


def should_send_reminder(row):
    assigned_date = row[2] if len(row) > 2 else ""
    task_name = row[3] if len(row) > 3 else ""
    assignee = row[4] if len(row) > 4 else ""
    task_type = row[5] if len(row) > 5 else ""
    status = row[6] if len(row) > 6 else ""
    estimate = row[8] if len(row) > 8 else ""
    date_finished = row[14] if len(row) > 14 else ""
    output_url = row[15] if len(row) > 15 else ""
    test_case_link = row[16] if len(row) > 16 else ""

    days = days_since(assigned_date)
    if days is None or days < 3:
        return False, None

    if task_type not in ["Sprint Deliverable", "Parking Lot Task"]:
        return False, None

    missing_items = []

    if not estimate.strip():
        missing_items.append("estimation")

    if status.lower() == "done":
        if not output_url.strip():
            missing_items.append("output reference")
        if not test_case_link.strip():
            missing_items.append("test case link")

    if missing_items:
        return True, {
            "assignee": assignee,
            "task": task_name,
            "days": days,
            "missing": missing_items
        }
    return False, None


def send_email(assignee, task, days, missing):
    recipient = "romiel@bposeats.com"
    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    subject = f"Task Reminder: {task}"

    body = f"Hey, {assignee}, you have a pending task: {task} for {days} days.\n\n"
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
        for row in rows:
            send_flag, info = should_send_reminder(row)
            if send_flag and info:
                send_email(
                    info["assignee"],
                    info["task"],
                    info["days"],
                    info["missing"]
                )
    except Exception as e:
        print(f"❌ Error processing sheet {sheet_id}: {str(e)}")


def main():
    credentials = authenticate()
    sheets = build("sheets", "v4", credentials=credentials)
    urls = get_sheet_urls(sheets)
    print(f"🔗 Found {len(urls)} sheet URLs")

    for url in urls:
        sheet_id = extract_sheet_id_from_url(url)
        if not sheet_id:
            print(f"⚠️ Invalid URL: {url}")
            continue
        print(f"📄 Processing sheet: {sheet_id}")
        process_sheet(sheets, sheet_id)


if __name__ == "__main__":
    main()
