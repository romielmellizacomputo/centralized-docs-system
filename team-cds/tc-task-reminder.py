import os
import sys
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Add parent directory to sys.path for custom imports if necessary
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from common import authenticate
from constants import UTILS_SHEET_ID, CDS_MASTER_ROSTER

# Validate required environment variables
if not UTILS_SHEET_ID:
    print("❌ UTILS_SHEET_ID is not set. Please set LEADS_CDS_SID environment variable.")
    sys.exit(1)

if not CDS_MASTER_ROSTER:
    print("❌ CDS_MASTER_ROSTER is not set. Please set CDS_MASTER_ROSTER environment variable.")
    sys.exit(1)

def get_sheet_ids(sheets):
    """
    Fetch sheet IDs listed in the UTILS sheet (range UTILS!B2:B).
    """
    result = sheets.spreadsheets().values().get(
        spreadsheetId=UTILS_SHEET_ID,
        range="UTILS!B2:B"
    ).execute()
    values = result.get("values", [])
    sheet_ids = [row[0].strip() for row in values if row and row[0].strip()]
    print(f"🔗 Found {len(sheet_ids)} valid sheet IDs in UTILS sheet")
    return sheet_ids

def get_assignee_email_map(sheets):
    """
    Fetch assignee-to-email mapping from the CDS_MASTER_ROSTER sheet.
    Sheet name: "Roster"
    Range: A4:B (A = assignee name, B = email)
    Returns dict {assignee_name: email}
    """
    try:
        result = sheets.spreadsheets().values().get(
            spreadsheetId=CDS_MASTER_ROSTER,
            range="Roster!A4:B"
        ).execute()
        rows = result.get("values", [])
        mapping = {}
        for row in rows:
            if len(row) >= 2:
                name = row[0].strip()
                email = row[1].strip()
                if name and email:
                    mapping[name] = email
        print(f"📋 Loaded {len(mapping)} assignee-email mappings from CDS_MASTER_ROSTER")
        return mapping
    except Exception as e:
        print(f"❌ Failed to load assignee-email mapping: {e}")
        return {}

def days_since(date_str):
    """
    Calculate days passed since date_str.
    Supports formats: "MM/DD/YYYY" and "Day, Mon DD, YYYY"
    Returns integer days or None if parsing fails.
    """
    try:
        return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%m/%d/%Y")).days
    except ValueError:
        try:
            return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%a, %b %d, %Y")).days
        except Exception as e:
            print(f"⚠️ Could not parse date '{date_str}': {e}")
            return None

def should_send_reminder(row):
    """
    Checks if a row (task) needs a reminder email.
    Returns (bool, info_dict) where info_dict contains assignee, task, days, missing items.
    """
    # Normalize row length to avoid index errors
    row += [""] * (17 - len(row))

    assigned_date = row[2].strip()
    task_name = row[3].strip()
    assignee = row[4].strip()
    task_type = row[5].strip().lower()
    status = row[6].strip().lower()
    estimate = row[8].strip()
    output_url = row[15].strip()
    test_case_link = row[16].strip()

    # Debug info
    print(f"Checking task: '{task_name}', Type: '{task_type}', Status: '{status}', Assigned Date: '{assigned_date}'")

    days = days_since(assigned_date)
    print(f"Days since assigned: {days}")

    if days is None or days < 3:
        print("  => No reminder: Task is too recent or invalid date.")
        return False, None

    # Only these task types need reminders
    if task_type not in ["sprint deliverable", "parking lot task"]:
        print("  => No reminder: Task type not applicable.")
        return False, None

    missing_items = []

    # Check missing estimation
    if not estimate:
        missing_items.append("Estimation")

    # For tasks marked done, check output and test case links
    if status == "done":
        if not output_url:
            missing_items.append("Output Reference")
        if not test_case_link:
            missing_items.append("Test Case Link")

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

def generate_email_html(assignee, tasks):
    """
    Generate an attractive HTML email body listing all pending tasks with missing info.
    """
    missing_colors = {
        "Estimation": "#e74c3c",         # Red
        "Output Reference": "#e67e22",   # Orange
        "Test Case Link": "#3498db"      # Blue
    }

    html = f"""
    <html>
    <head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f9f9f9;
            color: #333333;
            padding: 20px;
        }}
        .container {{
            background: #ffffff;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            max-width: 600px;
            margin: auto;
        }}
        h2 {{
            color: #2c3e50;
            text-align: center;
            margin-bottom: 25px;
        }}
        .task {{
            border: 1px solid #ddd;
            padding: 15px 20px;
            border-radius: 6px;
            margin-bottom: 20px;
            background-color: #fafafa;
        }}
        .task-header {{
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 8px;
            color: #34495e;
        }}
        .days-info {{
            font-style: italic;
            color: #7f8c8d;
            margin-bottom: 12px;
        }}
        ul {{
            padding-left: 20px;
            margin: 0;
        }}
        li {{
            margin-bottom: 6px;
            font-weight: 500;
        }}
        .missing-estimation {{
            color: {missing_colors['Estimation']};
            font-weight: 700;
        }}
        .missing-output {{
            color: {missing_colors['Output Reference']};
            font-weight: 700;
        }}
        .missing-testcase {{
            color: {missing_colors['Test Case Link']};
            font-weight: 700;
        }}
        .footer {{
            margin-top: 30px;
            font-size: 14px;
            text-align: center;
            color: #999999;
        }}
        .cta {{
            display: block;
            width: fit-content;
            margin: 25px auto 0;
            background-color: #27ae60;
            color: white !important;
            padding: 12px 24px;
            font-weight: 700;
            text-decoration: none;
            border-radius: 30px;
            box-shadow: 0 4px 10px rgba(39, 174, 96, 0.4);
            transition: background-color 0.3s ease;
        }}
        .cta:hover {{
            background-color: #2ecc71;
        }}
    </style>
    </head>
    <body>
    <div class="container">
        <h2>⚠️ Important: Pending Task Reminder for {assignee}</h2>
        <p>Dear <strong>{assignee}</strong>,</p>
        <p>You have <strong>{len(tasks)} pending task(s)</strong> that require your attention due to missing information. Please review the details below and update the necessary fields as soon as possible to ensure smooth progress.</p>
    """

    for task_info in tasks:
        task = task_info["task"]
        days = task_info["days"]
        missing = task_info["missing"]

        html += f"""
        <div class="task">
            <div class="task-header">{task}</div>
            <div class="days-info">Assigned <strong>{days} day{'s' if days > 1 else ''} ago</strong></div>
            <ul>
        """

        for miss in missing:
            class_name = ""
            if miss == "Estimation":
                class_name = "missing-estimation"
            elif miss == "Output Reference":
                class_name = "missing-output"
            elif miss == "Test Case Link":
                class_name = "missing-testcase"

            html += f'<li class="{class_name}">Missing {miss}</li>'

        html += "</ul></div>"

    html += """
        <p>Please update the missing information at your earliest convenience to avoid any delays.</p>
        <a href="https://docs.google.com/spreadsheets/d/" target="_blank" class="cta">Update Your Tasks Now</a>
        <p class="footer">Thank you for your prompt attention and dedication!<br>— The TC Task Management Team</p>
    </div>
    </body>
    </html>
    """
    return html

def send_email_combined(assignee, tasks, assignee_email):
    """
    Sends one combined email to the assignee listing all their pending tasks.
    """
    if not assignee_email:
        print(f"❌ No email found for assignee '{assignee}'. Skipping email.")
        return

    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")

    if not sender or not app_password:
        print("❌ GMAIL_SENDER or GMAIL_APP_PASSWORD environment variables not set.")
        return

    recipient = assignee_email
    subject = f"⏰ TC Task Reminder: {len(tasks)} Pending Task(s) Need Your Attention!"

    body_html = generate_email_html(assignee, tasks)

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body_html, "html"))

    print(f"Attempting to send combined HTML email from {sender} to {recipient} for assignee '{assignee}'")
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_password)
            server.sendmail(sender, recipient, msg.as_string())
        print(f"📧 Combined email sent to {recipient} for assignee '{assignee}' with {len(tasks)} tasks")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

def process_sheet(sheet_service, sheet_id, assignee_email_map):
    """
    Process each sheet by reading task rows,
    grouping tasks by assignee that require reminders,
    then sending emails accordingly.
    """
    try:
        result = sheet_service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range="Test Cases!A2:Q"
        ).execute()
        rows = result.get("values", [])
        print(f"Fetched {len(rows)} rows from sheet ID: {sheet_id}")

        assignee_tasks = {}

        for row in rows:
            send_flag, info = should_send_reminder(row)
            if send_flag and info:
                assignee = info["assignee"]
                if assignee not in assignee_tasks:
                    assignee_tasks[assignee] = []
                assignee_tasks[assignee].append(info)
            else:
                # No reminder needed for this row
                pass

        # Send emails per assignee
        for assignee, tasks in assignee_tasks.items():
            assignee_email = assignee_email_map.get(assignee)
            send_email_combined(assignee, tasks, assignee_email)

    except Exception as e:
        print(f"❌ Error processing sheet {sheet_id}: {str(e)}")

def main():
    credentials = authenticate()
    sheets = build("sheets", "v4", credentials=credentials)

    # Get assignee-to-email mapping once
    assignee_email_map = get_assignee_email_map(sheets)

    # Get sheet IDs to process
    sheet_ids = get_sheet_ids(sheets)

    # Process each sheet
    for sheet_id in sheet_ids:
        print(f"📄 Processing sheet: {sheet_id}")
        process_sheet(sheets, sheet_id, assignee_email_map)

if __name__ == "__main__":
    main()
