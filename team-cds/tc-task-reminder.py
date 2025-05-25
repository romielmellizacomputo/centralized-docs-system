import os
import sys
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from common import authenticate
from constants import UTILS_SHEET_ID, CDS_MASTER_ROSTER

if not UTILS_SHEET_ID:
    print("❌ UTILS_SHEET_ID is not set. Please set LEADS_CDS_SID environment variable.")
    sys.exit(1)

if not CDS_MASTER_ROSTER:
    print("❌ CDS_MASTER_ROSTER is not set. Please set CDS_MASTER_ROSTER environment variable.")
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

def get_assignee_email_map(sheets):
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
    try:
        return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%m/%d/%Y")).days
    except ValueError:
        try:
            return (datetime.datetime.now() - datetime.datetime.strptime(date_str, "%a, %b %d, %Y")).days
        except Exception as e:
            print(f"⚠️ Could not parse date '{date_str}': {e}")
            return None

def should_send_reminder(row):
    row += [""] * (17 - len(row))

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
        missing_items.append("Estimation")

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

def generate_email_html(assignee, tasks, sheet_id):
    missing_colors = {
        "Estimation": "#e74c3c",
        "Output Reference": "#e67e22",
        "Test Case Link": "#3498db"
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
            display: inline-block;
            background-color: #27ae60;
            color: white !important;
            padding: 10px 20px;
            font-weight: 700;
            text-decoration: none;
            border-radius: 30px;
            box-shadow: 0 4px 10px rgba(39, 174, 96, 0.4);
            transition: background-color 0.3s ease;
            margin-top: 15px;
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
        row_number = task_info["row_number"]

        # Construct URL to the specific row in the Google Sheet
        # Note: Google Sheets rows start at 1, header row is 1, tasks start at row 2, so row_number should be correct.
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit#gid=0&range=A{row_number}"

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

        html += f"""
            </ul>
            <a href="{sheet_url}" target="_blank" class="cta">Update Task Now</a>
        </div>
        """

    html += """
        <p>Please update the missing information at your earliest convenience to avoid any delays.</p>
        <p class="footer">Thank you for your prompt attention and dedication!<br>— The TC Task Management Team</p>
    </div>
    </body>
    </html>
    """
    return html

def send_email_combined(assignee, tasks, assignee_email, sheet_id):
    try:
        from_email = os.environ.get("EMAIL_FROM")
        email_password = os.environ.get("EMAIL_PASSWORD")
        smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", 587))

        if not from_email or not email_password:
            print("❌ Email credentials not set in environment variables")
            return

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "⏳ Reminder: Update Your Pending Tasks"
        msg["From"] = from_email
        msg["To"] = assignee_email

        html_content = generate_email_html(assignee, tasks, sheet_id)
        part2 = MIMEText(html_content, "html")
        msg.attach(part2)

        server = smtplib.SMTP(smtp_host, smtp_port)
        server.starttls()
        server.login(from_email, email_password)
        server.sendmail(from_email, assignee_email, msg.as_string())
        server.quit()

        print(f"✅ Email sent to {assignee} ({assignee_email}) with {len(tasks)} tasks")
    except Exception as e:
        print(f"❌ Failed to send email to {assignee}: {e}")

def main():
    sheets = authenticate()
    sheet_ids = get_sheet_ids(sheets)
    assignee_email_map = get_assignee_email_map(sheets)

    for sheet_id in sheet_ids:
        try:
            result = sheets.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range="Task Tracking!A2:Q"
            ).execute()

            rows = result.get("values", [])
            pending_tasks = {}

            for i, row in enumerate(rows, start=2):  # start=2 because row 1 is header
                send_reminder, details = should_send_reminder(row)
                if send_reminder:
                    assignee = details["assignee"]
                    if assignee not in pending_tasks:
                        pending_tasks[assignee] = []
                    # Add row number to details for direct link
                    details["row_number"] = i
                    pending_tasks[assignee].append(details)

            for assignee, tasks in pending_tasks.items():
                email = assignee_email_map.get(assignee)
                if email:
                    send_email_combined(assignee, tasks, email, sheet_id)
                else:
                    print(f"⚠️ No email found for assignee '{assignee}', skipping email")

        except Exception as e:
            print(f"❌ Failed to process sheet {sheet_id}: {e}")

if __name__ == "__main__":
    main()
