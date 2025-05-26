import os, sys, smtplib, datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from googleapiclient.discovery import build
from common import authenticate, get_sheet_ids, get_assignee_email_map, days_since
from constants import UTILS_SHEET_ID, CDS_MASTER_ROSTER

# Validate env vars
missing_vars = {"UTILS_SHEET_ID": UTILS_SHEET_ID, "CDS_MASTER_ROSTER": CDS_MASTER_ROSTER}
for name, val in missing_vars.items():
    if not val:
        print(f"❌ {name} is not set. Please set {name} environment variable.")
        sys.exit(1)

def get_priority_color(priority_label):
    lower = priority_label.lower()
    if "urg" in lower:
        return "#ffcccc"  # Red
    elif "high" in lower:
        return "#f8d7da"  # Light red
    elif "med" in lower:
        return "#e6ccff"  # Light violet
    elif "low" in lower:
        return "#d6ecff"  # Light blue
    return "#f9f9f9"     # Default light gray

def should_send_mr_reminder(row, row_number):
    row += [""] * 17  # pad row to expected length

    assigned_date = row[2].strip()
    assignee = row[3].strip()
    status = row[4].strip().lower()
    priority = row[5].strip()
    estimate = row[7].strip()
    backend_url = row[13].strip()
    frontend_url = row[14].strip()
    finished_date = row[16].strip()

    days = days_since(assigned_date)
    if days is None or days < 2:
        return False, None

    should_remind = False
    reasons = []

    if not estimate:
        reasons.append("Missing Estimation")
        should_remind = True

    if status in ["", "assigned", "on hold", "on-going discussion", "blocked"]:
        reasons.append(f"Status: '{status or 'Empty'}'")
        should_remind = True

    if status in ["passed", "failed"] and not finished_date:
        reasons.append(f"Missing Finished Date for status '{status.capitalize()}'")
        should_remind = True

    if not should_remind:
        return False, None

    task_display = f"Task ID: Row{row_number + 2} | {priority} - " \
                   f"<a href='{backend_url}' target='_blank'>Backend</a> / " \
                   f"<a href='{frontend_url}' target='_blank'>Frontend</a>"

    return True, {
        "assignee": assignee,
        "task": task_display,
        "days": days,
        "missing": reasons,
        "priority": priority
    }

def generate_mr_email_html(assignee, tasks):
    style = """
    body{font-family:sans-serif;padding:20px;background:#f5f5f5;color:#333}
    .container{background:#fff;padding:25px;border-radius:8px;max-width:700px;margin:auto;box-shadow:0 0 10px rgba(0,0,0,0.1)}
    h2{text-align:center;color:#2c3e50}
    .task{border:1px solid #ddd;padding:15px;border-radius:6px;margin:15px 0}
    .task-header{font-weight:bold;font-size:16px;margin-bottom:5px}
    .days-info{font-style:italic;color:#555;margin-bottom:8px}
    ul{margin:0;padding-left:20px}
    li{margin:4px 0}
    .footer{margin-top:30px;font-size:12px;text-align:center;color:#aaa}
    """

    body = f"<html><head><style>{style}</style></head><body><div class='container'>"
    body += f"<h2>🔔 MR Review Reminder for {assignee}</h2>"
    body += f"<p>Hello <strong>{assignee}</strong>, you have <strong>{len(tasks)} task(s)</strong> needing review:</p>"

    for task in tasks:
        issues = "".join(f"<li>{m}</li>" for m in task['missing'])
        plural = "s" if task["days"] > 1 else ""
        bg_color = get_priority_color(task['priority'])
        body += f"""
        <div class="task" style="background:{bg_color};">
            <div class="task-header">{task['task']}</div>
            <div class="days-info">Assigned <strong>{task['days']} day{plural} ago</strong></div>
            <ul>{issues}</ul>
        </div>
        """

    body += """
    <p>Please update the MR sheet to reflect the latest progress.</p>
    <div class="footer">
        This is an automated reminder sent by the CDS Monitoring System.<br>
        Kindly ignore if you’ve already updated your items.
    </div></div></body></html>
    """
    return body

def send_mr_email(assignee, tasks, recipient):
    sender, app_password = os.getenv("GMAIL_SENDER"), os.getenv("GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        print("❌ GMAIL_SENDER or GMAIL_APP_PASSWORD environment variables not set.")
        return
    if not recipient:
        print(f"❌ No email for assignee '{assignee}'. Skipping.")
        return

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = f"🔔 MR Review Reminder: {len(tasks)} Task(s) for {assignee}"
    msg.attach(MIMEText(generate_mr_email_html(assignee, tasks), "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_password)
            server.sendmail(sender, recipient, msg.as_string())
        print(f"📧 Reminder sent to {recipient}")
    except Exception as e:
        print(f"❌ Failed to send: {e}")

def process_mr_sheet(sheet_service, sheet_id, assignee_email_map):
    try:
        rows = sheet_service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="MR Review!A2:Q").execute().get("values", [])
        print(f"✅ Fetched {len(rows)} rows from MR Review")

        assignee_tasks = {}
        for idx, row in enumerate(rows):
            flag, info = should_send_mr_reminder(row, idx)
            if flag and info:
                assignee_tasks.setdefault(info["assignee"], []).append(info)

        for assignee, tasks in assignee_tasks.items():
            send_mr_email(assignee, tasks, assignee_email_map.get(assignee))

    except Exception as e:
        print(f"❌ Error processing MR sheet: {e}")

def main():
    credentials = authenticate()
    sheets = build("sheets", "v4", credentials=credentials)

    assignee_email_map = get_assignee_email_map(sheets)
    sheet_ids = get_sheet_ids(sheets)

    for sheet_id in sheet_ids:
        print(f"📄 Processing sheet ID: {sheet_id}")
        process_mr_sheet(sheets, sheet_id, assignee_email_map)

if __name__ == "__main__":
    main()
