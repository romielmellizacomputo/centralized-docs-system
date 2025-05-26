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

def should_send_reminder(row):
    row += [""] * (17 - len(row))
    assigned_date, task_name, assignee, task_type, status, estimate, output_url, test_case_link = \
        row[2].strip(), row[3].strip(), row[4].strip(), row[5].strip().lower(), row[6].strip().lower(), row[8].strip(), row[15].strip(), row[16].strip()

    print(f"Checking task: '{task_name}', Type: '{task_type}', Status: '{status}', Assigned Date: '{assigned_date}'")
    days = days_since(assigned_date)
    print(f"Days since assigned: {days}")
    if days is None or days < 3 or task_type not in ["sprint deliverable", "parking lot task"]:
        print("  => No reminder: Task is too recent, invalid date, or type not applicable.")
        return False, None

    missing = []
    if not estimate: missing.append("Estimation")
    if status == "done":
        if not output_url: missing.append("Output Reference")
        if not test_case_link: missing.append("Test Case Link")
    if missing:
        print(f"  => Reminder needed for missing: {missing}")
        return True, {"assignee": assignee, "task": task_name, "days": days, "missing": missing}
    print("  => No reminder: All required fields are present.")
    return False, None

def generate_email_html(assignee, tasks):
    colors = {
        "Estimation": "#e74c3c",
        "Output Reference": "#e67e22",
        "Test Case Link": "#3498db"
    }
    get_class = lambda m: f'missing-{"estimation" if m=="Estimation" else "output" if m=="Output Reference" else "testcase"}'

    style = "\n".join([
        "body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f9f9f9;color:#333;padding:20px}",
        ".container{background:#fff;padding:25px;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.1);max-width:600px;margin:auto}",
        "h2{color:#2c3e50;text-align:center;margin-bottom:25px}",
        ".task-table{width:100%;border-spacing:20px 20px;border-collapse:separate}",
        ".task{border:1px solid #ddd;padding:15px;border-radius:6px;background:#fafafa;width:100%;}",
        ".task-header{font-size:18px;font-weight:600;margin-bottom:8px;color:#34495e}",
        ".days-info{font-style:italic;color:#7f8c8d;margin-bottom:12px}",
        "ul{padding-left:20px;margin:0}li{margin-bottom:6px;font-weight:500}",
        f".missing-estimation{{color:{colors['Estimation']};font-weight:700}}",
        f".missing-output{{color:{colors['Output Reference']};font-weight:700}}",
        f".missing-testcase{{color:{colors['Test Case Link']};font-weight:700}}",
        ".footer{margin-top:30px;font-size:14px;text-align:center;color:#999}",
        ".cta{display:block;margin:25px auto 0;background:#27ae60;color:white!important;padding:12px 24px;font-weight:700;text-decoration:none;border-radius:30px;box-shadow:0 4px 10px rgba(39,174,96,0.4);transition:0.3s}.cta:hover{background:#2ecc71}",

        # Responsive layout
        "@media only screen and (max-width: 600px){ .task-table td{display:block;width:100%!important} }"
    ])

    # Generate task cells in table format (2 per row)
    task_cells = []
    for i, t in enumerate(tasks):
        plural = "s" if t["days"] > 1 else ""
        items = "".join([f'<li class="{get_class(m)}">Missing {m}</li>' for m in t["missing"]])
        html = f"""
        <td>
            <div class="task">
                <div class="task-header">{t["task"]}</div>
                <div class="days-info">Assigned <strong>{t["days"]} day{plural} ago</strong></div>
                <ul>{items}</ul>
            </div>
        </td>
        """
        task_cells.append(html)

    # Group every 2 cells into one <tr>
    rows_html = ""
    for i in range(0, len(task_cells), 2):
        first = task_cells[i]
        second = task_cells[i+1] if i + 1 < len(task_cells) else "<td></td>"
        rows_html += f"<tr>{first}{second}</tr>"

    task_table_html = f'<table class="task-table">{rows_html}</table>'

    return f"""<html><head><style>{style}</style></head><body><div class="container">
    <h2>⚠️ Important: Pending Task Reminder for {assignee}</h2>
    <p>Dear <strong>{assignee}</strong>,</p>
    <p>You have <strong>{len(tasks)} pending task(s)</strong> requiring your attention. Please review below and update missing fields.</p>
    {task_table_html}
    <p>Please update the missing information at your earliest convenience to avoid delays.</p>
    <a href="https://drive.google.com/drive/u/0/folders/1X7tChdqEcO_RvOl617W_haZ0ea7nl36m" target="_blank" class="cta">Update Your Tasks Now</a>
    <p class="footer">Thanks for your dedication!<br>— The TC Task Management Team</p>
    </div></body></html>"""

def send_email_combined(assignee, tasks, recipient):
    sender, app_password = os.getenv("GMAIL_SENDER"), os.getenv("GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        print("❌ GMAIL_SENDER or GMAIL_APP_PASSWORD environment variables not set.")
        return
    if not recipient:
        print(f"❌ No email for assignee '{assignee}'. Skipping.")
        return

    msg = MIMEMultipart("alternative")
    msg["From"], msg["To"], msg["Subject"] = sender, recipient, f"📜 TC Task Reminder: {len(tasks)} Pending Task(s)"
    msg.attach(MIMEText(generate_email_html(assignee, tasks), "html"))

    print(f"Sending email to {recipient} for '{assignee}' with {len(tasks)} tasks...")
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_password)
            server.sendmail(sender, recipient, msg.as_string())
        print(f"📧 Email sent to {recipient}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

def process_sheet(sheet_service, sheet_id, assignee_email_map):
    try:
        rows = sheet_service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="Test Cases!A2:Q").execute().get("values", [])
        print(f"Fetched {len(rows)} rows from sheet ID: {sheet_id}")

        assignee_tasks = {}
        for row in rows:
            flag, info = should_send_reminder(row)
            if flag and info:
                assignee_tasks.setdefault(info["assignee"], []).append(info)

        for assignee, tasks in assignee_tasks.items():
            send_email_combined(assignee, tasks, assignee_email_map.get(assignee))
    except Exception as e:
        print(f"❌ Error processing sheet: {e}")


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
