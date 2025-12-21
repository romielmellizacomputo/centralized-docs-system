import sys
import os
import re
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
    SHEET_SYNC_SID,
    DASHBOARD_SHEET,
    generate_timestamp_string
)

# Configuration
TC_REVIEW_SHEET = 'TC Review'
URL_COLUMN = 'H'  # Column H contains the URLs
LABEL_COLUMN = 'J'  # Column J will contain the labels

# Labels to process (matching the Apps Script)
LABELS_TO_PROCESS = [
    "To Do", "Doing", "Changes Requested",
    "Manual QA For Review", "QA Lead For Review",
    "Automation QA For Review", "Done", "On Hold", "Deprecated",
    "Automation Team For Review",
]

# Project name mapping for common variations
PROJECT_MAPPING = {
    'HQZEN': 'HQZEN',
    'BACKEND': 'BACKEND',
    'ANDROID': 'ANDROID',
    'DESKTOP': 'DESKTOP',
    'APPLYBPO': 'APPLYBPO',
    'MINISTRY': 'MINISTRY',
    'SCALEMA': 'SCALEMA',
    'BPOSEATS': 'BPOSEATS.COM'
}

def get_source_issues(sheets):
    """Get all issues from the source sheet (ALL ISSUES)"""
    print(f"📋 Fetching data from {SHEET_SYNC_SID} - ALL ISSUES!C4:T")
    
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SHEET_SYNC_SID,
        range="ALL ISSUES!C4:T"
    ).execute()
    
    values = result.get('values', [])
    if not values:
        print("⚠️ No data found in ALL ISSUES sheet")
        return []
    
    print(f"✅ Found {len(values)} rows in source sheet")
    return values

def build_source_lookup(source_data):
    """
    Build a lookup dictionary from source data
    Key format: "IID|PROJECT"
    Value: dict with issue details
    
    Column mapping from C4:T (0-indexed from C):
    C=0, D=1 (IID), E=2 (Title), F=3 (Author), G=4 (Assignee), 
    H=5 (Labels), I=6, J=7, K=8 (Status), L=9 (Created), 
    M=10, N=11 (Project), O=12, P=13, Q=14, R=15, S=16, T=17
    """
    lookup = {}
    
    for idx, row in enumerate(source_data, start=4):
        # Ensure row has enough columns
        if len(row) < 12:  # Need at least up to column N (index 11)
            continue
        
        iid = str(row[1]).strip() if len(row) > 1 and row[1] else None  # Column D
        project = str(row[11]).strip().upper() if len(row) > 11 and row[11] else None  # Column N
        
        if not iid or not project:
            continue
        
        # Build lookup key
        key = f"{iid}|{project}"
        
        lookup[key] = {
            'iid': iid,
            'title': row[2] if len(row) > 2 else 'No Title',  # Column E
            'author': row[3] if len(row) > 3 else 'Unknown',  # Column F
            'assignee': row[4] if len(row) > 4 else 'Unassigned',  # Column G
            'labels': row[5] if len(row) > 5 else 'None',  # Column H
            'status': row[8] if len(row) > 8 else 'Unknown',  # Column K
            'created_at': row[9] if len(row) > 9 else '',  # Column L
            'project': project,
            'row_index': idx
        }
    
    print(f"📊 Built lookup dictionary with {len(lookup)} issues")
    return lookup

def extract_url_from_cell(sheets, sheet_id, row, col_letter):
    """
    Extract URL from a cell, checking for hyperlinks first
    Note: Google Sheets API doesn't easily expose hyperlinks in plain values,
    so we'll work with the text content and formulas
    """
    # Try to get the formula (which might contain HYPERLINK)
    formula_range = f"{col_letter}{row}"
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TC_REVIEW_SHEET}'!{formula_range}",
        valueRenderOption='FORMULA'
    ).execute()
    
    formula_values = result.get('values', [['']])
    formula = formula_values[0][0] if formula_values and formula_values[0] else ''
    
    # Check if it's a HYPERLINK formula
    if formula and 'HYPERLINK' in formula.upper():
        # Extract URL from HYPERLINK("url", "text")
        match = re.search(r'HYPERLINK\s*\(\s*"([^"]+)"', formula, re.IGNORECASE)
        if match:
            return match.group(1)
    
    # Otherwise, get the plain value
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TC_REVIEW_SHEET}'!{formula_range}",
        valueRenderOption='UNFORMATTED_VALUE'
    ).execute()
    
    values = result.get('values', [['']])
    return values[0][0] if values and values[0] else ''

def parse_gitlab_url(url):
    """
    Parse GitLab URL to extract project name and IID
    Expected format: https://forge.bposeats.com/<group>/<project>/-/issues/<iid>
    """
    if not url:
        return None, None
    
    pattern = r'https://forge\.bposeats\.com/[^/]+/([^/]+)/-/issues/(\d+)'
    match = re.search(pattern, str(url))
    
    if not match:
        return None, None
    
    project_name = match.group(1)  # e.g., "hqzen.com"
    issue_iid = match.group(2)     # e.g., "11681"
    
    # Clean project name (remove .com, convert to uppercase)
    clean_project = project_name.replace('.com', '').upper().strip()
    
    # Apply project mapping
    clean_project = PROJECT_MAPPING.get(clean_project, clean_project)
    
    return issue_iid, clean_project

def get_tc_review_data(sheets, sheet_id):
    """Get all data from TC Review sheet starting from row 2"""
    print(f"📋 Fetching TC Review data from {sheet_id}")
    
    # Get the last row with data
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TC_REVIEW_SHEET}'!H2:H"
    ).execute()
    
    values = result.get('values', [])
    if not values:
        print("⚠️ No data found in TC Review sheet")
        return 0
    
    last_row = len(values) + 1  # +1 because we start from row 2
    print(f"✅ Found data up to row {last_row + 1}")
    return last_row + 1

def find_relevant_label(labels_str):
    """Find the first relevant label from the labels string"""
    if not labels_str:
        return None
    
    labels_list = [label.strip() for label in str(labels_str).split(',')]
    
    for label in LABELS_TO_PROCESS:
        if label in labels_list:
            return label
    
    return None

def update_tc_review_labels(sheets, sheet_id, source_lookup):
    """Update labels in TC Review sheet based on source data"""
    print(f"🔄 Processing TC Review sheet in {sheet_id}")
    
    last_row = get_tc_review_data(sheets, sheet_id)
    
    if last_row < 2:
        print("⚠️ No data to process in TC Review")
        return
    
    processed_count = 0
    updated_count = 0
    updates = []  # Collect all updates for batch processing
    
    # Get all URLs at once for efficiency
    url_range = f"'{TC_REVIEW_SHEET}'!H2:H{last_row}"
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=url_range,
        valueRenderOption='FORMULA'
    ).execute()
    
    url_values = result.get('values', [])
    
    for row_idx, row_data in enumerate(url_values, start=2):
        url = ''
        
        # Extract URL from formula or plain value
        if row_data and row_data[0]:
            cell_value = row_data[0]
            
            # Check if it's a HYPERLINK formula
            if 'HYPERLINK' in str(cell_value).upper():
                match = re.search(r'HYPERLINK\s*\(\s*"([^"]+)"', cell_value, re.IGNORECASE)
                if match:
                    url = match.group(1)
            else:
                url = str(cell_value).strip()
        
        if not url:
            print(f"Row {row_idx}: Empty URL, skipping...")
            continue
        
        print(f"Row {row_idx}: Processing URL: {url}")
        
        # Parse URL to get IID and project
        issue_iid, project_name = parse_gitlab_url(url)
        
        if not issue_iid or not project_name:
            print(f"Row {row_idx}: Invalid URL format")
            continue
        
        print(f"Row {row_idx}: IID={issue_iid}, Project={project_name}")
        
        # Look up issue in source data
        lookup_key = f"{issue_iid}|{project_name}"
        issue_data = source_lookup.get(lookup_key)
        
        if not issue_data:
            print(f"Row {row_idx}: No matching issue found for {lookup_key}")
            continue
        
        print(f"Row {row_idx}: Found matching issue - {issue_data['title']}")
        
        # Find relevant label
        relevant_label = find_relevant_label(issue_data['labels'])
        
        if relevant_label:
            # Prepare update for column J (label column)
            updates.append({
                'range': f"'{TC_REVIEW_SHEET}'!J{row_idx}",
                'values': [[relevant_label]]
            })
            
            # Update hyperlink format in column H to show #<issueIID>
            hyperlink_formula = f'=HYPERLINK("{url}", "#{issue_iid}")'
            updates.append({
                'range': f"'{TC_REVIEW_SHEET}'!H{row_idx}",
                'values': [[hyperlink_formula]]
            })
            
            print(f"Row {row_idx}: Will update with label '{relevant_label}'")
            updated_count += 1
        
        processed_count += 1
    
    # Batch update all changes
    if updates:
        print(f"📤 Applying {len(updates)} updates...")
        batch_update_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': updates
        }
        
        sheets.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id,
            body=batch_update_body
        ).execute()
        
        print(f"✅ Successfully applied all updates")
    
    print(f"📊 Processing complete: {processed_count} rows processed, {updated_count} labels updated")

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
        
        # Get source issues and build lookup from SHEET_SYNC_SID
        print(f"📋 Source: {SHEET_SYNC_SID} - ALL ISSUES")
        source_data = get_source_issues(sheets)
        
        if not source_data:
            print("❌ No source data available")
            return
        
        source_lookup = build_source_lookup(source_data)
        
        if not source_lookup:
            print("❌ Failed to build source lookup")
            return
        
        # Get all Team CDS sheet IDs from UTILS
        from common import get_all_team_cds_sheet_ids
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return
        
        # Process each sheet
        for sheet_id in sheet_ids:
            try:
                print(f"\n🔄 Processing: {sheet_id}")
                
                # Verify TC Review sheet exists
                titles = get_sheet_titles(sheets, sheet_id)
                
                if TC_REVIEW_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{TC_REVIEW_SHEET}' sheet")
                    continue
                
                # Update TC Review labels
                update_tc_review_labels(sheets, sheet_id, source_lookup)
                
                # Update timestamp in Dashboard
                update_timestamp(sheets, sheet_id)
                
                print(f"✅ Finished: {sheet_id}")
                
            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")
                import traceback
                traceback.print_exc()
        
        print("\n✅ Script completed successfully")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
