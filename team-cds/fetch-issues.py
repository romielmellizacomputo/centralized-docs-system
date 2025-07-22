import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from googleapiclient.discovery import build
from common import (
    authenticate,
    get_sheet_titles,
    get_all_team_cds_sheet_ids,
    get_selected_milestones
)
from constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_ISSUES_SHEET,
    DASHBOARD_SHEET,
    CBS_ID,  # Updated to use CBS_ID instead of CENTRAL_ISSUE_SHEET_ID
    GITLAB_ISSUES,  # Updated to use GITLAB_ISSUES instead of ALL_ISSUES
    generate_timestamp_string
)

def get_all_issues(sheets):
    """Get all issues from the new GitLab Issues source (CBS_ID)"""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=CBS_ID,
        range=GITLAB_ISSUES
    ).execute()
    
    values = result.get('values', [])
    if not values:
        raise Exception(f"No data found in range {GITLAB_ISSUES}")
    
    return values

def clear_g_issues(sheets, sheet_id):
    """Clear existing data in G-Issues sheet"""
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=f'{G_ISSUES_SHEET}!C4:U'
    ).execute()

def pad_row_to_u(row):
    """Pad row to 19 columns (C to U = 19 columns)"""
    full_length = 19
    return row + [''] * (full_length - len(row))

def insert_data_to_g_issues(sheets, sheet_id, data):
    """Insert processed data to G-Issues sheet"""
    padded_data = [pad_row_to_u(row) for row in data]
    print(f"📤 Inserting {len(padded_data)} rows to {G_ISSUES_SHEET}!C4")
    
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{G_ISSUES_SHEET}!C4',
        valueInputOption='RAW',
        body={'values': padded_data}
    ).execute()

def update_timestamp(sheets, sheet_id):
    """Update timestamp in Dashboard sheet"""
    timestamp = generate_timestamp_string()
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f'{DASHBOARD_SHEET}!W6',
        valueInputOption='RAW',
        body={'values': [[timestamp]]}
    ).execute()

def main():
    try:
        credentials = authenticate()
        sheets = build('sheets', 'v4', credentials=credentials)
        
        # Get sheet titles for validation
        get_sheet_titles(sheets, UTILS_SHEET_ID)
        sheet_ids = get_all_team_cds_sheet_ids(sheets, UTILS_SHEET_ID)
        
        if not sheet_ids:
            print("❌ No Team CDS sheet IDs found in UTILS!B2:B")
            return
        
        for sheet_id in sheet_ids:
            try:
                print(f"🔄 Processing: {sheet_id}")
                
                # Check if required sheets exist
                titles = get_sheet_titles(sheets, sheet_id)
                
                if G_MILESTONES not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_MILESTONES}' sheet")
                    continue
                
                if G_ISSUES_SHEET not in titles:
                    print(f"⚠️ Skipping {sheet_id} — missing '{G_ISSUES_SHEET}' sheet")
                    continue
                
                # Get selected milestones for filtering
                milestones = get_selected_milestones(sheets, sheet_id, G_MILESTONES)
                
                # Get all issues from the new GitLab source
                issues_data = get_all_issues(sheets)
                
                # Filter by milestones (Milestone is now in column F, index 5 in source data)
                filtered = [row for row in issues_data if len(row) > 5 and row[5] in milestones]
                
                # Map source columns to target columns
                processed = []
                for row in filtered:
                    # Ensure row has enough columns
                    padded_row = row + [''] * (22 - len(row))  # Pad to W column (22 total)
                    
                    # Map columns from source to target format
                    mapped_row = [
                        padded_row[1],   # C: ID (from B)
                        '',              # D: IID (skip - not in source)
                        padded_row[2],   # E: Issue Title (from C)
                        padded_row[3],   # F: Issue Author (from D)
                        '',              # G: Assignee (skip - not in source)
                        padded_row[4],   # H: Labels (from E)
                        padded_row[5],   # I: Milestone (from F)
                        padded_row[6],   # J: Status (from G)
                        padded_row[7],   # K: Created At (from H)
                        padded_row[8],   # L: Closed At (from I)
                        padded_row[10],  # M: Closed By (from K)
                        padded_row[9],   # N: Project (from J)
                        padded_row[17],  # O: Reviewer (from R)
                        padded_row[19],  # P: Reopened? (from T)
                        '',              # Q: Reopened By (skip - not in source)
                        '',              # R: Date Reopened (skip - not in source)
                        '',              # S: Local Status (skip - not in source)
                        '',              # T: Status Date (skip - not in source)
                        ''               # U: Duration/Status (skip - not in source)
                    ]
                    processed.append(mapped_row)
                
                # Clear existing data and insert new data
                clear_g_issues(sheets, sheet_id)
                insert_data_to_g_issues(sheets, sheet_id, processed)
                update_timestamp(sheets, sheet_id)
                
                print(f"✅ Finished: {sheet_id}")
                
            except Exception as e:
                print(f"❌ Error processing {sheet_id}: {str(e)}")
                
    except Exception as e:
        print(f"❌ Main failure: {str(e)}")

if __name__ == '__main__':
    main()
