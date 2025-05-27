from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from constants import SCOPES

def get_sheet_service(credentials_info):
    creds = Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)
