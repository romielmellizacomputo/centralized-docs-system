import os
import json
import sys

def load_json_env_var(var_name):
    raw = os.environ.get(var_name)
    if not raw:
        print(f"❌ ERROR: Environment variable '{var_name}' is not set or empty.")
        sys.exit(1)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        print(f"❌ ERROR: Environment variable '{var_name}' contains invalid JSON.")
        sys.exit(1)

sheet_data = load_json_env_var('SHEET_DATA')
credentials_info = load_json_env_var('TEST_CASE_SERVICE_ACCOUNT_JSON')
