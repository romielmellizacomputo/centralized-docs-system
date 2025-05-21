
from ..constants import (
    UTILS_SHEET_ID,
    G_MILESTONES,
    G_ISSUES_SHEET,
    G_MR_SHEET,
    NTC_SHEET,
    DASHBOARD_SHEET,
    CENTRAL_ISSUE_SHEET_ID,
    ALL_ISSUES,
    ALL_MR,
    ALL_NTC,
    generate_timestamp_string,
)


CONFIG = {
    "issues": {
        "range": ALL_ISSUES,
        "sheet_name": G_ISSUES_SHEET,
        "max_length": 18,
        "label_index": 6,
    },
    "mr": {
        "range": ALL_MR,
        "sheet_name": G_MR_SHEET,
        "max_length": 17,
        "label_index": 7,
    },
    "ntc": {
        "range": ALL_NTC,
        "sheet_name": NTC_SHEET,
        "max_length": 14,
        "label_index": 7,
        "filter_labels": [
            'needs test case',
            'needs test scenario',
            'test case needs update',
        ]
    }
}
