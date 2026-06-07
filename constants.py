from pathlib import Path


DATA_DIR = Path(__file__).parent
WORKBOOK_TITLE = "UPCOMING WFH - PLAN"
EMPLOYEES = []
VISIBLE_HEADERS = [
    "Name",
    "TID",
    "Mon",
    "Tue",
    "Wed",
    "Thu",
    "Fri",
]
METADATA_HEADERS = [
    "Submitted At",
    "Mail ID",
    "Week Start",
    "Week End",
    "Status",
]
SHEET_HEADERS = VISIBLE_HEADERS + METADATA_HEADERS
DAILY_WFH_LIMIT = 4
