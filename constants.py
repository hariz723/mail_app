from pathlib import Path


DATA_FILE = Path(__file__).with_name("wfh_requests.xlsx")
WORKBOOK_TITLE = "UPCOMING WFH - PLAN"
VISIBLE_HEADERS = [
    "Name",
    "Emp ID",
    "Mail ID",
    "Mon",
    "Tue",
    "Wed",
    "Thu",
    "Fri",
]
METADATA_HEADERS = [
    "Submitted At",
    "Work From Home",
    "WFH Date",
    "Week Start",
    "Week End",
    "Status",
]
SHEET_HEADERS = VISIBLE_HEADERS + METADATA_HEADERS
WEEKLY_WFH_LIMIT = 4
