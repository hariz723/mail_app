from pathlib import Path


DATA_DIR = Path(__file__).parent
WORKBOOK_TITLE = "UPCOMING WFH - PLAN"
EMPLOYEES = [
    {"name": "ARUNSURESH M", "tid": "T2116A0", "email": ""},
    {"name": "BALAMURUGAN A", "tid": "T0094MD", "email": ""},
    {"name": "CHANDRA SEKHAR P", "tid": "T0010Q0", "email": ""},
    {"name": "KARTHIKEYAN T", "tid": "T0244HS", "email": ""},
    {"name": "PRAKASH H", "tid": "T0133F2", "email": ""},
    {"name": "SATHISHKUMAR M", "tid": "T0221AF", "email": ""},
    {"name": "PRAVINKUMAR S", "tid": "T0122B4", "email": ""},
    {"name": "RAJKUMAR K", "tid": "T0075OC", "email": ""},
    {"name": "VINOD KUMAR K V", "tid": "T0098TL", "email": ""},
    {"name": "ASUWATH KUMAR J", "tid": "T0102D2", "email": ""},
    {"name": "SUNDARA PANDIAN", "tid": "T9206SC", "email": ""},
    {"name": "MEGHAVINA", "tid": "", "email": ""},
    {"name": "SNEHA", "tid": "T0452VW", "email": ""},
    {"name": "KARTEEK KAILASAKOTA", "tid": "T5756KK", "email": ""},
    {"name": "DINESH", "tid": "TA25886", "email": ""},
    {"name": "VASU KARTHIKEYAN", "tid": "TA29523", "email": ""},
    {"name": "ARUL VIGNESH", "tid": "TA32967", "email": ""},
    {"name": "SURYA", "tid": "TA31462", "email": ""},
    {"name": "PRAVIN B", "tid": "", "email": ""},
    {"name": "SARAVANAN V", "tid": "", "email": ""},
]
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
