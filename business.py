from dataclasses import dataclass
from datetime import date, datetime, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from constants import (
    DATA_FILE,
    METADATA_HEADERS,
    SHEET_HEADERS,
    VISIBLE_HEADERS,
    WEEKLY_WFH_LIMIT,
    WORKBOOK_TITLE,
)


@dataclass
class RequestResult:
    message: str
    category: str = "success"


METADATA_START_COLUMN = len(VISIBLE_HEADERS) + 1
STATUS_COLUMN = METADATA_START_COLUMN + METADATA_HEADERS.index("Status")
WEEK_START_COLUMN = METADATA_START_COLUMN + METADATA_HEADERS.index("Week Start")
WEEK_END_COLUMN = METADATA_START_COLUMN + METADATA_HEADERS.index("Week End")
WORK_FROM_HOME_COLUMN = METADATA_START_COLUMN + METADATA_HEADERS.index("Work From Home")
NAME_COLUMN = VISIBLE_HEADERS.index("Name") + 1
EMP_ID_COLUMN = VISIBLE_HEADERS.index("Emp ID") + 1
EMAIL_COLUMN = VISIBLE_HEADERS.index("Mail ID") + 1
WEEKDAY_START_COLUMN = VISIBLE_HEADERS.index("Mon") + 1
WEEKDAY_COLUMNS = {
    0: WEEKDAY_START_COLUMN,
    1: WEEKDAY_START_COLUMN + 1,
    2: WEEKDAY_START_COLUMN + 2,
    3: WEEKDAY_START_COLUMN + 3,
    4: WEEKDAY_START_COLUMN + 4,
}


def current_week_range():
    today = date.today()
    start = today - timedelta(days=today.weekday())
    end = start + timedelta(days=6)
    return start, end


def week_range_for(selected_date):
    start = selected_date - timedelta(days=selected_date.weekday())
    end = start + timedelta(days=6)
    return start, end


def get_week_label(week_start):
    return f"WK{week_start.isocalendar().week}"


def style_worksheet(sheet, week_start=None):
    title_fill = PatternFill("solid", fgColor="0070C0")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    metadata_fill = PatternFill("solid", fgColor="F2F4F7")
    thin_side = Side(style="thin", color="808080")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    visible_end_column = get_column_letter(len(VISIBLE_HEADERS))

    sheet.title = "WFH Plan"
    sheet.freeze_panes = "A4"
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
    sheet.merge_cells(f"A1:{visible_end_column}1")
    sheet.merge_cells(f"A2:{visible_end_column}2")

    sheet["A1"] = WORKBOOK_TITLE
    sheet["A1"].fill = title_fill
    sheet["A1"].font = Font(color="FFFFFF", bold=True, size=14)
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

    if week_start:
        sheet["A2"] = get_week_label(week_start)
    elif not sheet["A2"].value:
        sheet["A2"] = get_week_label(current_week_range()[0])

    sheet["A2"].font = Font(bold=True)
    sheet["A2"].alignment = Alignment(horizontal="center", vertical="center")

    for column, header in enumerate(SHEET_HEADERS, start=1):
        cell = sheet.cell(row=3, column=column, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        cell.fill = header_fill if column <= len(VISIBLE_HEADERS) else metadata_fill

    for column in range(1, len(SHEET_HEADERS) + 1):
        column_letter = get_column_letter(column)
        sheet.column_dimensions[column_letter].hidden = column > len(VISIBLE_HEADERS)

        if column == NAME_COLUMN:
            sheet.column_dimensions[column_letter].width = 18
        elif column == EMAIL_COLUMN:
            sheet.column_dimensions[column_letter].width = 34
        elif column == EMP_ID_COLUMN:
            sheet.column_dimensions[column_letter].width = 14
        elif column <= len(VISIBLE_HEADERS):
            sheet.column_dimensions[column_letter].width = 12
        else:
            sheet.column_dimensions[column_letter].width = 16


def style_data_row(sheet, row_number):
    thin_side = Side(style="thin", color="808080")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for column in range(1, len(SHEET_HEADERS) + 1):
        cell = sheet.cell(row=row_number, column=column)
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sheet.cell(row=row_number, column=NAME_COLUMN).alignment = Alignment(horizontal="left", vertical="center")
    sheet.cell(row=row_number, column=EMAIL_COLUMN).alignment = Alignment(horizontal="left", vertical="center")


def create_workbook():
    workbook = Workbook()
    sheet = workbook.active
    style_worksheet(sheet, current_week_range()[0])
    workbook.save(DATA_FILE)
    return workbook


def is_plan_format(sheet):
    visible_headers = [sheet.cell(row=3, column=column).value for column in range(1, len(VISIBLE_HEADERS) + 1)]
    return sheet["A1"].value == WORKBOOK_TITLE and visible_headers == VISIBLE_HEADERS


def is_legacy_plan_format(sheet):
    return sheet["A1"].value == WORKBOOK_TITLE and sheet["A3"].value == "Name"


def migrate_old_workbook(workbook):
    old_sheet = workbook.active
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    style_worksheet(new_sheet, current_week_range()[0])

    if is_legacy_plan_format(old_sheet):
        rows = list(old_sheet.iter_rows(min_row=4, values_only=True))
        for row in rows:
            if not row or not any(row):
                continue

            name = row[0]
            email = row[1] if len(row) > 1 else ""
            emp_id = row[8] if len(row) > 8 else ""
            wfh = row[9] if len(row) > 9 else ""
            wfh_date = row[10] if len(row) > 10 else ""
            week_start = row[11] if len(row) > 11 else ""
            week_end = row[12] if len(row) > 12 else ""
            status = row[13] if len(row) > 13 else ""
            submitted_at = row[7] if len(row) > 7 else ""

            selected_date = None
            if wfh_date:
                selected_date = datetime.strptime(str(wfh_date), "%Y-%m-%d").date()

            append_plan_row(
                new_sheet,
                name,
                email,
                emp_id,
                wfh,
                selected_date,
                week_start,
                week_end,
                status,
                submitted_at,
            )

        new_workbook.save(DATA_FILE)
        return new_workbook

    rows = list(old_sheet.iter_rows(min_row=2, values_only=True))
    for row in rows:
        if len(row) < 9:
            continue

        submitted_at, name, email, emp_id, wfh, wfh_date, week_start, week_end, status = row[:9]
        selected_date = None
        if wfh_date:
            selected_date = datetime.strptime(str(wfh_date), "%Y-%m-%d").date()

        append_plan_row(
            new_sheet,
            name,
            email,
            emp_id,
            wfh,
            selected_date,
            week_start,
            week_end,
            status,
            submitted_at,
        )

    new_workbook.save(DATA_FILE)
    return new_workbook


def get_workbook():
    if not DATA_FILE.exists():
        return create_workbook()

    workbook = load_workbook(DATA_FILE)
    if is_plan_format(workbook.active):
        style_worksheet(workbook.active)
        return workbook

    return migrate_old_workbook(workbook)


def count_approved_wfh_requests(sheet, week_start, week_end):
    count = 0
    week_start_text = week_start.isoformat()
    week_end_text = week_end.isoformat()

    for row in sheet.iter_rows(min_row=4, values_only=True):
        work_from_home = row[WORK_FROM_HOME_COLUMN - 1]
        row_week_start = row[WEEK_START_COLUMN - 1]
        row_week_end = row[WEEK_END_COLUMN - 1]
        status = row[STATUS_COLUMN - 1]

        if (
            work_from_home == "Yes"
            and row_week_start == week_start_text
            and row_week_end == week_end_text
            and status == "Approved"
        ):
            count += 1

    return count


def get_weekday_value(wfh_date, status):
    if not wfh_date or status != "Approved":
        return ""

    return "WFH"


def append_plan_row(
    sheet,
    name,
    email,
    emp_id,
    wfh,
    wfh_date,
    week_start,
    week_end,
    status,
    submitted_at=None,
):
    row_number = sheet.max_row + 1
    sheet.cell(row=row_number, column=NAME_COLUMN, value=name)
    sheet.cell(row=row_number, column=EMP_ID_COLUMN, value=emp_id)
    sheet.cell(row=row_number, column=EMAIL_COLUMN, value=email)

    if wfh_date and wfh_date.weekday() in WEEKDAY_COLUMNS:
        sheet.cell(
            row=row_number,
            column=WEEKDAY_COLUMNS[wfh_date.weekday()],
            value=get_weekday_value(wfh_date, status),
        )

    sheet.cell(row=row_number, column=METADATA_START_COLUMN, value=submitted_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    sheet.cell(row=row_number, column=WORK_FROM_HOME_COLUMN, value=wfh)
    sheet.cell(row=row_number, column=METADATA_START_COLUMN + METADATA_HEADERS.index("WFH Date"), value=wfh_date.isoformat() if wfh_date else "")
    sheet.cell(row=row_number, column=WEEK_START_COLUMN, value=week_start.isoformat() if hasattr(week_start, "isoformat") else week_start or "")
    sheet.cell(row=row_number, column=WEEK_END_COLUMN, value=week_end.isoformat() if hasattr(week_end, "isoformat") else week_end or "")
    sheet.cell(row=row_number, column=STATUS_COLUMN, value=status)
    style_data_row(sheet, row_number)


def save_request(name, email, emp_id, wfh, wfh_date, week_start, week_end, status):
    workbook = get_workbook()
    sheet = workbook.active
    style_worksheet(sheet, week_start)
    append_plan_row(sheet, name, email, emp_id, wfh, wfh_date, week_start, week_end, status)
    workbook.save(DATA_FILE)


def get_wfh_plan_table():
    workbook = get_workbook()
    sheet = workbook.active
    headers = [sheet.cell(row=3, column=column).value for column in range(1, len(VISIBLE_HEADERS) + 1)]
    rows = []

    for row in sheet.iter_rows(min_row=4, max_col=len(VISIBLE_HEADERS), values_only=True):
        row_values = [value or "" for value in row]
        if any(row_values):
            rows.append(row_values)

    return {
        "title": sheet["A1"].value or WORKBOOK_TITLE,
        "week": sheet["A2"].value or get_week_label(current_week_range()[0]),
        "headers": headers,
        "rows": rows,
    }


def fill_empty_weekdays_with_wfo():
    workbook = get_workbook()
    sheet = workbook.active

    for row_number in range(4, sheet.max_row + 1):
        name = sheet.cell(row=row_number, column=NAME_COLUMN).value
        email = sheet.cell(row=row_number, column=EMAIL_COLUMN).value

        if not name and not email:
            continue

        for column in WEEKDAY_COLUMNS.values():
            cell = sheet.cell(row=row_number, column=column)
            if not cell.value:
                cell.value = "WFO"

        style_data_row(sheet, row_number)

    workbook.save(DATA_FILE)


def delete_wfh_excel_file():
    if DATA_FILE.exists():
        DATA_FILE.unlink()


def get_wfh_plan_emails():
    table_data = get_wfh_plan_table()
    emails = []

    for row in table_data["rows"]:
        email = row[EMAIL_COLUMN - 1] if len(row) >= EMAIL_COLUMN else ""
        if "@" in email:
            emails.append(email)

    return sorted(set(emails))


def submit_wfh_request(name, email, emp_id, wfh, wfh_date_text):
    if not name or not email or not emp_id or wfh not in {"Yes", "No"}:
        return RequestResult("Please fill all required fields.", "error")

    if wfh == "No":
        save_request(name, email, emp_id, wfh, None, None, None, "Recorded")
        return RequestResult("Your response has been saved.")

    try:
        selected_date = datetime.strptime(wfh_date_text, "%Y-%m-%d").date()
    except ValueError:
        return RequestResult("Please select a valid work from home date.", "error")

    this_week_start, this_week_end = current_week_range()
    if not this_week_start <= selected_date <= this_week_end:
        return RequestResult("Please select a date from this week only.", "error")

    selected_week_start, selected_week_end = week_range_for(selected_date)
    workbook = get_workbook()
    sheet = workbook.active
    approved_count = count_approved_wfh_requests(
        sheet,
        selected_week_start,
        selected_week_end,
    )

    if approved_count >= WEEKLY_WFH_LIMIT:
        save_request(
            name,
            email,
            emp_id,
            wfh,
            selected_date,
            selected_week_start,
            selected_week_end,
            "Rejected - 4 members reached",
        )
        return RequestResult(
            "4 members reached. Work from home is not available for this week.",
            "error",
        )

    save_request(
        name,
        email,
        emp_id,
        wfh,
        selected_date,
        selected_week_start,
        selected_week_end,
        "Approved",
    )
    return RequestResult("Your work from home request has been saved.")
