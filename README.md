# Mail App

Flask application for selecting a weekly work-from-home roster, storing it in an Excel workbook, viewing the workbook data in the UI, and sending a WFH summary email to the configured employee mail IDs.

## Features

- Work-from-home roster table
- Employee columns: name and TID
- User-selected 7-day date range with weekday date columns and checkboxes
- Maximum 4 WFH selections per day
- Excel output in WFH plan format
- Table view for Excel data
- Send summary mail to all email IDs in the Excel file
- Fills empty weekday cells as `WFO` before sending mail
- Deletes the Excel file after mail is sent successfully

## Project Structure

```text
mail_app/
├── app.py
├── business.py
├── constants.py
├── env_loader.py
├── send_mail.py
├── requirements.txt
├── .env.example
├── templates/
│   ├── index.html
│   └── wfh_plan.html
└── 2026-06-08_to_2026-06-14.xlsx
```

## Setup

Create and activate a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Create your local `.env` file:

```bash
cp .env.example .env
```

Update `.env` with your SMTP details:

```bash
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USERNAME=your_email@example.com
SMTP_PASSWORD=your_app_password
SMTP_FROM_EMAIL=your_email@example.com
```

For Gmail, use an app password instead of your normal account password.

## Run

Start the Flask app:

```bash
python app.py
```

Open the app:

```text
http://localhost:5000
```

## Pages

WFH roster table:

```text
http://localhost:5000/
```

Select the start date and end date before saving the WFH plan. The range must be 7 days, such as `2026-06-08` to `2026-06-14`.

WFH table view:

```text
http://localhost:5000/wfh-plan
```

## Excel File

When you click `Save WFH Plan`, the app creates an Excel file named with the selected week range.

Example:

```text
2026-06-08_to_2026-06-14.xlsx
```

Visible Excel columns:

```text
Name | TID | Mon <date> | Tue <date> | Wed <date> | Thu <date> | Fri <date>
```

Hidden columns are used internally for submitted time, mail ID, week start, week end, and status. Add employee email IDs in `constants.py` to keep the summary mail recipient list populated.

## Mail Flow

On the WFH table page, click `Send Mail`.

The app will:

1. Fill empty weekday cells with `WFO`.
2. Build an HTML summary email.
3. Send the email to all mail IDs in the Excel file.
4. Delete the selected week range Excel file after successful mail send.
5. Redirect to the form page with success messages.

If mail sending fails, the Excel file is not deleted.
