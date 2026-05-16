# Mail App

Flask application for collecting weekly work-from-home requests, storing them in an Excel workbook, viewing the workbook data in the UI, and sending a WFH summary email to all employee mail IDs.

## Features

- Work-from-home request form
- Employee fields: name, employee ID, email
- Weekly WFH date selection
- Maximum 4 approved WFH requests per week
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
└── wfh_requests.xlsx
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

Request form:

```text
http://localhost:5000/
```

WFH table view:

```text
http://localhost:5000/wfh-plan
```

## Excel File

The app creates `wfh_requests.xlsx` automatically.

Visible Excel columns:

```text
Name | Emp ID | Mail ID | Mon | Tue | Wed | Thu | Fri
```

Hidden columns are used internally for submitted time, WFH date, week start, week end, and status.

## Mail Flow

On the WFH table page, click `Send Mail`.

The app will:

1. Fill empty weekday cells with `WFO`.
2. Build an HTML summary email.
3. Send the email to all mail IDs in the Excel file.
4. Delete `wfh_requests.xlsx` after successful mail send.
5. Redirect to the form page with success messages.

If mail sending fails, the Excel file is not deleted.
