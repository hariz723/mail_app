from html import escape
import smtplib
from email.message import EmailMessage

from env_loader import SMTPConfig


def send_mail(to_email, subject, body, html_body=None):
    SMTPConfig.validate()
    recipients = to_email if isinstance(to_email, list) else [to_email]

    message = EmailMessage()
    message["From"] = SMTPConfig.smtp_from_email
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    message.set_content(body)

    if html_body:
        message.add_alternative(html_body, subtype="html")

    with smtplib.SMTP(SMTPConfig.smtp_host, SMTPConfig.smtp_port) as server:
        server.starttls()
        server.login(SMTPConfig.smtp_username, SMTPConfig.smtp_password)
        server.send_message(message)


def build_wfh_summary_html(table_data):
    header_cells = "".join(f"<th>{escape(str(header))}</th>" for header in table_data["headers"])
    body_rows = []
    weekday_headers = {"Mon", "Tue", "Wed", "Thu", "Fri"}

    for index, row in enumerate(table_data["rows"], start=1):
        cells = [f"<td>{index}</td>"]
        for column_index, value in enumerate(row):
            header = table_data["headers"][column_index]
            display_value = value or ""
            css_class = ""

            if header in weekday_headers:
                display_value = display_value or "WFO"
                css_class = "wfh-day" if display_value == "WFH" else ""

            cells.append(f"<td class=\"{css_class}\">{escape(str(display_value))}</td>")

        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    return f"""
    <!doctype html>
    <html>
    <body style="font-family: Arial, sans-serif; color: #222;">
        <p>Hello Team,</p>
        <p>Kindly find the below <strong>WFH plan</strong> for our team members for <strong>{escape(str(table_data["week"]))}</strong>.</p>

        <table cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size: 13px; min-width: 640px;">
            <thead>
                <tr>
                    <th colspan="{len(table_data["headers"]) + 1}" style="background: #5b9bd5; color: #0b2239; border: 1px solid #444; padding: 10px; text-align: left;">
                        {escape(str(table_data["title"]))}
                    </th>
                </tr>
                <tr>
                    <th style="background: #d9eaf7; border: 1px solid #444; padding: 8px;">#</th>
                    {header_cells}
                </tr>
            </thead>
            <tbody>
                {''.join(body_rows)}
            </tbody>
        </table>

        <p>Regards,<br>Design Room</p>

        <style>
            th, td {{
                border: 1px solid #444;
                padding: 8px 10px;
                text-align: center;
            }}

            th {{
                background: #d9eaf7;
                font-weight: 700;
            }}

            td:nth-child(2),
            td:nth-child(3),
            td:nth-child(4) {{
                text-align: left;
            }}

            .wfh-day {{
                background: #fff200;
                font-weight: 700;
            }}
        </style>
    </body>
    </html>
    """


def send_wfh_summary_mail(recipients, table_data):
    subject = f"{table_data['title']} - {table_data['week']}"
    body = (
        "Hello Team,\n\n"
        f"Kindly find the below WFH plan for our team members for {table_data['week']}.\n\n"
        "Regards,\nDesign Room"
    )
    html_body = build_wfh_summary_html(table_data)
    send_mail(recipients, subject, body, html_body)


if __name__ == "__main__":
    receiver_email = input("To email: ").strip()
    mail_subject = input("Subject: ").strip()
    mail_body = input("Message: ").strip()

    send_mail(receiver_email, mail_subject, mail_body)
    print("Mail sent successfully.")
