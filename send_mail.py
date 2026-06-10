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
    weekday_prefixes = ("Mon", "Tue", "Wed", "Thu", "Fri")
    page_bg = "#f3f6f8"
    accent = "#0f766e"
    accent_dark = "#115e59"
    border = "#d8dee7"
    muted = "#667085"

    header_html = "".join(
        f'<th bgcolor="#d9eaf7" style="padding: 10px 8px; border: 1px solid {border}; '
        f'text-align: center; font-family: Arial, Helvetica, sans-serif; font-size: 12px; '
        f'font-weight: bold; color: #172033;">{escape(str(header))}</th>'
        for header in table_data["headers"]
    )

    rows_html = []
    for index, row in enumerate(table_data["rows"]):
        row_bg = "#ffffff" if index % 2 == 0 else "#f8fafc"
        cells = [
            f'<td bgcolor="{row_bg}" style="padding: 10px 8px; border: 1px solid {border}; '
            f'text-align: center; font-family: Arial, Helvetica, sans-serif; color: #667085; '
            f'font-size: 13px;">{index + 1}</td>'
        ]

        for col_idx, value in enumerate(row):
            header = table_data["headers"][col_idx]
            display_value = value or ""
            cell_bg = row_bg
            cell_color = "#151922"
            font_weight = "normal"

            if col_idx < 2:
                alignment = "left"
            else:
                alignment = "center"

            if str(header).startswith(weekday_prefixes):
                display_value = display_value or "WFO"
                if display_value == "WFH":
                    cell_bg = "#e6f5f2"
                    cell_color = accent_dark
                    font_weight = "bold"
                else:
                    cell_color = muted

            cells.append(
                f'<td bgcolor="{cell_bg}" style="padding: 10px 8px; border: 1px solid {border}; '
                f'text-align: {alignment}; font-family: Arial, Helvetica, sans-serif; '
                f'font-size: 14px; color: {cell_color}; font-weight: {font_weight};">'
                f'{escape(str(display_value))}</td>'
            )

        rows_html.append(f"<tr>{''.join(cells)}</tr>")

    return f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="x-ua-compatible" content="ie=edge">
    </head>
    <body bgcolor="{page_bg}" style="margin: 0; padding: 0; background-color: {page_bg};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{page_bg}" style="background-color: {page_bg};">
            <tr>
                <td align="center" style="padding: 24px 12px;">
                    <table role="presentation" width="760" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="width: 760px; max-width: 760px; background-color: #ffffff; border-collapse: collapse; border: 1px solid {border};">
                        <tr>
                            <td bgcolor="{accent}" style="padding: 24px; background-color: {accent}; font-family: Arial, Helvetica, sans-serif; color: #ffffff;">
                                <h1 style="margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; font-size: 24px; line-height: 30px; font-weight: bold; color: #ffffff;">
                                    {escape(str(table_data["title"]))}
                                </h1>
                                <p style="margin: 8px 0 0 0; padding: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 22px; color: #d9fffb;">
                                    Weekly attendance and work plan
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 22px 24px 8px 24px; font-family: Arial, Helvetica, sans-serif;">
                                <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                                    <tr>
                                        <td bgcolor="#e6f5f2" style="padding: 7px 12px; border: 1px solid #cde7e2; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; color: {accent_dark}; font-weight: bold;">
                                            {escape(str(table_data["week"]))}
                                        </td>
                                    </tr>
                                </table>
                                <p style="margin: 18px 0 18px 0; padding: 0; font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 22px; color: #475569;">
                                    Hello Team,<br><br>
                                    Please find the updated Work From Home (WFH) schedule for the upcoming week.
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 0 24px 22px 24px;">
                                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width: 100%; border-collapse: collapse;">
                                    <thead>
                                        <tr>
                                            <th bgcolor="#d9eaf7" style="padding: 10px 8px; border: 1px solid {border}; text-align: center; font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #172033;">#</th>
                                            {header_html}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {"".join(rows_html)}
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 0 24px 24px 24px; font-family: Arial, Helvetica, sans-serif;">
                                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; border-top: 1px solid {border};">
                                    <tr>
                                        <td style="padding-top: 18px; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 20px; color: {muted};">
                                            Best regards,<br>
                                            <strong style="color: #151922;">Design Room Team</strong>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="{accent_dark}" height="4" style="height: 4px; line-height: 4px; font-size: 0; background-color: {accent_dark};">&nbsp;</td>
                        </tr>
                    </table>
                    <table role="presentation" width="760" cellpadding="0" cellspacing="0" border="0" style="width: 760px; max-width: 760px;">
                        <tr>
                            <td align="center" style="padding: 14px 8px 0 8px; font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 18px; color: #94a3b8;">
                                This is an automated notification. Please do not reply directly to this email.
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
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
