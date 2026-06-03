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
    
    header_html = "".join([
        f'<th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: center; font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em;">{escape(str(header))}</th>'
        for header in table_data["headers"]
    ])
    
    rows_html = []
    for index, row in enumerate(table_data["rows"]):
        cells = [f'<td style="padding: 12px 8px; border-bottom: 1px solid #e2e8f0; text-align: center; color: #94a3b8; font-size: 13px;">{index + 1}</td>']
        for col_idx, value in enumerate(row):
            header = table_data["headers"][col_idx]
            display_value = value or ""
            
            # Base styles for cells
            cell_style = "padding: 12px 8px; border-bottom: 1px solid #e2e8f0; font-size: 14px; color: #1e293b;"
            
            # Alignment logic
            if col_idx < 2:
                cell_style += " text-align: left;"
            else:
                cell_style += " text-align: center;"

            # WFH Highlighting
            if str(header).startswith(weekday_prefixes):
                display_value = display_value or "WFO"
                if display_value == "WFH":
                    cell_style += " color: #0f766e; font-weight: 600; background-color: #f0fdfa;"
                else:
                    cell_style += " color: #64748b;"
            
            cells.append(f'<td style="{cell_style}">{escape(str(display_value))}</td>')
        
        bg_color = "#ffffff" if index % 2 == 0 else "#f8fafc"
        rows_html.append(f'<tr style="background-color: {bg_color};">{"".join(cells)}</tr>')

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin: 0; padding: 20px 0; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
        <div style="max-width: 800px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); border: 1px solid #e2e8f0;">
            <!-- Header -->
            <div style="padding: 32px 24px; background: linear-gradient(135deg, #0f766e 0%, #115e59 100%); color: #ffffff;">
                <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.025em;">{escape(str(table_data["title"]))}</h1>
                <p style="margin: 8px 0 0 0; font-size: 16px; color: #ccfbf1; opacity: 0.9;">Weekly attendance and work plan</p>
            </div>
            
            <!-- Content -->
            <div style="padding: 24px;">
                <div style="margin-bottom: 24px; display: inline-block; padding: 6px 12px; background-color: #f0fdfa; border: 1px solid #ccfbf1; border-radius: 20px; color: #0f766e; font-size: 14px; font-weight: 600;">
                    {escape(str(table_data["week"]))}
                </div>
                
                <p style="margin: 0 0 20px 0; font-size: 15px; color: #475569; line-height: 1.5;">
                    Hello Team, <br><br>
                    Please find the updated Work From Home (WFH) schedule for the upcoming week.
                </p>
                
                <div style="overflow-x: auto; -webkit-overflow-scrolling: touch;">
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0; min-width: 600px;">
                        <thead>
                            <tr>
                                <th style="padding: 12px 8px; border-bottom: 2px solid #e2e8f0; text-align: center; font-size: 12px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em;">#</th>
                                {header_html}
                            </tr>
                        </thead>
                        <tbody>
                            {"".join(rows_html)}
                        </tbody>
                    </table>
                </div>
                
                <!-- Footer -->
                <div style="margin-top: 32px; padding-top: 24px; border-top: 1px solid #e2e8f0;">
                    <p style="margin: 0; font-size: 14px; color: #64748b;">
                        Best regards,<br>
                        <strong style="color: #1e293b;">Design Room Team</strong>
                    </p>
                </div>
            </div>
            
            <!-- Bottom Accent -->
            <div style="height: 4px; background: linear-gradient(90deg, #0f766e, #14b8a6);"></div>
        </div>
        
        <div style="max-width: 800px; margin: 16px auto; text-align: center;">
            <p style="font-size: 12px; color: #94a3b8;">
                This is an automated notification. Please do not reply directly to this email.
            </p>
        </div>
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
