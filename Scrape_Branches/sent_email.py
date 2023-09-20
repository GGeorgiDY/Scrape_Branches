from email.mime.text import MIMEText
from check_email import get_valid_email
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib

EMAIL_BODY = """
    Hello,

    Attached you can find the Fibank Branches working on the weekends

    Best Regards
"""


def sent_email(excel_file_path):
    gmail_user = os.environ.get('GMAIL_USER')
    gmail_password = os.environ.get('GMAIL_APP_PASSWORD')
    recipient_email = get_valid_email()

    # Creating a connection to the SMTP server
    smtp_server = smtplib.SMTP("smtp.gmail.com", 587)
    smtp_server.starttls()

    # Login to your Gmail account
    smtp_server.login(gmail_user, gmail_password)

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = recipient_email
    msg['Subject'] = "Fibank Branches Data"
    email_body = EMAIL_BODY
    msg.attach(MIMEText(email_body, 'plain'))

    # Attach the Excel file to the email
    attachment = open(excel_file_path, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={excel_file_path}')
    msg.attach(part)

    # Send the email
    smtp_server.sendmail(gmail_user, recipient_email, msg.as_string())

    # Close the SMTP server connection
    smtp_server.quit()

    print("Data has been successfully extracted, saved to fibank_branches.xlsx, and sent via email.")