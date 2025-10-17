from email.mime.application import MIMEApplication  # Add this if not already imported
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from datetime import datetime, timedelta
import vars

EMAIL_HOST = "imap.gmail.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_USER = "niko.business08@gmail.com"
EMAIL_PASSWORD = "EMAIL_PASS"
REPORT_RECIPIENTS = ["ntsakiridis@whimsytrucking.com"]

date_today = (datetime.now() - timedelta(days=vars.days_back)).strftime("%m-%d-%Y")  # date of today in MM/DD/YYYY format

def send_email():
    filename = os.path.join(os.path.expanduser("~/Downloads/Hapag_XTs"), "Hapag XTs " + date_today + ".xlsx")  # Path to the file to attach
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = ', '.join(REPORT_RECIPIENTS)
        msg['Subject'] = "Hapag XTs " + str(date_today)

        # Email body
        body = f"""Hello, team.
        
Here are all the Hapag Crosstowns from {date_today} up to around 4:00PM.

Best regards,

~ Niko

"""

        msg.attach(MIMEText(body, 'plain'))
        
        if filename and os.path.exists(filename):
            with open(filename, "rb") as attachment:
                part = MIMEApplication(attachment.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(filename)}"'
                )
                msg.attach(part)
        
        # Send email
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_USER, REPORT_RECIPIENTS, text)
        server.quit()
        
        print(f"Email sent successfully to {', '.join(REPORT_RECIPIENTS)}")
        
    except Exception as e:
        print(f"Error sending email: {str(e)}")
