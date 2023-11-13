# libraries to be imported
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

load_dotenv()

smtp_server = "smtp.gmail.com"
port = 587  # For starttls

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")
email_adresses = [
    "maximilian.weber@bluewin.ch",
    "maximilian.weber@bf.uzh.ch",
    "mxjweber@gmail.com",
]
subject = "Test"
body = "Test"

# Python code to illustrate Sending mail with attachments
# from your Gmail account
for toaddr in email_adresses:
    print(toaddr)

    msg = MIMEMultipart()
    msg["From"] = FROM
    msg["To"] = toaddr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # open the file to be sent
    filename = "File_name_with_extension.pdf"
    attachment = open("/Users/mw/Downloads/dummy.pdf", "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase("application", "octet-stream")

    # To change the payload into encoded form
    p.set_payload((attachment).read())

    # encode into base64
    encoders.encode_base64(p)

    p.add_header("Content-Disposition", "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # creates SMTP session
    s = smtplib.SMTP("smtp.gmail.com", 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(FROM, PASSWORD)

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(FROM, toaddr, text)

    # terminating the session
    s.quit()
