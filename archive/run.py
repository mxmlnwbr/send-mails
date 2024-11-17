# libraries to be imported
from email.mime.image import MIMEImage
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os
import pandas as pd
from tqdm import tqdm

attach_files = True
add_image_to_body = False

load_dotenv()

smtp_server = "smtp.gmail.com"
port = 587  # For starttls

# df = pd.read_excel("/Users/mw/Documents/send-mails/input.xlsx", sheet_name=0)
# email_adresses = df["E-Mail"].tolist()
email_adresses = ["maximilian.weber@bluewin.ch", "mxjweber@gmail.com"]

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")

# Python code to illustrate Sending mail with attachments
# from your Gmail account
counter = 0
received_mail_adresses = []
for toaddr in tqdm(email_adresses):
    if toaddr in received_mail_adresses:
        continue
    counter += 1

    tqdm.write(f"Number {counter}: {toaddr}")

    subject = "RigiBeats 2024 - RigiBahn Ticket"
    body = """
    <html>
        <body>
            <p>👋 Hallo lieber Gast!,</p>
            <p>wir können es kaum erwarten, dich am 20.01.2024 auf der Rigi zu begrüssen und gemeinsam ein unvergessliches Event zu erleben! 🎉 </p>
            <p>In Anbetracht des grossen Interesses und des Ansturms auf die Tickets, möchten wir sicherstellen, dass jeder Gast die Möglichkeit hat, seine Tickets entsprechend anzupassen. </p>
            <p>Wenn du mehr Tickets gekauft hast, als du benötigst, kannst du diese bis zum 18.01.2024 um 18:00 Uhr gratis zurückgeben. Wir werden diese erneut in Umlauf bringen. </p>
            <p>Der Prozess dafür ist ganz einfach – schicke uns eine E-Mail mit den betroffenen Ticket-IDs an rigibeats.ch@gmail.com. Unser Team wird sich umgehend darum kümmern und die Rückerstattung für die überzähligen Tickets veranlassen. </p>
            <p>Wir bedanken uns für dein Verständnis und deine Mitarbeit in dieser Angelegenheit. Unser Ziel ist es, sicherzustellen, dass jeder Gast die bestmögliche Erfahrung bei unserem Event hat. </p>
            <p>Wir freuen uns schon riesig darauf, dich bald auf der Rigi zu sehen und gemeinsam eine grossartige Zeit zu haben! </p>
            <p>Dein RigiBeats Team ❤️</p>
            <br>
            <img src="cid:image1">
        </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = FROM
    msg["To"] = toaddr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    # ADD IMAGE TO BODY
    if add_image_to_body:
        with open("/Users/mw/Documents/send-mails/img.jpeg", "rb") as fp:
            img = MIMEImage(fp.read())
            img.add_header("Content-ID", "<image1>")
            msg.attach(img)

    # ATTACH FILE TO EMAIL
    if attach_files:
        # open the file to be sent
        filename = f"dummy{counter}.pdf"
        attachment = open(
            f"/Users/mw/Documents/send-mails/files/dummy{counter}.pdf",
            "rb",
        )

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
    received_mail_adresses.append(toaddr)
