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

attach_files = False
add_image_to_body = False

load_dotenv()

smtp_server = "smtp.gmail.com"
port = 587  # For starttls

df = pd.read_excel("/Users/mw/Documents/send-mails/survey_input.xlsx", sheet_name=0)
email_adresses = df["E-Mail"].tolist()
status = df["Status"].tolist()
email_adresses = [
    "maximilian.weber@bluewin.ch",
    "mxjweber@gmail.com",
    "kuettel_m@bluewin.ch",
]

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")

# Python code to illustrate Sending mail with attachments
# from your Gmail account
counter = 0
for toaddr in tqdm(email_adresses):
    counter += 1

    tqdm.write(f"Number {counter}: {toaddr}")

    subject = "RigiBeats 2024 - Feedback Umfrage"
    body = """
    <html>
        <body>
            <p>Liebe Rigibeats Teilnehmende</p>
            <p>Vielen Dank für deine Teilnahme an unserem Event. Wir hoffen, dass du eine angenehme Zeit hattest und dass Rigibeats deine Erwartungen erfüllen konnte.<br>
            Um sicherzustellen, dass wir auch zukünftig erstklassige Veranstaltungen anbieten, sind wir auf deine Rückmeldung angewiesen. Wir möchten dich daher bitten, an unserer kurzen Umfrage teilzunehmen, die uns helfen wird, Rigibeats zu verbessern.<br>
            Die Umfrage wird etwa 5 Minuten in Anspruch nehmen und ist anonym. Wir schätzen deine ehrlichen Antworten sehr.</p>
            <p>Bitte klicke auf den folgenden Link, um zur Umfrage zu gelangen: https://forms.gle/niP2LkXHEyYJb3vs6</p>
            <p><b>Unter allen Teilnehmern verlosen wir 1x 2 Rigibeats 2025 Tickets! 🤩</b></p>
            <p>Wenn du weitere Fragen oder Anregungen hast, kannst du uns gerne per E-Mail oder Instagram kontaktieren.</p>
            <p>Liebe Grüsse<br>
            Euer RigiBeats Team ❤️</p>
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
    df["Status"][counter - 1] = "sent"
    df.to_excel("/Users/mw/Documents/send-mails/survey_output.xlsx", index=False)
