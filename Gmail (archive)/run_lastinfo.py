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

df = pd.read_excel("/Users/mw/Documents/send-mails/HelferCodes.xlsx", sheet_name=0)
email_adresses = df["E-Mail"].tolist()
codes = df["Code"].tolist()
status = df["Status"].tolist()
# email_adresses = ["maximilian.weber@bluewin.ch", "mxjweber@gmail.com"]

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")

# Python code to illustrate Sending mail with attachments
# from your Gmail account
counter = 0
for toaddr in tqdm(email_adresses):
    counter += 1
    code = codes[counter - 1]

    tqdm.write(f"Number {counter}: {toaddr}")

    subject = "RigiBeats 2024 - Helfer:innen Tickets"
    body = f"""
    <html>
        <body>
            <p>Geschätztes Partyvolk,</p>
            <p>der Countdown nähert sich dem Ende – schon diesen Morgen steigt das RigiBeats 2024!</p>
            <p>Wir freuen uns darauf, mit euch das atemberaubende Bergpanorama zu geniessen und gemeinsam die kalte Winterluft zu vertreiben.</p>
            <p>Damit du optimal auf den Event vorbereitet bist, hier noch ein paar Informationen:</p>
            <p>- Packe Sonnenbrille ein und ziehe warme Winterausrüstung an, denn gutes, kaltes Wetter erwartet uns. So kannst du das RigiBeats bis zum Schluss in vollen Zügen geniessen.</p>
            <p>-  Wir haben Extrafahrten für euch organisiert, damit alle sicher wieder nach Hause kommen. Die Informationen zur An- und Abreise findest du auf Instagram (@rigibeats)</p>
            <p>- “En Bode zha” ist empfehlenswert, denn der RigiBeats Kebabstand erst ab 17:00 Uhr geöffnet. Wer auf dem Berg Mittagessen möchte, kann dies im Gasthaus Rigi Scheidegg oder Rigi Burggeist (bis 14:00 Uhr).</p>
            <p>So das wars. Wenn du weitere Informationen möchtest, findest du diese auf unserer Webseite oder Instagram (@rigibeats).</p>
            <p>Bis bald, auf der Tanzfläche, am Nagelstock, beim Beerpong oder an der Bar....</p>
            <p>Euer RigiBeats Team ❤️</p>
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
    df.to_excel("/Users/mw/Documents/send-mails/HelferCodes_output.xlsx", index=False)
