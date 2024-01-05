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

load_dotenv()

smtp_server = "smtp.gmail.com"
port = 587  # For starttls

df = pd.read_excel("/Users/mw/Documents/send-mails/input.xlsx", sheet_name=0)
email_adresses = df["E-Mail:"].tolist()

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")

# Python code to illustrate Sending mail with attachments
# from your Gmail account
counter = 0
for toaddr in tqdm(email_adresses):
    counter += 1

    tqdm.write(f"Number {counter}: {toaddr}")

    vorname = df["Vorname:"].tolist()[counter - 1].capitalize()
    if df["Geschlecht:"].tolist()[counter - 1] == "Weiblich":
        Anrede = "Liebe"
    else:
        Anrede = "Lieber"

    subject = "Weihnachtsgrüsse von der Flusswelle Muota!"
    body = f"""
    <html>
        <body>
            <p>{Anrede} {vorname},</p>
            <p>wir hoffen, dass diese Nachricht dich in bester Gesundheit erreicht und du eine schöne Vorweihnachtszeit hast. Wir wünschen frohe Weihnachten und ein strahlendes neues Jahr! Möge diese Zeit des Jahres mit Freude, Wärme und unvergesslichen Momenten mit deinen Liebsten gefüllt sein.</p>
            <p>Es gibt Updates bezüglich der Flusswelle: Der geplante Abriss der Muota Flusswelle wird nach wie vor vorangetrieben (EBS & Bezirk Schwyz). Allerdings konnten wir dank der Unterstützung dieses Vereins und euch als geschätzte Mitglieder Argumente vorbringen, die zu einem Gespräch mit den Projektverantwortlichen geführt haben. Der nächste Schritt besteht nun darin, auszuarbeiten, wie eine "neue" Welle im Rahmen des Projekts realisiert werden könnte.</p>
            <p>In diesem Sinne möchten wir euch herzlich danken. Gemeinsam machen wir die Flusswelle Muota zu einem Ort der Begeisterung und des Austauschs für Wassersportliebhaber. Wir freuen uns auf ein aufregendes neues Jahr voller Wellen und spannender Momente!</p>
            <p>Frohe Feiertage und einen fantastischen Start ins neue Jahr!</p>
            <p>P.S. Anbei findest du unsere Stellungnahme zum Vorprojekt des Bezirks Schwyz.</p>
            <p>Weihnachtliche Grüsse 🎅,<br>
            Flusswelle 🌊 Muota</p>
            <br>
            <img src="cid:image1">
        </body>
    </html>
    """
    subject = "Rücknahme zu viel gekaufter Tickets - RigiBeats 2024"
    body = """
    <html>
        <body>
            <p>👋 Hallo lieber Gast!,</p>
            <p>wir können es kaum erwarten, dich am 20.01.2024 auf der Rigi zu begrüssen und gemeinsam ein unvergessliches Event zu erleben! 🎉 </p>
            <p>In Anbetracht des grossen Interesses und des Ansturms auf die Tickets, möchten wir sicherstellen, dass jeder Gast die Möglichkeit hat, seine Tickets entsprechend anzupassen. </p>
            <p>Wenn du mehr Tickets gekauft hast, als du benötigst, kannst du diese bis zum 18.01.2024 um 18:00 Uhr gratis zurückgeben. Wir werden diese erneut in Umlauf bringen. </p>
            <p>Der Prozess dafür ist ganz einfach – schicke uns einfach eine E-Mail mit den betroffenen Ticket-IDs an rigibeats.ch@gmail.com. Unser Team wird sich umgehend darum kümmern und die Rückerstattung für die überzähligen Tickets veranlassen. </p>
            <p>Wir bedanken uns für dein Verständnis und deine Mitarbeit in dieser Angelegenheit. Unser Ziel ist es, sicherzustellen, dass jeder Gast die bestmögliche Erfahrung bei unserem Event hat. </p>
            <p>Wir freuen uns schon riesig darauf, dich bald auf der Rigi zu sehen und gemeinsam eine grossartige Zeit zu haben! </p>
            <p>Dein RigiBeats Team ❤️</p>
        </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = FROM
    msg["To"] = toaddr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))

    with open("/Users/mw/Documents/send-mails/img.jpeg", "rb") as fp:
        img = MIMEImage(fp.read())
        img.add_header("Content-ID", "<image1>")
        msg.attach(img)

    # open the file to be sent
    filename = "Stellungnahme_Vorprojekt_Bezirk_Schwyz.pdf"
    attachment = open(
        "/Users/mw/Documents/send-mails/Stellungnahme_Vorprojekt_Bezirk_Schwyz.pdf",
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
