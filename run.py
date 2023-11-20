# libraries to be imported
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

FROM = os.getenv("GMAIL_ACCOUNT")
PASSWORD = os.getenv("GMAIL_PASSWORD")
subject = "Ticketumtausch RigiBeats 2024"
body = "👋 Hallo lieber Gast! 🤩\n\nWir können es kaum erwarten, mit dir am 20.01.2024 auf der Rigi zu feiern! 🎉\nDer Ticketansturm war enorm, und die Eventfrog-Seite hatte Schwierigkeiten, mit dem Andrang mitzuhalten. Einige haben uns bereits mitgeteilt, dass sie nicht die gewünschte Ticketkategorie kaufen konnten und deshalb sich ein anderes Ticket sicherten.\n\nDa Fairness für uns oberste Priorität hat, bieten wir dir hier die Möglichkeit an, falsch gekaufte Tickets gegen die richtigen umzutauschen. Falls dich dies betrifft, findest du mehr Infos unter folgendem Link: https://docs.google.com/forms/d/e/1FAIpQLSfze4KtZJLOQMHrzRTw5q5uwcWXNoZBz8NeoM_kYIeWnDM0UQ/viewform?usp=sf_link\n\n**WICHTIG: AUSFÜLLEN BIS ZUM 26.11.2023 UM 20:00 UHR, ANSONSTEN KÖNNEN WIR DEINE ÄNDERUNGEN NICHT MEHR BERÜCKSICHTIGEN!**\n\nDanke für dein Verständnis & Bis Bald auf der Rigi 🏔️👑\nDein RigiBeats Team ❤️"

df = pd.read_excel(
    "/Users/mw/Documents/send-mails/eventfrog-ticketexport.xlsx", sheet_name="Tickets"
)
email_adresses = df["E-Mail"].tolist()

# Python code to illustrate Sending mail with attachments
# from your Gmail account
counter = 0
for toaddr in tqdm(email_adresses):
    counter += 1

    tqdm.write(f"Number {counter}: {toaddr}")

    msg = MIMEMultipart()
    msg["From"] = FROM
    msg["To"] = toaddr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # # open the file to be sent
    # filename = "File_name_with_extension.pdf"
    # attachment = open("/Users/mw/Downloads/dummy.pdf", "rb")

    # instance of MIMEBase and named as p
    # p = MIMEBase("application", "octet-stream")

    # To change the payload into encoded form
    # p.set_payload((attachment).read())

    # encode into base64
    # encoders.encode_base64(p)

    # p.add_header("Content-Disposition", "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    # msg.attach(p)

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
