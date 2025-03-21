import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv
from tqdm import tqdm
import logging
import pandas as pd

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 587  # Using 587 for TLS
EMAIL_ADDRESS = os.getenv("SMTP_ACCOUNT")
EMAIL_PASSWORD = os.getenv("SMTP_PASSWORD")

ATTACHMENTS_FOLDER = "attachments"  # Define the attachment folder
DATA_FOLDER = "data"  # Define the data folder
SENT_COLUMN_NAME = "Sent"  # The column in the Excel file that tracks sent emails


def read_email_addresses(excel_filename):
    """
    Reads email addresses from an Excel file and adds a 'Sent' column if it doesn't exist.

    Args:
        excel_filename (str): The name of the Excel file (located in the 'data' folder).

    Returns:
        tuple: A tuple containing the dataframe and a list of email addresses that have not been sent.
    """
    try:
        file_path = os.path.join(DATA_FOLDER, excel_filename)
        # Read the Excel file using pandas
        df = pd.read_excel(file_path)

        # Check if the 'Sent' column exists; if not, create it and set all values to 'No'
        if SENT_COLUMN_NAME not in df.columns:
            df[SENT_COLUMN_NAME] = "No"
            logging.info(f"Added '{SENT_COLUMN_NAME}' column to the Excel file.")

        # Filter the email addresses that are not sent yet (i.e., 'Sent' == 'No')
        email_addresses = df[df[SENT_COLUMN_NAME] != "Yes"]["E-Mail"].dropna().tolist()
        logging.info(
            f"Successfully read {len(email_addresses)} email addresses that need to be sent."
        )
        return df, email_addresses

    except Exception as e:
        logging.error(f"Error reading email addresses from {excel_filename}: {e}")
        return None, []


def get_all_attachments():
    """
    Retrieves all filenames in the ATTACHMENTS_FOLDER, excluding hidden files (those starting with a dot).

    Returns:
        list: A list of filenames in the ATTACHMENTS_FOLDER.
    """
    try:
        return [
            f
            for f in os.listdir(ATTACHMENTS_FOLDER)
            if os.path.isfile(os.path.join(ATTACHMENTS_FOLDER, f))
            and not f.startswith(".")
        ]
    except Exception as e:
        logging.error(f"Error retrieving attachments: {e}")
        return []


def update_sent_status(df, email):
    """
    Updates the 'Sent' status for an email in the dataframe to 'Yes'.

    Args:
        df (DataFrame): The dataframe containing email addresses.
        email (str): The email address to mark as sent.
    """
    df.loc[df["E-Mail"] == email, SENT_COLUMN_NAME] = "Yes"
    logging.info(f"Marked {email} as sent.")


def save_updated_excel(df, excel_filename):
    """
    Saves the updated dataframe back to the Excel file.

    Args:
        df (DataFrame): The updated dataframe with the 'Sent' status.
        excel_filename (str): The name of the Excel file to save to.
    """
    try:
        file_path = os.path.join(DATA_FOLDER, excel_filename)
        df.to_excel(file_path, index=False)
        logging.info(f"Updated Excel file saved to {file_path}.")
    except Exception as e:
        logging.error(f"Error saving updated Excel file: {e}")


def add_attachments(msg, attachments):
    """
    Add attachments to the email.

    Args:
        msg (MIMEMultipart): The email message object.
        attachments (list): A list of filenames to attach (located in ATTACHMENTS_FOLDER).
    """
    for filename in attachments:
        file_path = os.path.join(ATTACHMENTS_FOLDER, filename)
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)
            logging.info(f"Attached file: {filename}")
        except FileNotFoundError:
            logging.warning(f"Attachment file not found: {filename}")
        except Exception as e:
            logging.error(f"Error attaching file {filename}: {e}")


def send_personalized_email(df, row_index, subject_template, html_body_template, attachments=None):
    """
    Sends a personalized email to a single recipient based on their row data.

    Args:
        df (DataFrame): The dataframe containing email addresses and their data
        row_index (int): The index of the row to process
        subject_template (str): The template for the email subject with placeholders
        html_body_template (str): The template for the email body with placeholders
        attachments (list): A list of filenames to attach (default: None)
    """
    row = df.iloc[row_index]
    to_email = row["E-Mail"]

    # Skip if already sent
    if row[SENT_COLUMN_NAME] == "Yes":
        logging.info(f"E-Mail already sent to {to_email}. Skipping.")
        return

    try:
        # Create a copy of the row data and modify the Vorname to first name only
        row_data = row.to_dict()
        row_data["Vorname"] = row_data["Vorname"].split()[0]  # Get first name only

        # Replace placeholders in subject and body with row data
        subject = subject_template.format(**row_data)
        personalized_body = html_body_template.format(**row_data)

        # Create the email message
        msg = MIMEMultipart('alternative')
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_email
        msg["Subject"] = subject
        msg["Bcc"] = "hi@rigibeats.ch"  # Add BCC

        # Add HTML content
        msg.attach(MIMEText(personalized_body, 'html', 'utf-8'))

        if attachments:
            add_attachments(msg, attachments)

        # Create and send the email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
            logging.info(f"E-Mail successfully sent to {to_email}")

            # Mark as sent and save
            update_sent_status(df, to_email)
            save_updated_excel(df, "email_list.xlsx")

    except Exception as e:
        logging.error(f"Failed to send email to {to_email}: {e}")


# Read email addresses from the Excel file
df, _ = read_email_addresses("email_list.xlsx")

if df is not None:
    # Get attachments
    attachments = get_all_attachments()
    
    # E-Mail templates with placeholders
    subject_template = "RigiBeats 2025 Feedback"
    
    html_body_template = """
    <p>Hey {Vorname}! 👋</p>

    <p><strong>Rigibeats 2025 ist Geschichte! 🎉</strong></p>

    <p>Leider hat uns der Föhn dieses Jahr einen Strich durch die Rechnung gemacht, doch das haltet uns nicht auf!<br>
    Ein riesiges Dankeschön an alle, die am Sonntag dabei waren – die Stimmung war einfach fantastisch!<br>
    Ihr seid die beste Community, und für das sagen wir: <strong>Danke!</strong> 🥳</p>

    <p>Euer Feedback ist uns wichtig  – egal, ob ihr dabei wart oder nicht!<br>
    Nehmt euch kurz Zeit für unsere Umfrage (dauert maximal 5 Minuten):<br>
    <a href="https://docs.google.com/forms/d/e/1FAIpQLSdONIpMxHbSyKgNwAxrHwMj0Y4-yRbpgKjYPi9vIuta2r0XwQ/viewform?usp=dialog">Rigibeats 2025 Feedback</a></p>

    <p>Danke und bis zum nächsten Jahr! 🚀</p>

    <p>Liebe Grüsse<br>
    Max vom Rigibeats Team 👑🏔️❤️‍🔥</p>
    """

    # Process each row
    for index in tqdm(range(len(df)), desc="Processing entries", unit="email"):
        send_personalized_email(df, index, subject_template, html_body_template, attachments)
        
    print("All emails have been processed!")
else:
    print("Failed to read the Excel file. Please check the file path and format.")
