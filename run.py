import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv
import markdown
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
        email_addresses = df[df[SENT_COLUMN_NAME] != "Yes"]["Email"].dropna().tolist()
        logging.info(
            f"Successfully read {len(email_addresses)} email addresses that need to be sent."
        )
        return df, email_addresses

    except Exception as e:
        logging.error(f"Error reading email addresses from {excel_filename}: {e}")
        return None, []


def update_sent_status(df, email):
    """
    Updates the 'Sent' status for an email in the dataframe to 'Yes'.

    Args:
        df (DataFrame): The dataframe containing email addresses.
        email (str): The email address to mark as sent.
    """
    df.loc[df["Email"] == email, SENT_COLUMN_NAME] = "Yes"
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


def send_emails(df, to_emails, subject, markdown_body, attachments=None):
    """
    Sends an email to a list of recipients with optional attachments.

    Args:
        df (DataFrame): The dataframe containing email addresses and their 'Sent' status.
        to_emails (list): A list of email addresses to send the email to.
        subject (str): The subject of the email.
        markdown_body (str): The email body written in Markdown format.
        attachments (list): A list of filenames to attach (default: None).
    """
    # Convert Markdown to HTML
    html_body = markdown.markdown(markdown_body)

    try:
        # Connect to the SMTP server
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Upgrade the connection to secure
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

            # Send email to each recipient with a progress bar
            for to_email in tqdm(to_emails, desc="Sending Emails", unit="email"):
                # Skip emails that have already been sent
                if df.loc[df["Email"] == to_email, SENT_COLUMN_NAME].values[0] == "Yes":
                    logging.info(f"Email already sent to {to_email}. Skipping.")
                    continue

                logging.info(f"Sending email to: {to_email}")

                # Create the email for each recipient
                msg = MIMEMultipart()
                msg["From"] = EMAIL_ADDRESS
                msg["To"] = to_email
                msg["Subject"] = subject

                # Attach the email body
                msg.attach(MIMEText(html_body, "html"))

                # Add attachments if any
                if attachments:
                    add_attachments(msg, attachments)

                # Send the email
                server.send_message(msg)
                logging.info(f"Email successfully sent to {to_email}")

                # Mark this email as sent in the dataframe
                update_sent_status(df, to_email)

                # Save the updated dataframe after each email to avoid losing progress
                save_updated_excel(df, "email_list.xlsx")

    except Exception as e:
        logging.error(f"Failed to send emails: {e}")


# Read email addresses from the Excel file located in the 'data' folder
df, email_list = read_email_addresses(
    "email_list.xlsx"
)  # Replace with the actual filename

# Markdown Email Body
markdown_body = """
# Hello, Maximilian!

This is a **test email** with *Markdown* formatting and attachments.

- Item 1
- Item 2

[Click here for more info](https://example.com).

Best regards,  
**Your Team**
"""

# File Attachments (from /attachments folder)
attachments = ["example.pdf", "image.png"]  # Just filenames, not full paths

if email_list:
    send_emails(
        df, email_list, "Test Email with Attachments", markdown_body, attachments
    )
else:
    logging.warning("No email addresses found. Please check the Excel file.")
