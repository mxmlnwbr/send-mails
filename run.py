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


def send_emails(to_emails, subject, markdown_body, attachments=None):
    """
    Sends an email to a list of recipients with optional attachments.

    Args:
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

    except Exception as e:
        logging.error(f"Failed to send emails: {e}")


# Usage
email_list = ["maximilian.weber@bluewin.ch", "mxjweber@gmail.com"]

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

send_emails(email_list, "Test Email with Attachments", markdown_body, attachments)
