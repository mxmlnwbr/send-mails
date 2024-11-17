import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
import markdown  # Import markdown to convert Markdown to HTML
from tqdm import tqdm

# Load environment variables
load_dotenv()

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 587  # Using 587 for TLS
EMAIL_ADDRESS = os.getenv("SMTP_ACCOUNT")
EMAIL_PASSWORD = os.getenv("SMTP_PASSWORD")


def send_emails(to_emails, subject, markdown_body):
    """
    Sends an email to a list of recipients with a Markdown-formatted body.

    Args:
        to_emails (list): A list of email addresses to send the email to.
        subject (str): The subject of the email.
        markdown_body (str): The email body written in Markdown format.
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
                # Display the email address being sent
                tqdm.write(f"Sending email to: {to_email}")

                # Create the email for each recipient
                msg = MIMEMultipart()
                msg["From"] = EMAIL_ADDRESS
                msg["To"] = to_email
                msg["Subject"] = subject
                msg.attach(MIMEText(html_body, "html"))

                # Send the email
                server.send_message(msg)

    except Exception as e:
        print(f"Failed to send emails: {e}")


# Usage
email_list = ["maximilian.weber@bluewin.ch", "mxjweber@gmail.com"]

# Markdown Email Body
markdown_body = """
Hi! 👋

This is a **test email** with *Markdown* formatting.

- Item 1
- **Item 2**

[Click here for more info](https://example.com).

Liebe Grüsse
Max von Rigibeats Team 👑🏔️
"""

send_emails(email_list, "Test Markdown Email", markdown_body)
