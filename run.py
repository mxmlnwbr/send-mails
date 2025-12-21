import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
from tqdm import tqdm
import logging
import gspread
from google.auth.credentials import Credentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
import secrets
import string

load_dotenv()

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("SMTP_ACCOUNT")
EMAIL_PASSWORD = os.getenv("SMTP_PASSWORD")

# Google Sheets Configuration
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

# Column names in the sheet
EMAIL_COLUMN = "Email"
STATUS_COLUMN = "Status"
ACCESS_KEY_COLUMN = "Access Key"
STATUS_OPEN = "1. Offen"
STATUS_SENT = "2. Verschickt"


def generate_access_key(length=16):
    """Generate a unique access key."""
    characters = string.ascii_uppercase + string.digits
    key = ''.join(secrets.choice(characters) for _ in range(length))
    return f"RIGIBEATS-{key[:4]}-{key[4:8]}-{key[8:12]}-{key[12:16]}"


def connect_to_sheet():
    """Connect to Google Sheets using service account credentials."""
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = ServiceAccountCredentials.from_service_account_file(
            CREDENTIALS_FILE, scopes=scopes
        )
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1
        logging.info("Successfully connected to Google Sheets")
        return sheet
    except Exception as e:
        logging.error(f"Error connecting to Google Sheets: {e}")
        return None


def is_valid_email(email):
    """Basic email validation."""
    if not email or not isinstance(email, str):
        return False
    email = email.strip()
    return "@" in email and "." in email and len(email) > 5


def send_email(to_email, access_key):
    """Send email with access key to recipient."""
    try:
        subject = "RigiBeats 2024 - Helfer:innen Tickets"
        body = f"""
        <html>
            <body>
                <p>Hey 👋!</p>
                <p>Erstmal herzlichen Dank, dass du uns RigiBeats 2024 als Helfer:in unterstützt 🎉. Ohne deine Hilfe könnten wir den Event nicht durchführen. Nun schicken wir dir dein gratis Helfer:innen Ticket.</p>
                <p>Dein persönlicher Zugangsschlüssel für 1 Ticket lautet: <strong>{access_key}</strong></p>
                <p>Bitte einfach hier einlösen: https://eventfrog.ch/de/p/party/house-techno/rigibeats-2024-7121106425825163433.html</p>
                <p>Dein RigiBeats Team ❤️</p>
            </body>
        </html>
        """

        msg = MIMEMultipart('alternative')
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
            logging.info(f"Email successfully sent to {to_email}")
            return True
    except Exception as e:
        logging.error(f"Failed to send email to {to_email}: {e}")
        return False


def process_entries(sheet):
    """Process all entries with status '1. Offen' and send access keys."""
    try:
        # Get all records as dictionaries
        all_records = sheet.get_all_records()
        
        if not all_records:
            logging.warning("No records found in the sheet")
            return

        # Get header row to find column indices
        headers = sheet.row_values(1)
        
        # Find column indices
        try:
            email_col_idx = headers.index(EMAIL_COLUMN) + 1
            status_col_idx = headers.index(STATUS_COLUMN) + 1
            
            # Check if Access Key column exists, if not create it
            if ACCESS_KEY_COLUMN in headers:
                key_col_idx = headers.index(ACCESS_KEY_COLUMN) + 1
            else:
                key_col_idx = len(headers) + 1
                sheet.update_cell(1, key_col_idx, ACCESS_KEY_COLUMN)
                logging.info(f"Created '{ACCESS_KEY_COLUMN}' column")
        except ValueError as e:
            logging.error(f"Required column not found: {e}")
            return

        processed_count = 0
        skipped_count = 0

        # Process each row (starting from row 2, since row 1 is headers)
        for idx, record in enumerate(tqdm(all_records, desc="Processing entries"), start=2):
            status = record.get(STATUS_COLUMN, "").strip()
            email = record.get(EMAIL_COLUMN, "").strip()

            # Only process entries with status "1. Offen"
            if status != STATUS_OPEN:
                continue

            # Skip if no valid email
            if not is_valid_email(email):
                logging.warning(f"Row {idx}: Invalid or missing email, skipping")
                skipped_count += 1
                continue

            # Check if access key already exists
            existing_key = record.get(ACCESS_KEY_COLUMN, "").strip()
            if existing_key:
                access_key = existing_key
                logging.info(f"Row {idx}: Using existing access key for {email}")
            else:
                # Generate new access key
                access_key = generate_access_key()
                sheet.update_cell(idx, key_col_idx, access_key)
                logging.info(f"Row {idx}: Generated new access key for {email}")

            # Send email
            if send_email(email, access_key):
                # Update status to "2. Verschickt"
                sheet.update_cell(idx, status_col_idx, STATUS_SENT)
                logging.info(f"Row {idx}: Updated status to '{STATUS_SENT}' for {email}")
                processed_count += 1
            else:
                logging.error(f"Row {idx}: Failed to send email to {email}, status not updated")
                skipped_count += 1

        logging.info(f"\n{'='*50}")
        logging.info(f"Processing complete!")
        logging.info(f"Emails sent: {processed_count}")
        logging.info(f"Skipped: {skipped_count}")
        logging.info(f"{'='*50}")

    except Exception as e:
        logging.error(f"Error processing entries: {e}")


def main():
    """Main function to run the script."""
    logging.info("Starting email distribution script...")
    
    # Validate environment variables
    if not all([SMTP_SERVER, EMAIL_ADDRESS, EMAIL_PASSWORD, SPREADSHEET_ID]):
        logging.error("Missing required environment variables. Check your .env file.")
        return

    # Check if credentials file exists
    if not os.path.exists(CREDENTIALS_FILE):
        logging.error(f"Credentials file not found: {CREDENTIALS_FILE}")
        logging.error("Please follow the setup guide to create service account credentials.")
        return

    # Connect to Google Sheets
    sheet = connect_to_sheet()
    if not sheet:
        logging.error("Failed to connect to Google Sheets. Exiting.")
        return

    # Process entries
    process_entries(sheet)
    
    logging.info("Script finished.")


if __name__ == "__main__":
    main()
