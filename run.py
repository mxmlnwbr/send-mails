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
import argparse

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
    if not email:
        return False
    
    # Convert to string if it's not already (handles phone numbers as integers)
    email_str = str(email).strip()
    
    # Check if it looks like a phone number (only digits, +, -, spaces, parentheses)
    if email_str.replace('+', '').replace('-', '').replace(' ', '').replace('(', '').replace(')', '').isdigit():
        return False
    
    # Basic email validation
    return "@" in email_str and "." in email_str and len(email_str) > 5


def send_email(to_email, access_key, test_mode=False):
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

        if test_mode:
            logging.info(f"[TEST MODE] Would send email to: {to_email}")
            logging.info(f"[TEST MODE] Subject: {subject}")
            logging.info(f"[TEST MODE] Access Key: {access_key}")
            logging.info(f"[TEST MODE] Body preview: {body[:200]}...")
            return True

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


def generate_keys_only(sheet):
    """Generate access keys for all entries without sending emails."""
    try:
        logging.info("\n" + "="*60)
        logging.info("🔑 KEY GENERATION MODE - Only generating keys, no emails sent!")
        logging.info("="*60 + "\n")
        
        # Get header row first to handle empty columns
        headers = sheet.row_values(1)
        
        # Get all data rows
        all_data = sheet.get_all_values()[1:]  # Skip header row
        
        # Create records manually, handling empty headers
        all_records = []
        for row_data in all_data:
            record = {}
            for idx, header in enumerate(headers):
                if header:  # Only process non-empty headers
                    record[header] = row_data[idx] if idx < len(row_data) else ""
            all_records.append(record)
        
        if not all_records:
            logging.warning("No records found in the sheet")
            return
        
        # Find column indices
        try:
            status_col_idx = headers.index(STATUS_COLUMN) + 1
            
            # Check if Access Key column exists, if not create it
            if ACCESS_KEY_COLUMN in headers:
                key_col_idx = headers.index(ACCESS_KEY_COLUMN) + 1
            else:
                key_col_idx = len(headers) + 1
                sheet.update_cell(1, key_col_idx, ACCESS_KEY_COLUMN)
                logging.info(f"Created '{ACCESS_KEY_COLUMN}' column at position {key_col_idx}")
        except ValueError as e:
            logging.error(f"Required column not found: {e}")
            return

        generated_count = 0
        skipped_count = 0

        # Process each row (starting from row 2, since row 1 is headers)
        for idx, record in enumerate(tqdm(all_records, desc="Generating keys"), start=2):
            status = str(record.get(STATUS_COLUMN, "")).strip()
            email_raw = record.get(EMAIL_COLUMN, "")
            email = str(email_raw).strip() if email_raw else ""

            # Only process entries with status "1. Offen"
            if status != STATUS_OPEN:
                continue

            # Skip if no valid email
            if not is_valid_email(email_raw):
                email_check = str(email_raw).strip()
                if email_check and email_check.replace('+', '').replace('-', '').replace(' ', '').replace('(', '').replace(')', '').isdigit():
                    logging.warning(f"Row {idx}: Phone number detected ({email_check}), skipping")
                else:
                    logging.warning(f"Row {idx}: Invalid or missing email ({email_check}), skipping")
                skipped_count += 1
                continue

            # Check if access key already exists
            existing_key_raw = record.get(ACCESS_KEY_COLUMN, "")
            existing_key = str(existing_key_raw).strip() if existing_key_raw else ""
            
            if existing_key:
                logging.info(f"Row {idx}: Access key already exists for {email}")
                skipped_count += 1
            else:
                # Generate new access key
                access_key = generate_access_key()
                sheet.update_cell(idx, key_col_idx, access_key)
                logging.info(f"Row {idx}: Generated key {access_key} for {email}")
                generated_count += 1

        logging.info(f"\n{'='*60}")
        logging.info(f"🔑 Key Generation Complete!")
        logging.info(f"Keys generated: {generated_count}")
        logging.info(f"Skipped (already has key or invalid): {skipped_count}")
        logging.info(f"{'='*60}")
        logging.info("\n⚠️  NEXT STEPS:")
        logging.info("1. Upload the generated keys to your website")
        logging.info("2. Run 'uv run python run.py' to send emails with the keys")

    except Exception as e:
        logging.error(f"Error generating keys: {e}")


def process_entries(sheet, test_mode=False):
    """Process all entries with status '1. Offen' and send access keys."""
    try:
        if test_mode:
            logging.info("\n" + "="*60)
            logging.info("🧪 TEST MODE ENABLED - No emails will be sent!")
            logging.info("🧪 Sheet status will NOT be updated!")
            logging.info("="*60 + "\n")
        
        # Get header row first to handle empty columns
        headers = sheet.row_values(1)
        
        # Get all data rows
        all_data = sheet.get_all_values()[1:]  # Skip header row
        
        # Create records manually, handling empty headers
        all_records = []
        for row_data in all_data:
            record = {}
            for idx, header in enumerate(headers):
                if header:  # Only process non-empty headers
                    record[header] = row_data[idx] if idx < len(row_data) else ""
            all_records.append(record)
        
        if not all_records:
            logging.warning("No records found in the sheet")
            return
        
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
            # Convert to string and strip to handle both strings and numbers
            status = str(record.get(STATUS_COLUMN, "")).strip()
            email_raw = record.get(EMAIL_COLUMN, "")
            email = str(email_raw).strip() if email_raw else ""

            # Only process entries with status "1. Offen"
            if status != STATUS_OPEN:
                continue

            # Skip if no valid email
            if not is_valid_email(email_raw):
                # Check if it's a phone number
                email_check = str(email_raw).strip()
                if email_check and email_check.replace('+', '').replace('-', '').replace(' ', '').replace('(', '').replace(')', '').isdigit():
                    logging.warning(f"Row {idx}: Phone number detected ({email_check}), skipping (SMS not supported)")
                else:
                    logging.warning(f"Row {idx}: Invalid or missing email ({email_check}), skipping")
                skipped_count += 1
                continue

            # Check if access key exists (required for sending)
            existing_key_raw = record.get(ACCESS_KEY_COLUMN, "")
            access_key = str(existing_key_raw).strip() if existing_key_raw else ""
            
            if not access_key:
                logging.warning(f"Row {idx}: No access key found for {email}, skipping. Run with --generate-keys first!")
                skipped_count += 1
                continue
            
            logging.info(f"Row {idx}: Using access key {access_key} for {email}")

            # Send email
            if send_email(email, access_key, test_mode):
                if not test_mode:
                    # Update status to "2. Verschickt"
                    sheet.update_cell(idx, status_col_idx, STATUS_SENT)
                    logging.info(f"Row {idx}: Updated status to '{STATUS_SENT}' for {email}")
                else:
                    logging.info(f"Row {idx}: [TEST] Would update status to '{STATUS_SENT}' for {email}")
                processed_count += 1
            else:
                logging.error(f"Row {idx}: Failed to send email to {email}, status not updated")
                skipped_count += 1

        logging.info(f"\n{'='*60}")
        if test_mode:
            logging.info(f"🧪 TEST MODE - Simulation Complete!")
            logging.info(f"Would have sent {processed_count} emails")
        else:
            logging.info(f"Processing complete!")
            logging.info(f"Emails sent: {processed_count}")
        logging.info(f"Skipped: {skipped_count}")
        logging.info(f"{'='*60}")

    except Exception as e:
        logging.error(f"Error processing entries: {e}")


def main():
    """Main function to run the script."""
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Send access key emails from Google Sheets',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Step 1: Generate keys only (no emails):
  uv run python run.py --generate-keys
  
  # Step 2: Test mode (simulate without sending):
  uv run python run.py --test
  
  # Step 3: Production mode (sends emails):
  uv run python run.py
        """
    )
    parser.add_argument(
        '--generate-keys',
        action='store_true',
        help='Generate access keys only without sending emails (Step 1)'
    )
    parser.add_argument(
        '--test', '--dry-run',
        action='store_true',
        dest='test_mode',
        help='Test mode: simulate sending without actually sending emails or updating sheet status'
    )
    args = parser.parse_args()
    
    # Determine mode
    if args.generate_keys:
        logging.info("🔑 Running in KEY GENERATION MODE - Only generating keys")
    elif args.test_mode:
        logging.info("🧪 Running in TEST MODE - No emails will be sent, no statuses will be updated")
    else:
        logging.info("🚀 Running in PRODUCTION MODE - Emails will be sent and statuses updated")
    
    logging.info("Starting email distribution script...")
    
    # Validate environment variables
    required_vars = [SPREADSHEET_ID]
    if not args.test_mode and not args.generate_keys:
        required_vars.extend([SMTP_SERVER, EMAIL_ADDRESS, EMAIL_PASSWORD])
    
    if not all(required_vars):
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

    # Run appropriate mode
    if args.generate_keys:
        generate_keys_only(sheet)
    else:
        process_entries(sheet, test_mode=args.test_mode)
    
    logging.info("Script finished.")


if __name__ == "__main__":
    main()
