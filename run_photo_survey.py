"""
Feedback Survey Email Sender for Rigibeats 2026
Sends HTML formatted email with Google Forms link and event photo
Asks attendees to share feedback about what went well and what could be improved
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import os
from dotenv import load_dotenv
from tqdm import tqdm
import logging
import pandas as pd
import argparse
from datetime import datetime

load_dotenv()

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("SMTP_ACCOUNT")
EMAIL_PASSWORD = os.getenv("SMTP_PASSWORD")

# File paths
CSV_FILE = "data/24 Jan 2026 1300 - Rigibeats 2026 - Orders.csv"
IMAGE_FILE = "data/_DSC0127.jpg"
OUTPUT_FILE = "data/24 Jan 2026 1300 - Rigibeats 2026 - Orders_feedback_survey_sent.csv"

# Google Forms link
FORMS_LINK = "https://docs.google.com/forms/d/e/1FAIpQLSdG-ARZABazBueDmIJZT42bK08MMznhuRXs93bT-r2hoLf8KA/viewform?usp=dialog"


def is_valid_email(email):
    """Basic email validation."""
    if not email or pd.isna(email):
        return False
    email_str = str(email).strip()
    return "@" in email_str and "." in email_str and len(email_str) > 5


def send_feedback_survey_email(to_email, first_name="", image_path=IMAGE_FILE, test_mode=False):
    """Send feedback survey email with HTML formatting and attached image."""
    try:
        # Personalized greeting
        greeting = f"Hey {first_name}!" if first_name else "Hey!"
        
        subject = "Rigibeats 2026 - Dein Feedback 💭"
        
        # HTML email body with embedded image
        body = f"""
        <html>
            <head>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        line-height: 1.6;
                        color: #333;
                    }}
                    .container {{
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 30px;
                        border-radius: 10px 10px 0 0;
                        text-align: center;
                    }}
                    .content {{
                        background-color: #ffffff;
                        padding: 30px;
                        border-radius: 0 0 10px 10px;
                    }}
                    .cta-button {{
                        display: inline-block;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 15px 30px;
                        text-decoration: none;
                        border-radius: 5px;
                        font-weight: bold;
                        margin: 20px 0;
                    }}
                    .footer {{
                        text-align: center;
                        margin-top: 30px;
                        color: #666;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1>💭 Dein Feedback zählt! 💭</h1>
                    </div>
                    <div class="content">
                        <p>{greeting}</p>
                        
                        <p>Was für ein unvergesslicher Tag auf der Rigi! 🏔️🎉</p>
                        
                        <p><strong>Wie hat dir das Event gefallen?</strong></p>
                        
                        <p>Deine Meinung ist uns wichtig! Wir möchten von dir erfahren, was dir besonders 
                        gut gefallen hat und wo wir uns noch verbessern können. Dein ehrliches Feedback 
                        hilft uns, die nächsten Events noch besser zu machen.</p>
                        
                        <p>Das Ausfüllen dauert nur wenige Minuten und ist für uns unglaublich wertvoll! 🙏</p>
                        
                        <p style="text-align: center;">
                            <a href="{FORMS_LINK}" class="cta-button">
                                📝 Jetzt Feedback geben
                            </a>
                        </p>
                        
                        <p>Vielen Dank, dass du dabei warst und diese Nacht so besonders gemacht hast! ❤️</p>
                        
                        <div class="footer">
                            <p>See you next time! 🎵<br>
                            <strong>Rigibeats Team</strong></p>
                        </div>
                    </div>
                </div>
            </body>
        </html>
        """

        if test_mode:
            logging.info(f"✉️  Test Email Preview:")
            logging.info(f"To: {to_email}")
            logging.info(f"Name: {first_name}")
            logging.info(f"Subject: {subject}")
            logging.info(f"Image: {image_path}")
            logging.info(f"Forms Link: {FORMS_LINK}")
            logging.info(f"{'─'*60}")
            return True

        msg = MIMEMultipart('alternative')
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_email
        msg["Bcc"] = "hi@rigibeats.ch"
        msg["Subject"] = subject
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
            logging.info(f"✅ Email sent to {to_email}")
            return True
            
    except Exception as e:
        logging.error(f"❌ Failed to send email to {to_email}: {e}")
        return False


def process_csv_and_send_emails(dry_run=False, test_email=None, simulate=False):
    """Process CSV file and send feedback survey emails."""
    try:
        # Read CSV
        df = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
        logging.info(f"Loaded {len(df)} records from CSV")
        
        # Add 'Feedback Survey Sent' column if it doesn't exist
        if 'Feedback Survey Sent' not in df.columns:
            df['Feedback Survey Sent'] = ''
            logging.info("Added 'Feedback Survey Sent' column")
        
        # If dry run with test email, send only to test email
        if dry_run and test_email:
            logging.info(f"\n{'='*60}")
            logging.info("🧪 DRY RUN MODE - Sending test email")
            logging.info(f"{'='*60}\n")
            
            success = send_feedback_survey_email(
                to_email=test_email,
                first_name="Max",
                test_mode=False
            )
            
            if success:
                logging.info(f"\n✅ Test email sent successfully to {test_email}!")
                logging.info("\nPlease check the inbox to verify the format!")
            else:
                logging.error(f"\n❌ Failed to send test email")
            
            return
        
        # If simulate mode, show what would happen
        if simulate:
            logging.info(f"\n{'='*60}")
            logging.info(f"🔍 SIMULATION MODE - Preview")
            logging.info(f"{'='*60}\n")
            
            # Track unique emails
            email_tracker = {}
            total_entries = 0
            already_sent = 0
            
            for idx, row in df.iterrows():
                email = row.get('Email', '')
                first_name = row.get('First name', '')
                
                if not is_valid_email(email):
                    continue
                
                if row.get('Feedback Survey Sent', ''):
                    already_sent += 1
                    continue  # Skip already sent
                
                total_entries += 1
                
                if email in email_tracker:
                    email_tracker[email] += 1
                else:
                    email_tracker[email] = 1
            
            # Count duplicates
            duplicate_count = sum(count - 1 for count in email_tracker.values() if count > 1)
            
            logging.info(f"📊 Simulation Summary:")
            logging.info(f"   Total entries with valid emails: {total_entries}")
            logging.info(f"   Unique email addresses: {len(email_tracker)}")
            logging.info(f"   Duplicate entries: {duplicate_count}")
            logging.info(f"   Already sent: {already_sent}")
            logging.info(f"\n   ✅ DEDUPLICATION: Each email will receive ONLY ONE email!")
            logging.info(f"   📧 Total emails that will be sent: {len(email_tracker)}")
            logging.info(f"\n   Image file: {IMAGE_FILE}")
            logging.info(f"   Image size: {os.path.getsize(IMAGE_FILE) / 1024 / 1024:.2f} MB")
            logging.info(f"   Forms link: {FORMS_LINK}")
            logging.info(f"{'='*60}\n")
            
            return
        
        # Production mode - process all records
        logging.info(f"\n{'='*60}")
        logging.info("🚀 PRODUCTION MODE - Sending emails to all recipients")
        logging.info(f"{'='*60}\n")
        
        # Group by email to avoid duplicates
        logging.info("📋 Analyzing emails and grouping duplicates...")
        email_groups = {}
        total_rows = 0
        already_sent_count = 0
        
        for idx, row in df.iterrows():
            email = row.get('Email', '')
            if not is_valid_email(email):
                continue
            if row.get('Feedback Survey Sent', ''):
                already_sent_count += 1
                continue  # Skip already sent
            
            total_rows += 1
            
            if email not in email_groups:
                email_groups[email] = {
                    'indices': [idx],
                    'first_name': row.get('First name', '')
                }
            else:
                email_groups[email]['indices'].append(idx)
        
        # Calculate duplicates
        duplicate_entries = total_rows - len(email_groups)
        
        logging.info(f"\n📊 Email Analysis:")
        logging.info(f"   Total entries with valid emails: {total_rows}")
        logging.info(f"   Unique email addresses: {len(email_groups)}")
        logging.info(f"   Duplicate entries (will be skipped): {duplicate_entries}")
        logging.info(f"   Already sent (will be skipped): {already_sent_count}")
        logging.info(f"   ✅ Each person will receive ONLY ONE email!\n")
        
        sent_count = 0
        skipped_count = 0
        
        # Send emails
        for email, info in tqdm(email_groups.items(), desc="Sending emails"):
            first_name = info['first_name']
            indices = info['indices']
            
            if send_feedback_survey_email(email, first_name, test_mode=False):
                # Mark ALL rows with this email as sent
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                for idx in indices:
                    df.at[idx, 'Feedback Survey Sent'] = timestamp
                sent_count += 1
            else:
                skipped_count += 1
        
        # Save updated CSV with timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d_%H%M')
        output_file_timestamped = OUTPUT_FILE.replace('.csv', f'_{timestamp}.csv')
        df.to_csv(output_file_timestamped, index=False, encoding='utf-8-sig')
        
        # Also save as Excel
        output_excel = output_file_timestamped.replace('.csv', '.xlsx')
        df.to_excel(output_excel, index=False, engine='openpyxl')
        
        logging.info(f"\n{'='*60}")
        logging.info("📊 Summary:")
        logging.info(f"✅ Emails sent: {sent_count}")
        logging.info(f"❌ Skipped/Failed: {skipped_count}")
        logging.info(f"💾 Output saved to:")
        logging.info(f"   - CSV: {output_file_timestamped}")
        logging.info(f"   - Excel: {output_excel}")
        logging.info(f"{'='*60}\n")
        
    except Exception as e:
        logging.error(f"Error processing CSV: {e}")
        raise


def main():
    """Main function."""
    parser = argparse.ArgumentParser(
        description='Send photo survey emails for Rigibeats 2026',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simulate - Check what would be sent:
  python run_photo_survey.py --simulate
  
  # Dry run - Send test email to check format:
  python run_photo_survey.py --dry-run
  
  # Production - Send to all recipients:
  python run_photo_survey.py
        """
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Dry run mode: send test email to maximilian.weber@bluewin.ch only'
    )
    parser.add_argument(
        '--simulate',
        action='store_true',
        help='Simulate mode: analyze CSV without sending emails'
    )
    args = parser.parse_args()
    
    if args.simulate:
        logging.info("🔍 Starting in SIMULATION mode")
        process_csv_and_send_emails(simulate=True)
    elif args.dry_run:
        logging.info("🧪 Starting in DRY RUN mode")
        process_csv_and_send_emails(
            dry_run=True, 
            test_email="maximilian.weber@bluewin.ch"
        )
    else:
        logging.info("🚀 Starting in PRODUCTION mode")
        confirm = input("⚠️  This will send emails to ALL recipients. Continue? (yes/no): ")
        if confirm.lower() != 'yes':
            logging.info("Cancelled by user")
            return
        process_csv_and_send_emails(dry_run=False)
    
    logging.info("Script finished.")


if __name__ == "__main__":
    main()
