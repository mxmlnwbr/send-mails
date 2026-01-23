"""
Event Information Email Sender for Rigibeats 2026
Sends HTML formatted event information emails in German
- All guests: General event preparation info
- Goldau guests only: Special notice about Eventfrog ticket = RigiBahn ticket
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
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

# CSV file path
CSV_FILE = "data/24 Jan 2026 1300 - Rigibeats 2026 - Orders.csv"
OUTPUT_FILE = "data/24 Jan 2026 1300 - Rigibeats 2026 - Orders_eventinfo_sent.csv"


def is_valid_email(email):
    """Basic email validation."""
    if not email or pd.isna(email):
        return False
    email_str = str(email).strip()
    return "@" in email_str and "." in email_str and len(email_str) > 5


def is_goldau_ticket(category):
    """Check if ticket category is for Goldau route."""
    if not category or pd.isna(category):
        return False
    category_str = str(category).lower()
    return "goldau" in category_str


def send_event_info_email(to_email, first_name="", is_goldau=False, test_mode=False):
    """Send event information email with HTML formatting."""
    try:
        # Personalized greeting
        greeting = f"Hey {first_name}!" if first_name else "Hey!"
        
        # Goldau-specific notice (only for Goldau tickets)
        goldau_notice = ""
        if is_goldau:
            goldau_notice = """
            <p style="background-color: #ffebee; border-left: 4px solid #d32f2f; padding: 15px; margin: 20px 0;">
                <strong style="color: #d32f2f;">⚠️ WICHTIG für Goldau-Reisende:</strong><br>
                Dein Eventfrog-Ticket gilt neu auch als <strong>RigiBahn-Ticket</strong>! 🎫🚂<br>
                Dir wird also kein separates Bahnticket mehr zugestellt.
            </p>
            """
        
        subject = "Rigibeats 2026 - Event Information 🎉❤️"
        
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
                        <h1>🏔️ Rigibeats 2026 🎵</h1>
                    </div>
                    <div class="content">
                        <p>{greeting}</p>
                        
                        {goldau_notice}
                        
                        <p>Wir sind gerade dabei, alles auf der Rigi vorzubereiten und freuen uns riesig auf einen unvergesslichen Event mit euch allen! 🎉</p>
                        
                        <p>Die letzten Vorbereitungen laufen auf Hochtouren:</p>
                        <ul>
                            <li>🎵 Die Bühne wird aufgebaut</li>
                            <li>🍻 Die Bars werden eingerichtet</li>
                            <li>✨ Die Deko wird angebracht</li>
                            <li>🔊 Das Soundsystem wird getestet</li>
                        </ul>
                        
                        <p><strong>Wir sehen uns bald auf der Rigi!</strong> 🏔️❄️</p>
                        
                        <div class="footer">
                            <p>See you on the dancefloor ❤️<br>
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
            logging.info(f"Goldau Ticket: {is_goldau}")
            logging.info(f"Subject: {subject}")
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
    """Process CSV file and send event info emails."""
    try:
        # Read CSV
        df = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
        logging.info(f"Loaded {len(df)} records from CSV")
        
        # Add 'EventInfo Sent' column if it doesn't exist
        if 'EventInfo Sent' not in df.columns:
            df['EventInfo Sent'] = ''
            logging.info("Added 'EventInfo Sent' column")
        
        # If dry run with test email, send only to test email
        if dry_run and test_email:
            logging.info(f"\n{'='*60}")
            logging.info("🧪 DRY RUN MODE - Sending test emails")
            logging.info(f"{'='*60}\n")
            
            # Send test email as Goldau example
            logging.info("📧 Sending GOLDAU version to maximilian.weber@bluewin.ch...")
            success1 = send_event_info_email(
                to_email="maximilian.weber@bluewin.ch",
                first_name="Max",
                is_goldau=True,
                test_mode=False
            )
            
            # Send test email as Gersau example
            logging.info("\n📧 Sending GERSAU version to mxjweber@gmail.com...")
            success2 = send_event_info_email(
                to_email="mxjweber@gmail.com",
                first_name="Max",
                is_goldau=False,
                test_mode=False
            )
            
            if success1 and success2:
                logging.info(f"\n✅ Both test emails sent successfully!")
                logging.info("   - GOLDAU version → maximilian.weber@bluewin.ch (with red notice)")
                logging.info("   - GERSAU version → mxjweber@gmail.com (without red notice)")
                logging.info("\nPlease check both inboxes to verify the format difference!")
            else:
                logging.error(f"\n❌ Failed to send one or both test emails")
            
            # Create test output files to show what the tracking looks like
            logging.info("\n📄 Creating test output files (CSV & Excel) with sample data...")
            df_test = df.copy()
            df_test['EventInfo Sent'] = ''
            # Mark first few entries as "sent" for demonstration
            df_test.at[0, 'EventInfo Sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            df_test.at[1, 'EventInfo Sent'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # Save test files
            timestamp = datetime.now().strftime('%Y-%m-%d_%H%M')
            test_csv = f"data/TEST_eventinfo_sent_{timestamp}.csv"
            test_excel = f"data/TEST_eventinfo_sent_{timestamp}.xlsx"
            
            df_test.to_csv(test_csv, index=False, encoding='utf-8-sig')
            df_test.to_excel(test_excel, index=False, engine='openpyxl')
            
            logging.info(f"✅ Test files created:")
            logging.info(f"   - CSV: {test_csv}")
            logging.info(f"   - Excel: {test_excel}")
            logging.info(f"\n💡 Open the Excel file to see how the 'EventInfo Sent' column will look!")
            
            return
        
        # If simulate mode, show what would happen
        if simulate:
            logging.info(f"\n{'='*60}")
            logging.info(f"🔍 SIMULATION MODE - Showing email content variations")
            logging.info(f"{'='*60}\n")
            
            # Track unique emails to detect duplicates
            email_tracker = {}
            goldau_count = 0
            gersau_count = 0
            
            for idx, row in df.iterrows():
                email = row.get('Email', '')
                first_name = row.get('First name', '')
                category = row.get('Category', '')
                
                if not is_valid_email(email):
                    continue
                
                # Track for duplicates
                if email in email_tracker:
                    email_tracker[email] += 1
                else:
                    email_tracker[email] = 1
                
                is_goldau = is_goldau_ticket(category)
                
                if is_goldau:
                    goldau_count += 1
                else:
                    gersau_count += 1
                
                # Show first Goldau and first Gersau examples
                if (is_goldau and goldau_count == 1) or (not is_goldau and gersau_count == 1):
                    route = "GOLDAU" if is_goldau else "GERSAU"
                    logging.info(f"\n{'─'*60}")
                    logging.info(f"📧 Example {route} Email:")
                    logging.info(f"{'─'*60}")
                    logging.info(f"To: {email}")
                    logging.info(f"Name: {first_name}")
                    logging.info(f"Category: {category}")
                    logging.info(f"Includes Goldau Notice: {'YES (Red box)' if is_goldau else 'NO'}")
                    logging.info(f"{'─'*60}\n")
            
            # Check for duplicates
            duplicates = {email: count for email, count in email_tracker.items() if count > 1}
            
            logging.info(f"\n{'='*60}")
            logging.info("📊 Simulation Summary:")
            logging.info(f"Total unique emails: {len(email_tracker)}")
            logging.info(f"Total entries in CSV: {goldau_count + gersau_count}")
            logging.info(f"Goldau recipients (with red notice): {goldau_count}")
            logging.info(f"Gersau recipients (no red notice): {gersau_count}")
            logging.info(f"Duplicate email addresses: {len(duplicates)}")
            if duplicates:
                logging.info(f"\n✅ DEDUPLICATION: Each email will receive only ONE email")
                logging.info(f"   Total emails that will be sent: {len(email_tracker)}")
                logging.info(f"   Duplicate entries that will be skipped: {(goldau_count + gersau_count) - len(email_tracker)}")
            else:
                logging.info("✅ No duplicate emails - each person will receive only ONE email")
            logging.info(f"{'='*60}\n")
            
            return
        
        # Production mode - process all records
        logging.info(f"\n{'='*60}")
        logging.info("🚀 PRODUCTION MODE - Sending emails to all recipients")
        logging.info(f"{'='*60}\n")
        
        # First pass: Group by email and prioritize Goldau tickets
        logging.info("📋 Analyzing tickets by email address...")
        email_groups = {}
        for idx, row in df.iterrows():
            email = row.get('Email', '')
            if not is_valid_email(email):
                continue
            if row.get('EventInfo Sent', ''):
                continue  # Skip already sent
                
            category = row.get('Category', '')
            is_goldau = is_goldau_ticket(category)
            
            if email not in email_groups:
                email_groups[email] = {
                    'indices': [idx],
                    'first_name': row.get('First name', ''),
                    'is_goldau': is_goldau
                }
            else:
                email_groups[email]['indices'].append(idx)
                # Prioritize Goldau: if any ticket is Goldau, send Goldau version
                if is_goldau:
                    email_groups[email]['is_goldau'] = True
        
        logging.info(f"Found {len(email_groups)} unique email addresses to process")
        
        sent_count = 0
        skipped_count = 0
        already_sent_count = 0
        duplicate_count = 0
        
        # Second pass: Send emails with correct version
        for email, info in tqdm(email_groups.items(), desc="Sending emails"):
            first_name = info['first_name']
            is_goldau = info['is_goldau']
            indices = info['indices']
            
            # Send email
            if send_event_info_email(email, first_name, is_goldau, test_mode=False):
                # Mark ALL rows with this email as sent
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                for idx in indices:
                    df.at[idx, 'EventInfo Sent'] = timestamp
                
                sent_count += 1
                if len(indices) > 1:
                    duplicate_count += len(indices) - 1
            else:
                skipped_count += 1
        
        # Save updated CSV with timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d_%H%M')
        output_file_timestamped = OUTPUT_FILE.replace('.csv', f'_{timestamp}.csv')
        df.to_csv(output_file_timestamped, index=False, encoding='utf-8-sig')
        
        # Also save as Excel for better visibility
        output_excel = output_file_timestamped.replace('.csv', '.xlsx')
        df.to_excel(output_excel, index=False, engine='openpyxl')
        
        logging.info(f"\n{'='*60}")
        logging.info("📊 Summary:")
        logging.info(f"✅ Emails sent: {sent_count}")
        logging.info(f"⏭️  Already sent: {already_sent_count}")
        logging.info(f"🔄 Duplicates skipped: {duplicate_count}")
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
        description='Send event information emails for Rigibeats 2026',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simulate - Check content variations and duplicates:
  python run_eventinfo.py --simulate
  
  # Dry run - Send test email to check format:
  python run_eventinfo.py --dry-run
  
  # Production - Send to all recipients:
  python run_eventinfo.py
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
        help='Simulate mode: analyze CSV and show content variations without sending emails'
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
