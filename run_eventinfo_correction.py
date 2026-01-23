"""
Correction Email for Celine Camenzind
Sends the Goldau-specific railway ticket information
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
import logging

load_dotenv()

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("SMTP_ACCOUNT")
EMAIL_PASSWORD = os.getenv("SMTP_PASSWORD")


def send_correction_email(to_email="celine.camenzind@bluewin.ch", first_name="Celine"):
    """Send correction email with Goldau railway ticket info."""
    try:
        greeting = f"Hey {first_name}!" if first_name else "Hey!"
        
        subject = "Rigibeats 2026 - Wichtige Info für Goldau-Tickets 🎫🚂"
        
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
                    .important-notice {{
                        background-color: #ffebee;
                        border-left: 4px solid #d32f2f;
                        padding: 15px;
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
                        <h1>🏔️ Rigibeats 2026 🎵</h1>
                    </div>
                    <div class="content">
                        <p>{greeting}</p>
                        
                        <p>Wir haben gesehen, dass du sowohl Goldau- als auch Gersau-Tickets gekauft hast. Hier noch eine wichtige Info für deine Goldau-Tickets:</p>
                        
                        <div class="important-notice">
                            <strong style="color: #d32f2f;">⚠️ WICHTIG für Goldau-Reisende:</strong><br>
                            Dein Eventfrog-Ticket gilt neu auch als <strong>RigiBahn-Ticket</strong>! 🎫🚂<br>
                            Dir wird also kein separates Bahnticket mehr zugestellt.
                        </div>
                        
                        <p>Für deine Gersau-Tickets ändert sich nichts - diese Info betrifft nur die Anreise via Goldau.</p>
                        
                        <p>Wir freuen uns auf dich! 🎉</p>
                        
                        <div class="footer">
                            <p>See you on the dancefloor ❤️<br>
                            <strong>Rigibeats Team</strong></p>
                        </div>
                    </div>
                </div>
            </body>
        </html>
        """

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
            logging.info(f"✅ Correction email sent to {to_email}")
            return True
            
    except Exception as e:
        logging.error(f"❌ Failed to send correction email to {to_email}: {e}")
        return False


def main():
    """Main function."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Send correction email for Goldau ticket info')
    parser.add_argument('--demo', action='store_true', help='Send demo to maximilian.weber@bluewin.ch')
    args = parser.parse_args()
    
    if args.demo:
        logging.info("🧪 DEMO MODE - Sending to maximilian.weber@bluewin.ch")
        logging.info("="*60)
        success = send_correction_email(to_email="maximilian.weber@bluewin.ch", first_name="Max")
    else:
        logging.info("📧 Sending correction email to Celine Camenzind...")
        logging.info("="*60)
        
        # Confirm before sending
        print("\n⚠️  About to send correction email to:")
        print("   Email: celine.camenzind@bluewin.ch")
        print("   Name: Celine")
        print("   Content: Goldau railway ticket information")
        print()
        confirm = input("Continue? (yes/no): ")
        
        if confirm.lower() != 'yes':
            logging.info("Cancelled by user")
            return
        
        # Send email
        success = send_correction_email()
    
    if success:
        logging.info("\n" + "="*60)
        logging.info("✅ Correction email sent successfully!")
        logging.info("="*60)
    else:
        logging.error("\n" + "="*60)
        logging.error("❌ Failed to send correction email")
        logging.error("="*60)


if __name__ == "__main__":
    main()
