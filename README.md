# Send-Mails

**Send-Mails** is a Python-based email-sending utility that sends personalized emails with unique access keys directly from Google Sheets. Perfect for event ticket distribution, access code management, and automated email campaigns.

## Features

- 📧 Send personalized emails with unique access keys
- 📊 Direct Google Sheets integration (no Excel downloads needed)
- 🔐 Automatic access key generation
- ✅ Status tracking ("1. Offen" → "2. Verschickt")
- 🚫 Smart email validation (skips invalid/missing emails)
- 🔄 Resume support (uses existing access keys if present)

## Prerequisites

- Python 3.10 or higher
- A valid SMTP account (e.g., Gmail, Infomaniak, Outlook)
- UV (Python package manager)
- Google Cloud Project with Sheets API enabled (see [SETUP.md](SETUP.md))

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/send-mails.git
   cd send-mails
   ```

2. Install the required dependencies:

   ```bash
   uv sync
   ```

3. Set up Google Sheets API credentials:
   
   **Follow the detailed guide in [SETUP.md](SETUP.md)** for step-by-step instructions on:
   - Creating a Google Cloud Project
   - Enabling Sheets API
   - Creating Service Account
   - Downloading credentials
   - Sharing your Google Sheet with the service account

4. Configure environment variables:

   ```bash
   cp .env.example .env
   ```
   
   Then edit `.env` with your values:
   - SMTP server details
   - Google Spreadsheet ID
   - Path to credentials.json

## Usage

### Test Mode (Recommended First)

Before sending real emails, test the setup:

```bash
uv run python run.py --test
# or
uv run python run.py --dry-run
```

Test mode will:
- ✅ Connect to Google Sheets
- ✅ Read and validate data
- ✅ Show which emails would be sent
- ✅ Display email content previews
- ❌ NOT send actual emails
- ❌ NOT update sheet status

### Production Mode

When ready to send real emails:

```bash
uv run python run.py
```

Production mode will:
1. Connect to your Google Sheet
2. Find all rows with Status = "1. Offen"
3. Validate email addresses
4. Generate unique access keys (format: RIGIBEATS-XXXX-XXXX-XXXX-XXXX)
5. Send personalized emails with access keys
6. Update Status to "2. Verschickt"

## Google Sheet Requirements

Your sheet should have these columns:
- **Email**: Email addresses of recipients (phone numbers will be skipped)
- **Status**: "1. Offen" for pending, "2. Verschickt" for sent
- **Access Key**: Auto-created by the script (stores generated keys)

**Note:** Phone numbers in the Email column will be automatically detected and skipped with a warning. Only valid email addresses will be processed.

## File Structure

```bash
send-mails/
├── run.py                  # Main script with Google Sheets integration
├── SETUP.md                # Detailed setup guide for Google Sheets API
├── .env.example            # Environment variables template
├── credentials.json        # Google service account credentials (you create this)
├── .gitignore              # Git ignore file (excludes .env and credentials.json)
├── uv.lock                 # Project dependencies (lock file)
├── pyproject.toml          # Project configuration
├── README.md               # This file
├── archive/                # Previous versions
└── attachments/            # Optional email attachments
```

## Troubleshooting

See [SETUP.md](SETUP.md) for detailed troubleshooting steps.

Common issues:
- **"Credentials file not found"**: Make sure `credentials.json` is in the project root
- **"Permission denied"**: Share the Google Sheet with your service account email
- **SMTP authentication failed**: Use App Password for Gmail (not regular password)
- **Phone numbers in Email column**: Phone numbers are automatically detected and skipped (SMS not supported)

### SMS Support

Currently, only email is supported. If you need SMS support for phone numbers, you would need to:
1. Integrate an SMS service (e.g., Twilio, AWS SNS)
2. Modify the script to detect and send to phone numbers
3. Add SMS credentials to your `.env` file
