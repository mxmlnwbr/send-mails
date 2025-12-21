# Send-Mails

Automated email distribution with unique access keys from Google Sheets.

## Prerequisites

- Python 3.10+, UV, Gmail/SMTP account
- Google Cloud Project with Sheets API enabled

## Setup

```bash
# Install
uv sync

# Google Cloud setup:
# 1. Create project: https://console.cloud.google.com/
# 2. Enable APIs: Google Sheets API + Google Drive API
# 3. Create Service Account: https://console.cloud.google.com/iam-admin/serviceaccounts
# 4. Create JSON key, save as credentials.json in project root
# 5. Share your Google Sheet with service account email (Editor access)

# Configure
cp .env.example .env
# Edit .env with your SMTP and SPREADSHEET_ID
```

**.env file:**
```env
SMTP_SERVER="smtp.gmail.com"
SMTP_ACCOUNT="your_email@gmail.com"
SMTP_PASSWORD="your_app_password"
SPREADSHEET_ID="your_spreadsheet_id"
GOOGLE_CREDENTIALS_FILE="credentials.json"
```

Get SPREADSHEET_ID from URL: `docs.google.com/spreadsheets/d/[ID]/edit`  
Gmail users: Create App Password at https://myaccount.google.com/security

## Usage

```bash
# Step 1: Generate all access keys (no emails sent)
uv run python run.py --generate-keys

# Step 2: Upload keys to your website (manual step)

# Step 3: Test email sending (simulation only)
uv run python run.py --test

# Step 4: Send emails (production)
uv run python run.py
```

## Sheet Requirements

Required columns:
- **Email** - Valid emails only (phone numbers skipped)
- **Status** - "1. Offen" (pending) or "2. Verschickt" (sent)
- **Access Key** - Auto-created by `--generate-keys` command

## Workflow

1. **Generate keys**: Creates Access Key column and generates unique keys for all "1. Offen" entries
2. **Upload to website**: Manually upload the generated keys to activate them on your platform
3. **Test sending**: Verify email content and recipients without actually sending
4. **Send emails**: Emails are sent with existing keys, status updated to "2. Verschickt"

## Troubleshooting

- **No access key found**: Run `--generate-keys` first before sending emails
- **Credentials not found**: Check `credentials.json` is in project root
- **Permission denied**: Share sheet with service account email (Editor access)
- **SMTP failed**: Use Gmail App Password, not regular password
- **Phone numbers**: Automatically skipped (SMS not supported)
