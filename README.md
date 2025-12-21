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
- Google Cloud Project with Sheets API enabled

## Installation

### Step 1: Clone and Install Dependencies

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/send-mails.git
   cd send-mails
   ```

2. Install the required dependencies:

   ```bash
   uv sync
   ```

### Step 2: Create Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click on "Select a project" dropdown at the top
3. Click "NEW PROJECT"
4. Enter a project name (e.g., "RigiBeats Email System")
5. Click "CREATE"

### Step 3: Enable Google Sheets API

1. With your project selected, go to the [API Library](https://console.cloud.google.com/apis/library)
2. Search for "Google Sheets API"
3. Click on it and press "ENABLE"
4. Also search for "Google Drive API" and enable it

### Step 4: Create Service Account

1. Go to [Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Click "CREATE SERVICE ACCOUNT"
3. Enter a name (e.g., "rigibeats-email-bot")
4. Add a description (optional)
5. Click "CREATE AND CONTINUE"
6. Skip the optional steps and click "DONE"

### Step 5: Create and Download Credentials

1. Find your newly created service account in the list
2. Click on it to open details
3. Go to the "KEYS" tab
4. Click "ADD KEY" → "Create new key"
5. Select "JSON" format
6. Click "CREATE"
7. The JSON file will download automatically

### Step 6: Setup Credentials in Your Project

1. Rename the downloaded JSON file to `credentials.json`
2. Move it to your project root directory:
   ```bash
   mv ~/Downloads/your-project-*.json /path/to/send-mails/credentials.json
   ```
3. **IMPORTANT**: The file is already in `.gitignore` to avoid committing it

### Step 7: Share Google Sheet with Service Account

1. Open the downloaded `credentials.json` file
2. Find the `client_email` field (looks like: `service-account-name@project-id.iam.gserviceaccount.com`)
3. Copy this email address
4. Open your Google Sheet
5. Click the "Share" button (top right)
6. Paste the service account email
7. Give it "Editor" permissions
8. Click "Send"

### Step 8: Configure Environment Variables

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and fill in your values:
   ```env
   SMTP_SERVER="smtp.gmail.com"  # or your SMTP server
   SMTP_ACCOUNT="your_email@gmail.com"
   SMTP_PASSWORD="your_app_password"
   
   SPREADSHEET_ID="your_spreadsheet_id_here"
   GOOGLE_CREDENTIALS_FILE="credentials.json"
   ```

**Getting the Spreadsheet ID:**

The Spreadsheet ID is in the URL of your Google Sheet:
```
https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
```

**Gmail SMTP Setup (if using Gmail):**

If you're using Gmail, you need to create an App Password:

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable "2-Step Verification" if not already enabled
3. Go to "App passwords"
4. Generate a new app password for "Mail"
5. Use this password in your `.env` file

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

### Common Issues

**"Credentials file not found"**
- Make sure `credentials.json` is in the project root directory
- Check the `GOOGLE_CREDENTIALS_FILE` path in `.env`

**"Permission denied" or "Unable to parse range"**
- Make sure you shared the Google Sheet with the service account email
- Verify the service account has "Editor" permissions

**"Authentication failed" (SMTP)**
- If using Gmail, make sure you're using an App Password, not your regular password
- Check that 2-Step Verification is enabled for Gmail
- Verify SMTP server and port are correct

**"Invalid email" warnings**
- The script will skip rows without valid email addresses
- Check the "Email" column in your sheet for missing or malformed emails

**Phone numbers in Email column**
- Phone numbers are automatically detected and skipped (SMS not supported)
- Script will log: `"Phone number detected (079 xxx xxxx), skipping (SMS not supported)"`

### SMS Support

Currently, only email is supported. If you need SMS support for phone numbers, you would need to:
1. Integrate an SMS service (e.g., Twilio, AWS SNS)
2. Modify the script to detect and send to phone numbers
3. Add SMS credentials to your `.env` file

## Testing

Always test first with the `--test` flag before sending real emails:

```bash
uv run python run.py --test
```

This simulates the entire process without:
- Sending actual emails
- Updating sheet status
- Writing access keys to the sheet (if new)

You'll see output like:
```
🧪 Running in TEST MODE - No emails will be sent, no statuses will be updated
[TEST MODE] Would send email to: user@example.com
[TEST MODE] Access Key: RIGIBEATS-ABCD-1234-EFGH-5678
Row 5: Phone number detected (079 234 5678), skipping (SMS not supported)
🧪 TEST MODE - Simulation Complete!
Would have sent 5 emails
Skipped: 2
```
