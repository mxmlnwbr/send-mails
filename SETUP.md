# Google Sheets API Setup Guide

This guide will walk you through setting up Google Sheets API access using a Service Account.

## Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click on "Select a project" dropdown at the top
3. Click "NEW PROJECT"
4. Enter a project name (e.g., "RigiBeats Email System")
5. Click "CREATE"

## Step 2: Enable Google Sheets API

1. With your project selected, go to the [API Library](https://console.cloud.google.com/apis/library)
2. Search for "Google Sheets API"
3. Click on it and press "ENABLE"
4. Also search for "Google Drive API" and enable it

## Step 3: Create Service Account

1. Go to [Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Click "CREATE SERVICE ACCOUNT"
3. Enter a name (e.g., "rigibeats-email-bot")
4. Add a description (optional)
5. Click "CREATE AND CONTINUE"
6. Skip the optional steps and click "DONE"

## Step 4: Create and Download Credentials

1. Find your newly created service account in the list
2. Click on it to open details
3. Go to the "KEYS" tab
4. Click "ADD KEY" → "Create new key"
5. Select "JSON" format
6. Click "CREATE"
7. The JSON file will download automatically

## Step 5: Setup Credentials in Your Project

1. Rename the downloaded JSON file to `credentials.json`
2. Move it to your project root directory:
   ```bash
   mv ~/Downloads/your-project-*.json /Users/mxmlnwbr/Documents/projects/send-mails/credentials.json
   ```
3. **IMPORTANT**: Make sure `credentials.json` is in your `.gitignore` file to avoid committing it!

## Step 6: Share Google Sheet with Service Account

1. Open the downloaded `credentials.json` file
2. Find the `client_email` field (looks like: `service-account-name@project-id.iam.gserviceaccount.com`)
3. Copy this email address
4. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1a0WTV7WWd76UiInqn0qzVGLt-2XhDAbivU6ST0Ah2MU/edit
5. Click the "Share" button (top right)
6. Paste the service account email
7. Give it "Editor" permissions
8. Click "Send"

## Step 7: Configure Environment Variables

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and fill in your values:
   ```env
   SMTP_SERVER="smtp.gmail.com"  # or your SMTP server
   SMTP_ACCOUNT="your_email@gmail.com"
   SMTP_PASSWORD="your_app_password"
   
   SPREADSHEET_ID="1a0WTV7WWd76UiInqn0qzVGLt-2XhDAbivU6ST0Ah2MU"
   GOOGLE_CREDENTIALS_FILE="credentials.json"
   ```

### Getting the Spreadsheet ID

The Spreadsheet ID is in the URL of your Google Sheet:
```
https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
```

For your sheet, the ID is: `1a0WTV7WWd76UiInqn0qzVGLt-2XhDAbivU6ST0Ah2MU`

### Gmail SMTP Setup (if using Gmail)

If you're using Gmail, you need to create an App Password:

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable "2-Step Verification" if not already enabled
3. Go to "App passwords"
4. Generate a new app password for "Mail"
5. Use this password in your `.env` file

## Step 8: Run the Script

```bash
uv run python run.py
```

## Troubleshooting

### "Credentials file not found"
- Make sure `credentials.json` is in the project root directory
- Check the `GOOGLE_CREDENTIALS_FILE` path in `.env`

### "Permission denied" or "Unable to parse range"
- Make sure you shared the Google Sheet with the service account email
- Verify the service account has "Editor" permissions

### "Authentication failed" (SMTP)
- If using Gmail, make sure you're using an App Password, not your regular password
- Check that 2-Step Verification is enabled for Gmail
- Verify SMTP server and port are correct

### "Invalid email" warnings
- The script will skip rows without valid email addresses
- Check the "Email" column in your sheet for missing or malformed emails

## Sheet Structure Requirements

Your Google Sheet should have these columns:
- **Email**: Email addresses of recipients
- **Status**: Current status ("1. Offen" for pending, "2. Verschickt" for sent)
- **Access Key**: Will be auto-created by the script if it doesn't exist

The script will:
1. Only process rows with Status = "1. Offen"
2. Skip rows with invalid/missing emails
3. Generate unique access keys (or reuse existing ones)
4. Send emails with the access key
5. Update Status to "2. Verschickt" after successful send
