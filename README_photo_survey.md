# Feedback Survey Email Script

This script sends a feedback survey email to Rigibeats 2026 attendees, asking them to share their opinions about what went well and what could be improved.

## Features

- Sends personalized HTML emails in German
- Includes embedded event photo (data/_DSC0127.jpg)
- Contains link to Google Forms for feedback submission
- Tracks sent status in CSV to avoid duplicates
- Groups by email address (no duplicate sends)

## Usage

### 1. Simulate (check what will be sent)
```bash
uv run python run_photo_survey.py --simulate
```

### 2. Test (send one test email to verify format)
```bash
uv run python run_photo_survey.py --dry-run
```
This sends a test email to `maximilian.weber@bluewin.ch`

### 3. Production (send to all recipients)
```bash
uv run python run_photo_survey.py
```
⚠️ You will be asked to confirm before sending

## Email Content

The email includes:
- Personalized greeting (using first name from CSV)
- Event photo embedded in email
- German text asking for feedback about what went well and what could be improved
- Call-to-action button linking to Google Forms
- Forms link: https://docs.google.com/forms/d/e/1FAIpQLSdG-ARZABazBueDmIJZT42bK08MMznhuRXs93bT-r2hoLf8KA/viewform?usp=dialog

## Output

After sending, the script creates:
- CSV file with timestamp tracking when emails were sent
- Excel file for easy viewing of send status

Files are saved as:
- `data/24 Jan 2026 1300 - Rigibeats 2026 - Orders_feedback_survey_sent_YYYY-MM-DD_HHMM.csv`
- `data/24 Jan 2026 1300 - Rigibeats 2026 - Orders_feedback_survey_sent_YYYY-MM-DD_HHMM.xlsx`

## Note

The image file is quite large (44 MB), so email sending may take a bit longer per email.
