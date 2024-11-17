# Send-Mails

**Send-Mails** is a Python-based email-sending utility that enables you to easily send emails using SMTP servers. This project supports environment-based configuration and allows sending both plain text and multipart messages.

## Prerequisites

- Python 3.7 or higher
- A valid SMTP account (e.g., Infomaniak, Gmail, Outlook)
- Poetry

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/send-mails.git
   cd send-mails
   ```

2. Install the required dependencies:

   ```bash
   poetry shell
   poetry install
   ```

3. Rename the .env.template and fill out the Variables. (Note: For Gmail accounts, you might need to enable "Less Secure Apps" or generate an App Password.)

## Usage

1. Activate Environment:

   ```bash
   poetry shell
   ```

2. Run the script:

   ```bash
   python run.py
   ```

## File Structure

    send-mails/
    ├── run.py # Main script containing the email-sending logic
    ├── .env # Environment variables (ignored by version control)
    ├── .gitignore
    ├── poetry.lock # Project dependencies
    ├── pyprohect.toml # Project dependencies
    └── README.md # Project documentation
