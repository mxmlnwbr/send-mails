import os
from dotenv import load_dotenv
import logging
import gspread
from google.oauth2.service_account import Credentials as ServiceAccountCredentials

load_dotenv()

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Google Sheets Configuration
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

# Column names in the sheet
ACCESS_KEY_COLUMN = "Access Key"


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


def export_access_keys(sheet, output_file="access_keys.txt"):
    """Export all access keys to a comma-separated text file."""
    try:
        # Get header row
        headers = sheet.row_values(1)
        
        # Find Access Key column
        access_key_indices = [i for i, h in enumerate(headers) if h == ACCESS_KEY_COLUMN]
        
        if not access_key_indices:
            logging.error(f"'{ACCESS_KEY_COLUMN}' column not found in the sheet")
            return
        
        key_col_idx = access_key_indices[0] + 1
        logging.info(f"Found '{ACCESS_KEY_COLUMN}' column at position {key_col_idx}")
        
        # Get all values from the Access Key column (skip header)
        all_keys = sheet.col_values(key_col_idx)[1:]  # Skip header row
        
        # Filter out empty values and strip whitespace
        valid_keys = [key.strip() for key in all_keys if key and key.strip()]
        
        if not valid_keys:
            logging.warning("No access keys found in the sheet")
            return
        
        # Create comma-separated string
        keys_string = ", ".join(valid_keys)
        
        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(keys_string)
        
        logging.info(f"\n{'='*60}")
        logging.info(f"✅ Export Complete!")
        logging.info(f"Total keys exported: {len(valid_keys)}")
        logging.info(f"Output file: {output_file}")
        logging.info(f"{'='*60}")
        
        # Show preview of first few keys
        preview_count = min(5, len(valid_keys))
        logging.info(f"\nPreview (first {preview_count} keys):")
        for i, key in enumerate(valid_keys[:preview_count], 1):
            logging.info(f"  {i}. {key}")
        
        if len(valid_keys) > preview_count:
            logging.info(f"  ... and {len(valid_keys) - preview_count} more")
        
    except Exception as e:
        logging.error(f"Error exporting access keys: {e}")


def main():
    """Main function to run the export script."""
    logging.info("Starting access key export script...")
    
    # Validate environment variables
    if not SPREADSHEET_ID:
        logging.error("Missing SPREADSHEET_ID environment variable. Check your .env file.")
        return

    # Check if credentials file exists
    if not os.path.exists(CREDENTIALS_FILE):
        logging.error(f"Credentials file not found: {CREDENTIALS_FILE}")
        return

    # Connect to Google Sheets
    sheet = connect_to_sheet()
    if not sheet:
        logging.error("Failed to connect to Google Sheets. Exiting.")
        return

    # Export access keys
    export_access_keys(sheet)
    
    logging.info("\nScript finished.")


if __name__ == "__main__":
    main()
