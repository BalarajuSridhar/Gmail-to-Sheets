import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
CREDENTIALS_PATH = BASE_DIR / "credentials" / "credentials.json"
TOKEN_PATH = BASE_DIR / "credentials" / "token.json"
STATE_PATH = BASE_DIR / "data" / "last_processed.json"

# Scopes for API access
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Google Sheet configuration
SPREADSHEET_ID = "1-v4qcX1d2pRu0YLd-L9CdKp1mXKjHIDHVoJKcCUMbok"  # You'll get this from your sheet URL
SHEET_NAME = "Email Log"

# Email processing
MAX_RESULTS = 50  # Max emails to process per run