import sys
from pathlib import Path
import time
import logging

sys.path.append(str(Path(__file__).parent.parent))

from src.gmail_service import GmailService
from src.sheets_service import SheetsService
from src.email_parser import EmailParser
from src.state_manager import StateManager

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    logger.info("Starting Gmail to Sheets automation...")
    
    # Configuration
    SPREADSHEET_ID = "1m7JX9A9I1JiHjzk3Asc-MLXIdBLZJioWBY37kZDbIHw"
    SHEET_NAME = "Sheet1"  # Your sheet tab name
    
    try:
        # Initialize services
        gmail = GmailService()
        sheets = SheetsService(gmail.creds)
        parser = EmailParser()
        state = StateManager()
        
        # Load last processed date
        last_processed = state.load_state()
        logger.info(f"Last processed: {last_processed}")
        
        # Get unread emails since last run
        messages = gmail.get_unread_emails(last_processed)
        
        if not messages:
            logger.info("No new unread emails to process.")
            return
        
        logger.info(f"Found {len(messages)} unread messages")
        
        # Get existing message IDs to prevent duplicates
        existing_ids = sheets.get_existing_message_ids(SPREADSHEET_ID, SHEET_NAME)
        
        processed_count = 0
        latest_email_date = None
        
        # Process emails (limit to reasonable number)
        for i, msg in enumerate(messages[:20]):  # Process max 20 emails per run
            msg_id = msg['id']
            
            # Skip if already processed
            if msg_id in existing_ids:
                logger.debug(f"Skipping already processed email: {msg_id}")
                continue
            
            # Get email details
            email_details = gmail.get_email_details(msg_id)
            if not email_details:
                continue
            
            # Parse email
            parsed_email = parser.parse_email(email_details)
            
            # Append to Google Sheets
            success = sheets.append_row(SPREADSHEET_ID, SHEET_NAME, parsed_email)
            
            if success:
                # Mark as read in Gmail
                gmail.mark_as_read(msg_id)
                processed_count += 1
                
                # Track latest email date for state update
                if latest_email_date is None or parsed_email['date'] > latest_email_date:
                    latest_email_date = parsed_email['date']
                
                logger.info(f"Processed [{i+1}/{len(messages[:20])}]: {parsed_email['subject'][:50]}...")
            else:
                logger.error(f"Failed to append email: {msg_id}")
            
            # Rate limiting
            time.sleep(0.5)
        
        # Update state
        if latest_email_date:
            state.update_last_processed(latest_email_date)
        elif processed_count > 0:
            state.save_state()
        
        logger.info(f"Processing complete! {processed_count} emails added to sheet.")
        
    except Exception as e:
        logger.error(f"Error in main process: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()