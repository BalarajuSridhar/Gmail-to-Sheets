# main.py
import sys
from pathlib import Path
import time

sys.path.append(str(Path(__file__).parent.parent))

from src.gmail_service import GmailService
from src.sheets_service import SheetsService
from src.email_parser import EmailParser
from src.state_manager import StateManager

def main():
    print("Starting Gmail to Sheets automation...")
    print("=" * 50)
    
    # Configuration
    SPREADSHEET_ID = "1m7JX9A9I1JiHjzk3Asc-MLXIdBLZJioWBY37kZDbIHw"  # Your Sheet ID
    SHEET_NAME = "Sheet1"  # Your sheet tab name
    
    try:
        # Initialize services
        print("1. Authenticating with Gmail...")
        gmail = GmailService()
        
        print("2. Initializing Sheets service...")
        sheets = SheetsService(gmail.creds)
        
        print("3. Formatting Google Sheet...")
        # Format or create sheet with proper headers
        sheets.create_or_reset_sheet(SPREADSHEET_ID, SHEET_NAME)
        
        parser = EmailParser()
        state = StateManager()
        
        # Load last processed date
        last_processed = state.load_state()
        print(f"4. Last processed: {last_processed}")
        
        # Get unread emails since last run
        print("5. Fetching unread emails...")
        messages = gmail.get_unread_emails(last_processed)
        
        if not messages:
            print("✓ No new unread emails to process.")
            return
        
        print(f"6. Found {len(messages)} unread messages")
        
        # Get existing message IDs to prevent duplicates
        existing_ids = sheets.get_existing_message_ids(SPREADSHEET_ID, SHEET_NAME)
        print(f"7. Found {len(existing_ids)} already processed emails")
        
        processed_count = 0
        latest_email_date = None
        
        # Process emails (limit to reasonable number)
        max_emails = min(10, len(messages))  # Process max 10 emails
        print(f"8. Processing {max_emails} emails...")
        
        for i, msg in enumerate(messages[:max_emails]):
            msg_id = msg['id']
            
            # Skip if already processed
            if msg_id in existing_ids:
                print(f"   [{i+1}/{max_emails}] Skipping already processed: {msg_id[:10]}...")
                continue
            
            # Get email details
            email_details = gmail.get_email_details(msg_id)
            if not email_details:
                print(f"   [{i+1}/{max_emails}] Failed to fetch email details")
                continue
            
            # Parse email
            parsed_email = parser.parse_email(email_details)
            parsed_email['message_id'] = msg_id  # Add message_id for duplicate tracking
            
            print(f"   [{i+1}/{max_emails}] Processing: {parsed_email['subject'][:40]}...")
            
            # Append to Google Sheets
            success = sheets.append_row(SPREADSHEET_ID, SHEET_NAME, parsed_email)
            
            if success:
                # Mark as read in Gmail
                gmail.mark_as_read(msg_id)
                processed_count += 1
                
                # Track latest email date for state update
                if latest_email_date is None or parsed_email['date'] > latest_email_date:
                    latest_email_date = parsed_email['date']
                
                print(f"     ✓ Added to sheet")
            else:
                print(f"     ✗ Failed to add to sheet")
            
            # Rate limiting
            time.sleep(0.5)
        
        # Update state
        if latest_email_date:
            state.update_last_processed(latest_email_date)
        elif processed_count > 0:
            state.save_state()
        
        print("\n" + "=" * 50)
        print(f"✅ Processing complete! {processed_count} emails added to sheet.")
        print(f"📊 Sheet URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n✗ Error in main process: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()