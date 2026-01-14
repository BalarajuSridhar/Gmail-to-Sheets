# src/gmail_service.py
import os
import base64
from pathlib import Path
import sys
import pickle

sys.path.append(str(Path(__file__).parent.parent))

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

class GmailService:
    def __init__(self):
        self.creds = None
        self.service = None
        self._authenticate()
    
    def _authenticate(self):
        """Authenticate using OAuth 2.0"""
        BASE_DIR = Path(__file__).parent.parent
        CREDENTIALS_PATH = BASE_DIR / "credentials" / "credentials.json"
        TOKEN_PATH = BASE_DIR / "credentials" / "token.json"
        
        # Updated scopes - now includes modify permission
        SCOPES = [
            'https://www.googleapis.com/auth/gmail.modify',  # Changed from readonly
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        
        if os.path.exists(TOKEN_PATH):
            with open(TOKEN_PATH, 'rb') as token:
                self.creds = pickle.load(token)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(CREDENTIALS_PATH), SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            # Save the token
            with open(TOKEN_PATH, 'wb') as token:
                pickle.dump(self.creds, token)
        
        self.service = build('gmail', 'v1', credentials=self.creds)
        print("Gmail authentication successful!")
    
    def get_unread_emails(self, last_processed_date=None):
        """Fetch unread emails since last processed date"""
        query = 'is:unread in:inbox'
        if last_processed_date:
            query += f' after:{last_processed_date.strftime("%Y/%m/%d")}'
        
        try:
            results = self.service.users().messages().list(
                userId='me',
                q=query,
                maxResults=50
            ).execute()
            
            messages = results.get('messages', [])
            print(f"Found {len(messages)} unread messages")
            return messages
            
        except Exception as e:
            print(f"Error fetching emails: {e}")
            return []
    
    def get_email_details(self, msg_id):
        """Get full email details"""
        try:
            message = self.service.users().messages().get(
                userId='me',
                id=msg_id,
                format='full'
            ).execute()
            return message
        except Exception as e:
            print(f"Error fetching email {msg_id}: {e}")
            return None
    
    def mark_as_read(self, msg_id):
        """Mark email as read"""
        try:
            self.service.users().messages().modify(
                userId='me',
                id=msg_id,
                body={'removeLabelIds': ['UNREAD']}
            ).execute()
            print(f"✓ Marked email {msg_id} as read")
            return True
        except Exception as e:
            print(f"Error marking email as read: {e}")
            return False