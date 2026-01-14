import base64
import email
from email.header import decode_header
import re
from bs4 import BeautifulSoup
from datetime import datetime
import pytz

class EmailParser:
    @staticmethod
    def decode_subject(encoded_subject):
        """Decode email subject"""
        if encoded_subject is None:
            return ""
        
        decoded_parts = decode_header(encoded_subject)
        subject = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                subject += part.decode(encoding or 'utf-8', errors='replace')
            else:
                subject += part
        return subject
    
    @staticmethod
    def get_body(message):
        """Extract plain text body from email"""
        if 'parts' in message['payload']:
            for part in message['payload']['parts']:
                if part['mimeType'] == 'text/plain':
                    data = part['body'].get('data', '')
                    if data:
                        return base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
                elif part['mimeType'] == 'multipart/alternative':
                    for subpart in part['parts']:
                        if subpart['mimeType'] == 'text/plain':
                            data = subpart['body'].get('data', '')
                            if data:
                                return base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
        
        # Fallback to HTML or snippet
        if 'body' in message['payload']:
            data = message['payload']['body'].get('data', '')
            if data:
                html_content = base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
                # Convert HTML to plain text
                soup = BeautifulSoup(html_content, 'html.parser')
                return soup.get_text()[:1000]  # Limit to 1000 chars
        
        return message.get('snippet', '')[:500]
    
    @staticmethod
    def parse_email(message_details):
        """Parse email details into structured format"""
        headers = {header['name']: header['value'] 
                  for header in message_details['payload']['headers']}
        
        # Parse date
        date_str = headers.get('Date', '')
        try:
            # Try to parse date
            date_obj = email.utils.parsedate_to_datetime(date_str)
        except:
            date_obj = datetime.now(pytz.UTC)
        
        # Parse sender
        from_header = headers.get('From', '')
        sender_match = re.search(r'<(.+?)>', from_header)
        if sender_match:
            sender_email = sender_match.group(1)
        else:
            sender_email = from_header
        
        return {
            'from': sender_email,
            'subject': EmailParser.decode_subject(headers.get('Subject', '')),
            'date': date_obj.isoformat(),
            'content': EmailParser.get_body(message_details),
            'message_id': message_details['id'],
            'thread_id': message_details['threadId']
        }