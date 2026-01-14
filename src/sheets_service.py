from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os

class SheetsService:
    def __init__(self, credentials):
        self.creds = credentials
        self.service = build('sheets', 'v4', credentials=self.creds)
    
    def append_row(self, spreadsheet_id, sheet_name, row_data):
        """Append a row to Google Sheet"""
        try:
            # Prepare the row in correct column order
            values = [[
                row_data['from'],
                row_data['subject'],
                row_data['date'],
                row_data['content'][:50000]  # Sheets cell limit
            ]]
            
            body = {
                'values': values
            }
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:D",
                valueInputOption="USER_ENTERED",
                body=body,
                insertDataOption="INSERT_ROWS"
            ).execute()
            
            print(f"Appended row. Updated cells: {result.get('updates', {}).get('updatedCells')}")
            return True
            
        except Exception as e:
            print(f"Error appending to sheet: {e}")
            return False
    
    def get_existing_message_ids(self, spreadsheet_id, sheet_name):
        """Get all already processed message IDs from sheet"""
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!E:E"
            ).execute()
            
            values = result.get('values', [])
            return [item for sublist in values[1:] for item in sublist if item]
        except:
            return []