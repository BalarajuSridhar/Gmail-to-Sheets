# sheets_service.py
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
                row_data['content'][:50000],  # Sheets cell limit
                row_data.get('message_id', '')  # Store message_id in column E
            ]]
            
            body = {
                'values': values
            }
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:E",  # Changed to A:E to include column E
                valueInputOption="USER_ENTERED",
                body=body,
                insertDataOption="INSERT_ROWS"
            ).execute()
            
            print(f"✓ Appended row. Updated cells: {result.get('updates', {}).get('updatedCells')}")
            return True
            
        except Exception as e:
            print(f"✗ Error appending to sheet: {e}")
            return False
    
    def get_existing_message_ids(self, spreadsheet_id, sheet_name):
        """Get all already processed message IDs from sheet"""
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!E:E"
            ).execute()
            
            values = result.get('values', [])
            # Flatten list and skip header
            existing_ids = []
            for i, row in enumerate(values):
                if i == 0:  # Skip header row
                    continue
                if row and row[0]:
                    existing_ids.append(row[0])
            return existing_ids
        except Exception as e:
            print(f"Note: Could not retrieve existing message IDs: {e}")
            return []
    
    def format_sheet(self, spreadsheet_id, sheet_name):
        """Format the Google Sheet with proper headers and styling"""
        try:
            # First, get sheet ID by name
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            
            sheet_id = None
            for sheet in spreadsheet['sheets']:
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                print(f"Sheet '{sheet_name}' not found")
                return False
            
            # 1. Set headers
            headers = [['From', 'Subject', 'Date', 'Content', 'Message_ID']]
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1:E1",
                valueInputOption="USER_ENTERED",
                body={'values': headers}
            ).execute()
            
            # 2. Apply formatting
            requests = [
                # Format header row
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 5
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                                "textFormat": {"bold": True},
                                "horizontalAlignment": "CENTER"
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                    }
                },
                # Set column widths
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1
                        },
                        "properties": {"pixelSize": 200},
                        "fields": "pixelSize"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 1,
                            "endIndex": 2
                        },
                        "properties": {"pixelSize": 300},
                        "fields": "pixelSize"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,
                            "endIndex": 3
                        },
                        "properties": {"pixelSize": 150},
                        "fields": "pixelSize"
                    }
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 3,
                            "endIndex": 4
                        },
                        "properties": {"pixelSize": 400},
                        "fields": "pixelSize"
                    }
                },
                # Hide Message_ID column (E)
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": 4,
                            "endIndex": 5
                        },
                        "properties": {"hiddenByUser": True},
                        "fields": "hiddenByUser"
                    }
                },
                # Freeze header row
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "frozenRowCount": 1
                            }
                        },
                        "fields": "gridProperties.frozenRowCount"
                    }
                }
            ]
            
            # Apply all formatting requests
            body = {"requests": requests}
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            print(f"✅ Sheet '{sheet_name}' formatted successfully!")
            print("   - Headers: From | Subject | Date | Content | Message_ID")
            print("   - Column E (Message_ID) is hidden")
            print("   - Header row is frozen")
            return True
            
        except Exception as e:
            print(f"✗ Error formatting sheet: {e}")
            return False
    
    def create_or_reset_sheet(self, spreadsheet_id, sheet_name):
        """Create or reset a sheet with proper formatting"""
        try:
            # First, try to format existing sheet
            success = self.format_sheet(spreadsheet_id, sheet_name)
            
            if not success:
                print("Creating new sheet...")
                # If sheet doesn't exist, create it
                requests = [{
                    "addSheet": {
                        "properties": {
                            "title": sheet_name
                        }
                    }
                }]
                
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": requests}
                ).execute()
                
                # Now format it
                self.format_sheet(spreadsheet_id, sheet_name)
            
            return True
            
        except Exception as e:
            print(f"Error creating/resetting sheet: {e}")
            return False