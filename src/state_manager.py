import json
import os
from datetime import datetime
from pathlib import Path

class StateManager:
    def __init__(self):
        BASE_DIR = Path(__file__).parent.parent
        self.state_file = BASE_DIR / "data" / "last_processed.json"
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
        
    def load_state(self):
        """Load last processed timestamp"""
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, 'r') as f:
                    state = json.load(f)
                    last_processed_str = state.get('last_processed')
                    if last_processed_str:
                        return datetime.fromisoformat(last_processed_str)
            except:
                pass
        return None
    
    def save_state(self, timestamp=None):
        """Save current timestamp as last processed"""
        if timestamp is None:
            timestamp = datetime.now()
        
        state = {
            'last_processed': timestamp.isoformat(),
            'updated_at': datetime.now().isoformat()
        }
        
        with open(self.state_file, 'w') as f:
            json.dump(state, f, indent=2)
        
        print(f"State saved. Last processed: {timestamp}")
    
    def update_last_processed(self, email_date):
        """Update state with the latest email date"""
        try:
            email_dt = datetime.fromisoformat(email_date.replace('Z', '+00:00'))
            self.save_state(email_dt)
        except:
            self.save_state()