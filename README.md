# Gmail to Google Sheets Automation

## Architecture


## Setup Instructions

### 1. Prerequisites
- Python 3.8+
- Google Account
- VS Code (recommended)

### 2. Google Cloud Setup
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create new project
3. Enable Gmail API and Google Sheets API
4. Create OAuth 2.0 Desktop App credentials
5. Download `credentials.json` to `credentials/` folder

### 3. Local Setup
```bash
# Clone repository
git clone <your-repo-url>
cd gmail-to-sheets

# Create virtual environment
python -m venv venv

# Activate (Windows)
venv\Scripts\activate
# Activate (Mac/Linux)
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt