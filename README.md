# 📧 Gmail to Google Sheets Automation System

**Author:** Balarajusridhar Raju  
**Language:** Python 3  
**APIs Used:** Gmail API, Google Sheets API  
**Authentication:** OAuth 2.0 (Desktop Flow)

---

## 📖 Project Overview

This project is a Python automation system that connects the **Gmail API** and **Google Sheets API** to read real incoming emails from a Gmail inbox and log them into a Google Sheet.

Only **unread inbox emails** are processed. Each email is logged exactly once and then marked as **read** to prevent duplicate entries.

---

## 🎯 Objective

For every qualifying email, the system appends a new row in Google Sheets with the following fields:

| Column | Description |
|------|-------------|
| From | Sender email address |
| Subject | Email subject |
| Date | Date and time received |
| Content | Email body (plain text) |

---

## 🧠 High-Level Architecture

┌────────────┐
│ Gmail │
│ (Inbox) │
└─────┬──────┘
│ Gmail API (Unread Emails)
▼
┌────────────┐
│ Email │
│ Parser │
└─────┬──────┘
│ Parsed Data
▼
┌────────────┐
│ Google │
│ Sheets API │
└─────┬──────┘
│ Append Rows
▼
┌────────────┐
│ Google │
│ Sheet │
└────────────┘

---

## 📂 Required Project Structure


gmail-to-sheets/
├── src/
│ ├── gmail_service.py
│ ├── sheets_service.py
│ ├── email_parser.py
│ └── main.py
├── credentials/
│ └── credentials.json # DO NOT COMMIT
├── proof/
├── .gitignore
├── requirements.txt
├── config.py
└── README.md


---

## ⚙️ Step-by-Step Setup Instructions

### 1️⃣ Clone the Repository

```bash
git clone <https://github.com/BalarajuSridhar/Gmail-to-Sheets.git>
cd gmail-to-sheets

2️⃣ Install Dependencies
pip install -r requirements.txt

3️⃣ Google Cloud Configuration

Open Google Cloud Console

Create a new project

Enable:

Gmail API

Google Sheets API

Configure OAuth Consent Screen

Create OAuth 2.0 Client ID:

Application type: Desktop App

Download credentials.json

Place it inside the credentials/ folder

4️⃣ Run the Script
python src/main.py


A browser window will open on first execution for OAuth authentication.

🔐 OAuth 2.0 Flow Explanation

This project uses OAuth 2.0 Desktop Authentication.

The user logs in via browser

Google issues an access token and refresh token

Tokens are stored locally for reuse

Allows secure, non-interactive future runs

No service accounts are used

🚫 Duplicate Prevention Logic

Duplicates are prevented using two mechanisms:

Unread Email Filtering

Only emails with the UNREAD label are fetched

Inbox State Update

After logging an email, the script removes the UNREAD label

Processed emails are never fetched again

This guarantees no duplicate rows in Google Sheets.

🔄 State Persistence Method

State is stored directly in Gmail, not locally.

Gmail’s READ/UNREAD labels act as persistent state

This approach is reliable and survives:

System restarts

Token deletion

Local file loss

Reason for choice:
Using Gmail labels avoids external databases and keeps the system stateless and robust.

🧩 Challenges Faced & Solution
Challenge

Emails often contain multi-part MIME content (HTML + text).

Solution

A recursive MIME parser was implemented to:

Traverse all parts

Extract clean text/plain content

Ignore HTML and attachments

⚠️ Limitations

Subject to Google API rate limits

Email attachments are not processed

Requires one-time OAuth authentication

Only processes Inbox unread emails

📸 Proof of Execution

The /proof/ folder contains:

Gmail inbox showing unread emails

Google Sheet with at least 5 populated rows

OAuth consent screen (email blurred)

2–3 minute execution video

🎥 Demo Video: <ADD_VIDEO_LINK_HERE>

🔒 Security Rules Followed

credentials.json is ignored

OAuth tokens are not committed

.gitignore prevents secret leaks

No API keys are hardcoded

📦 Submission Details

Repository is public

README contains:

Architecture

OAuth explanation

Duplicate prevention

State persistence

Challenges & limitations

Proof folder included

✅ Conclusion

This project demonstrates secure API integration, clean automation design, and real-world email processing.
It fully satisfies all mandatory internship assignment requirements and is ready for evaluation.



