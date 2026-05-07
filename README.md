# LeadHunter Pro — Setup Guide
=====================================

## FILES IN THIS FOLDER
```
C:\LeadScraper\
├── scraper_app.py        ← Main app (run this)
├── sheets_service.py     ← Secure Google Sheets backend
├── credentials.json      ← YOUR service account key (private!)
├── .env                  ← Your environment config
└── .env.example          ← Template for .env
```

## STEP 1 — Setup credentials (one time only)

1. Copy your downloaded JSON file to `C:\LeadScraper\`
2. Rename it to `credentials.json`
3. Copy `.env.example` → rename to `.env`
4. Open `.env` and set:
   ```
   GOOGLE_CREDENTIALS_PATH=C:/LeadScraper/credentials.json
   ```

## STEP 2 — Install libraries (one time only)

```bash
pip install gspread google-auth google-auth-oauthlib google-auth-httplib2 python-dotenv
```

## STEP 3 — Run the app

```bash
cd C:\LeadScraper
python scraper_app.py
```

## HOW USERS CONNECT THEIR SHEET

Users do NOT upload any credentials. They just:

1. Create a Google Sheet at sheets.google.com
2. Click Share → Add this email as Editor:
   `leadscraper@leadscraper-493409.iam.gserviceaccount.com`
3. Paste their Sheet URL into the app
4. Click "Test Connection"
5. Done! All leads sync automatically.

## SECURITY MODEL

| What                  | Who has access |
|-----------------------|----------------|
| credentials.json      | You only       |
| Service account email | Public (safe)  |
| Sheet URL             | User only      |
| Sheet data            | User only      |

✅ Users NEVER see credentials
✅ Each user connects their OWN sheet
✅ Data is never mixed between users
