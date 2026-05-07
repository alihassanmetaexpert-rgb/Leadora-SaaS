"""
LeadHunter Pro — Google Sheets Service (OAuth Edition)
=======================================================

ARCHITECTURE:
  OLD WAY (bad):
    User must manually share sheet with service account email
    Every user has to do 5 steps before connecting

  NEW WAY (this file):
    User just pastes their sheet URL
    App opens Google login in browser (one click)
    User approves access once
    Done — leads sync automatically forever

HOW IT WORKS:
  1. User pastes their Google Sheet URL
  2. App opens browser → Google login page
  3. User clicks "Allow" 
  4. Token saved locally (never expires unless user revokes)
  5. App writes leads directly to their sheet

SETUP (one time, by you):
  1. Go to console.cloud.google.com
  2. APIs & Services → Credentials
  3. Create OAuth 2.0 Client ID
  4. Application type: Desktop App
  5. Download the client_secret.json file
  6. Put it in C:/LeadScraper/client_secret.json
  7. Set in .env: OAUTH_CLIENT_SECRET_PATH=C:/LeadScraper/client_secret.json

USERS:
  - Just paste their Google Sheet URL
  - Click Connect
  - Browser opens → they log in with Google → click Allow
  - Done forever
"""

import os
import re
import json
import time
import logging
import threading
import webbrowser
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger("LeadHunterSheets")

# ── Load env ──────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR = Path(os.environ.get(
    "LEADSCRAPER_DIR",
    os.path.join(os.path.expanduser("~"), "Desktop", "LeadResults")
))
BASE_DIR.mkdir(parents=True, exist_ok=True)

# OAuth client secret — you create this once in Google Cloud Console
OAUTH_CLIENT_SECRET = os.environ.get(
    "OAUTH_CLIENT_SECRET_PATH",
    str(Path(__file__).parent / "client_secret.json")
)

# Where user tokens are stored per user
TOKENS_DIR = BASE_DIR / "user_tokens"
TOKENS_DIR.mkdir(exist_ok=True)

# Google API scopes needed
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

# Sheet headers
SHEET_HEADERS = [
    "Name", "Category", "City", "Address",
    "Phone", "Email", "Website", "Rating",
    "Source", "Maps URL", "Date Added"
]

SERVICE_ACCOUNT_EMAIL = "leadscraper@leadscraper-493409.iam.gserviceaccount.com"


# ── Custom Exceptions ─────────────────────────────────────────────────────────

class InvalidSheetURLError(Exception):
    pass

class NotAuthenticatedError(Exception):
    pass

class SheetNotSharedError(Exception):
    pass

class SheetsAPIError(Exception):
    pass


# ── URL Extraction ────────────────────────────────────────────────────────────

def extract_spreadsheet_id(sheet_url: str) -> str:
    """
    Extract spreadsheet ID from any Google Sheets URL format.
    Accepts all formats users might paste.
    """
    if not sheet_url or not sheet_url.strip():
        raise InvalidSheetURLError("Sheet URL cannot be empty.")

    url = sheet_url.strip().strip('"').strip("'")

    # Raw spreadsheet ID pasted directly
    if re.match(r'^[a-zA-Z0-9_-]{20,60}$', url):
        return url

    # Standard URL: extract ID after /d/
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if match:
        return match.group(1)

    # Fallback: any long alphanumeric segment
    segments = re.findall(r'[a-zA-Z0-9_-]{25,}', url)
    if segments:
        return segments[0]

    raise InvalidSheetURLError(
        "Could not find spreadsheet ID in URL.\n"
        "Please paste the full URL from your browser address bar.\n"
        "Example: https://docs.google.com/spreadsheets/d/YOUR_ID/edit"
    )


# ── OAuth Authentication ──────────────────────────────────────────────────────

def get_token_path(user_id: str) -> Path:
    """Get path to token file for a specific user."""
    safe_id = re.sub(r'[^a-zA-Z0-9_-]', '_', user_id)
    return TOKENS_DIR / f"token_{safe_id}.json"


def is_user_authenticated(user_id: str) -> bool:
    """Check if user has a valid saved token."""
    token_path = get_token_path(user_id)
    if not token_path.exists():
        return False
    try:
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        return creds and (creds.valid or creds.refresh_token)
    except Exception:
        return False


def authenticate_user(user_id: str, log_fn=None) -> bool:
    """
    Authenticate a user via OAuth browser flow.

    Opens browser → user logs into Google → user clicks Allow
    Token saved locally for future use.

    Args:
        user_id: Unique identifier for this user/session
        log_fn:  Optional logging callback

    Returns:
        True if authenticated successfully
    """
    def log(msg):
        logger.info(msg)
        if log_fn:
            log_fn(msg)

    # Check if already authenticated
    if is_user_authenticated(user_id):
        log("✅ Already authenticated with Google")
        return True

    # Check OAuth client secret exists
    if not os.path.exists(OAUTH_CLIENT_SECRET):
        log(f"❌ OAuth client secret not found: {OAUTH_CLIENT_SECRET}")
        log("   Please set up OAuth credentials in Google Cloud Console")
        return False

    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials

        log("🔐 Opening Google login in browser...")
        log("   Please sign in and click Allow to grant access")

        flow = InstalledAppFlow.from_client_secrets_file(
            OAUTH_CLIENT_SECRET, SCOPES)

        # Run local server to capture OAuth callback
        creds = flow.run_local_server(
            port=0,
            prompt='consent',
            access_type='offline'
        )

        # Save token for this user
        token_path = get_token_path(user_id)
        with open(str(token_path), 'w') as f:
            f.write(creds.to_json())

        log(f"✅ Google account connected successfully!")
        return True

    except Exception as e:
        log(f"❌ Authentication failed: {e}")
        return False


def get_authenticated_client(user_id: str):
    """
    Get authenticated gspread client for a specific user.
    Refreshes token if expired.
    """
    try:
        import gspread
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request

        token_path = get_token_path(user_id)

        if not token_path.exists():
            raise NotAuthenticatedError(
                "User not authenticated. Please connect Google account first.")

        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

        # Refresh token if expired
        if not creds.valid:
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                # Save refreshed token
                with open(str(token_path), 'w') as f:
                    f.write(creds.to_json())
            else:
                raise NotAuthenticatedError(
                    "Token expired. Please reconnect Google account.")

        client = gspread.authorize(creds)
        return client

    except ImportError:
        raise ImportError(
            "Required libraries not installed.\n"
            "Run: pip install gspread google-auth google-auth-oauthlib"
        )


def revoke_user_token(user_id: str) -> bool:
    """Remove saved token for a user (disconnect)."""
    token_path = get_token_path(user_id)
    if token_path.exists():
        token_path.unlink()
        logger.info(f"Token revoked for user {user_id}")
        return True
    return False


# ── Sheet Operations ──────────────────────────────────────────────────────────

def _get_or_create_worksheet(spreadsheet, sheet_name: str):
    """Get existing tab or create it with headers."""
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except Exception:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=5000, cols=15)

    # Ensure headers exist
    try:
        first_row = ws.row_values(1)
        if not first_row or first_row[0] != "Name":
            ws.insert_row(SHEET_HEADERS, index=1)
            ws.format("A1:K1", {
                "textFormat": {"bold": True, "fontSize": 11,
                               "foregroundColor": {"red":1,"green":1,"blue":1}},
                "backgroundColor": {"red":0.102,"green":0.451,"blue":0.910},
                "horizontalAlignment": "CENTER"
            })
    except Exception as e:
        logger.warning(f"Could not format headers: {e}")

    return ws


def _get_existing_keys(worksheet) -> set:
    """Get (name, phone) pairs already in sheet to avoid duplicates."""
    try:
        all_rows = worksheet.get_all_values()
        return {(r[0].strip().lower(), r[4].strip())
                for r in all_rows[1:] if len(r) >= 5}
    except Exception:
        return set()


def _leads_to_rows(leads: List[Dict], existing_keys: set) -> List[List]:
    """Convert leads to sheet rows, skipping duplicates."""
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")
    rows = []
    for lead in leads:
        name  = str(lead.get("name","")).strip()
        phone = str(lead.get("phone","")).strip()
        key   = (name.lower(), phone)
        if key in existing_keys:
            continue
        existing_keys.add(key)
        rows.append([
            name,
            str(lead.get("category","")),
            str(lead.get("city","")),
            str(lead.get("address","")),
            phone,
            str(lead.get("email","")),
            str(lead.get("website","")),
            str(lead.get("rating","")),
            str(lead.get("source","Google Maps")),
            str(lead.get("maps_url","")),
            now,
        ])
    return rows


# ── Main Public Functions ─────────────────────────────────────────────────────

def test_sheet_connection(
    sheet_url: str,
    user_id:   str,
    sheet_name: str = "Leads",
    log_fn = None
) -> Dict:
    """
    Test if user's sheet is accessible.
    User must be authenticated via OAuth first.

    Args:
        sheet_url:  User's Google Sheet URL
        user_id:    Unique user identifier
        sheet_name: Tab name
        log_fn:     Optional log callback

    Returns:
        Dict: {success, message, sheet_title}
    """
    def log(msg):
        logger.info(msg)
        if log_fn: log_fn(msg)

    try:
        log("   Extracting sheet ID from URL...")
        spreadsheet_id = extract_spreadsheet_id(sheet_url)

        log("   Getting authenticated client...")
        client = get_authenticated_client(user_id)

        log("   Opening spreadsheet...")
        try:
            spreadsheet = client.open_by_key(spreadsheet_id)
        except Exception as e:
            err = str(e).lower()
            if "forbidden" in err or "403" in err:
                return {
                    "success": False,
                    "message": "forbidden: You don't have access to this sheet. Make sure you're opening YOUR sheet."
                }
            if "404" in err or "not found" in err:
                return {
                    "success": False,
                    "message": "Sheet not found. Please check the URL is correct."
                }
            return {"success": False, "message": f"Could not open sheet: {e}"}

        log("   Writing test row...")
        try:
            ws = _get_or_create_worksheet(spreadsheet, sheet_name)
            ws.append_row([
                "✅ LeadHunter Connected!", "Test", "Test City",
                "Test Address", "000-0000", "test@test.com",
                "www.test.com", "5.0", "Connection Test", "",
                datetime.now().strftime("%Y-%m-%d %H:%M")
            ], value_input_option="RAW")
        except Exception as e:
            return {"success": False, "message": f"Could not write to sheet: {e}"}

        log(f"   ✅ Connected to: {spreadsheet.title}")
        return {
            "success": True,
            "message": f"Connected to '{spreadsheet.title}'",
            "sheet_title": spreadsheet.title,
        }

    except NotAuthenticatedError:
        return {
            "success": False,
            "message": "not_authenticated: Please connect your Google account first"
        }
    except InvalidSheetURLError as e:
        return {"success": False, "message": f"Invalid URL: {e}"}
    except ImportError as e:
        return {"success": False, "message": f"gspread not installed: {e}"}
    except Exception as e:
        return {"success": False, "message": f"Error: {e}"}


def append_leads_to_sheet(
    sheet_url:  str,
    leads:      List[Dict],
    user_id:    str,
    sheet_name: str = "Leads",
    log_fn = None,
) -> Dict:
    """
    Append leads to user's Google Sheet.
    Uses OAuth — user's own Google account writes to their own sheet.

    Args:
        sheet_url:  User's Google Sheet URL
        leads:      List of lead dicts
        user_id:    Unique user identifier
        sheet_name: Tab name
        log_fn:     Optional log callback

    Returns:
        Dict: {success, added, skipped, message}
    """
    def log(msg):
        logger.info(msg)
        if log_fn: log_fn(msg)

    if not leads:
        return {"success": True, "added": 0, "skipped": 0,
                "message": "No leads to append"}

    try:
        spreadsheet_id = extract_spreadsheet_id(sheet_url)
        client         = get_authenticated_client(user_id)

        try:
            spreadsheet = client.open_by_key(spreadsheet_id)
        except Exception as e:
            err = str(e).lower()
            if "forbidden" in err or "403" in err:
                return {
                    "success": False, "added": 0, "skipped": 0,
                    "message": "Cannot access sheet. Please reconnect your Google account."
                }
            return {
                "success": False, "added": 0, "skipped": 0,
                "message": f"Could not open sheet: {e}"
            }

        ws            = _get_or_create_worksheet(spreadsheet, sheet_name)
        existing_keys = _get_existing_keys(ws)
        rows          = _leads_to_rows(leads, existing_keys)
        skipped       = len(leads) - len(rows)

        if not rows:
            return {
                "success": True, "added": 0, "skipped": skipped,
                "message": f"All {len(leads)} leads already in sheet"
            }

        # Append in batches
        added = 0
        for i in range(0, len(rows), 50):
            batch = rows[i:i+50]
            ws.append_rows(batch, value_input_option="RAW")
            added += len(batch)
            log(f"   Appended {added}/{len(rows)} rows...")
            time.sleep(1)

        msg = (f"✅ {added} leads added to '{spreadsheet.title}'"
               + (f" ({skipped} duplicates skipped)" if skipped else ""))
        log(f"   {msg}")
        return {"success": True, "added": added, "skipped": skipped, "message": msg}

    except NotAuthenticatedError:
        return {
            "success": False, "added": 0, "skipped": 0,
            "message": "not_authenticated: Please connect Google account first"
        }
    except Exception as e:
        return {"success": False, "added": 0, "skipped": 0,
                "message": f"Error: {e}"}


def install_gspread(log_fn=None):
    """Install required Google libraries."""
    import subprocess
    def log(msg):
        if log_fn: log_fn(msg)
        else: print(msg)
    log("Installing Google libraries...")
    try:
        subprocess.check_call([
            "pip", "install", "gspread", "google-auth",
            "google-auth-oauthlib", "google-auth-httplib2",
            "python-dotenv", "--quiet"
        ])
        log("✅ Libraries installed successfully!")
        return True
    except Exception as e:
        log(f"❌ Install failed: {e}")
        return False
