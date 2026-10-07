"""Google Sheets access via OAuth (read-only scope) using gspread."""
import json
import os
import re
from pathlib import Path

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
BASE = Path(__file__).resolve().parent.parent
CRED_PATH = BASE / "credentials.json"
TOKEN_PATH = BASE / "tokens.json"
REDIRECT_URI = os.environ.get(
    "BS_REDIRECT_URI", "http://localhost:8000/oauth/callback"
)


class AuthError(RuntimeError):
    pass


def credentials_ready():
    return CRED_PATH.exists()


def _flow():
    if not credentials_ready():
        raise AuthError("credentials.json chưa được đặt vào thư mục tool")
    flow = Flow.from_client_secrets_file(str(CRED_PATH), SCOPES)
    flow.redirect_uri = REDIRECT_URI
    return flow


def _load_credentials():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_info(
            json.loads(TOKEN_PATH.read_text()), SCOPES
        )
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_PATH.write_text(json.dumps(creds.to_json_dict(), indent=2))
    return creds if creds and creds.valid else None


def client():
    creds = _load_credentials()
    if not creds:
        raise AuthError("Chưa đăng nhập Google (hoặc token hết hạn, đăng nhập lại)")
    return gspread.authorize(creds)


def list_sheets(spreadsheet_id):
    return [ws.title for ws in client().open_by_key(spreadsheet_id).worksheets()]


def read_sheet(spreadsheet_id, title):
    return client().open_by_key(spreadsheet_id).worksheet(title).get_all_values()


def start_login():
    url, _ = _flow().authorization_url(
        access_type="offline", prompt="consent"
    )
    return url


def finish_login(code):
    flow = _flow()
    flow.fetch_token(code=code)
    TOKEN_PATH.write_text(json.dumps(flow.credentials.to_json_dict(), indent=2))


def clear_token():
    if TOKEN_PATH.exists():
        TOKEN_PATH.unlink()


def extract_spreadsheet_id(url_or_id):
    """Accept a full Google Sheets URL or a bare spreadsheet id."""
    s = str(url_or_id).strip()
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", s)
    return m.group(1) if m else s
