#!/usr/bin/env python3
"""Create the overseas intake Google Sheet after the user authorizes.

Never print tokens, client secrets, or authorization codes.
Credentials come from flags or environment variables, not from this file.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from urllib.parse import urlencode

SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
)
REDIRECT_URI = "http://localhost"
TOKEN_URI = "https://oauth2.googleapis.com/token"
AUTH_URI = "https://accounts.google.com/o/oauth2/v2/auth"
SHEET_TITLE = "Overseas Intake — sync to Feishu"
SHARE_WITH = os.environ.get("GOOGLE_SHARE_WITH", "liwenliang@laoyuegou.com")

ROOT = Path(__file__).resolve().parent
XLSX_PATH = ROOT / "overseas-intake.xlsx"
TOKEN_PATH = ROOT.parent / ".google-tokens.json"


def _client_id(args: argparse.Namespace) -> str:
    value = args.client_id or os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "")
    if not value:
        raise SystemExit("Missing GOOGLE_OAUTH_CLIENT_ID / --client-id")
    return value


def _client_secret(args: argparse.Namespace) -> str:
    value = args.client_secret or os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "")
    if not value:
        raise SystemExit("Missing GOOGLE_OAUTH_CLIENT_SECRET / --client-secret")
    return value


def print_auth_url(args: argparse.Namespace) -> None:
    params = {
        "client_id": _client_id(args),
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": " ".join(SCOPES),
        "access_type": "offline",
        "prompt": "consent",
        "include_granted_scopes": "true",
    }
    print(f"{AUTH_URI}?{urlencode(params)}")


def exchange_code(args: argparse.Namespace) -> None:
    import requests

    code = args.code or os.environ.get("GOOGLE_OAUTH_CODE", "")
    if not code:
        raise SystemExit("Missing --code (the value after code= on the localhost URL)")

    response = requests.post(
        TOKEN_URI,
        data={
            "code": code,
            "client_id": _client_id(args),
            "client_secret": _client_secret(args),
            "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code",
        },
        timeout=30,
    )
    if not response.ok:
        raise SystemExit(f"Token exchange failed: HTTP {response.status_code}")
    payload = response.json()
    if "refresh_token" not in payload or "access_token" not in payload:
        raise SystemExit("Token exchange succeeded but refresh_token was missing. Re-consent with prompt=consent.")
    TOKEN_PATH.write_text(json.dumps(payload), encoding="utf-8")
    TOKEN_PATH.chmod(0o600)
    print("Authorization saved. Tokens were not printed.")


def _creds_from_oauth(args: argparse.Namespace):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    raw = None
    if TOKEN_PATH.exists():
        raw = json.loads(TOKEN_PATH.read_text(encoding="utf-8"))
    elif os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN"):
        raw = {
            "refresh_token": os.environ["GOOGLE_OAUTH_REFRESH_TOKEN"],
            "token": os.environ.get("GOOGLE_OAUTH_ACCESS_TOKEN", ""),
            "token_uri": TOKEN_URI,
        }
    if not raw:
        raise SystemExit("No OAuth tokens. Run auth-url → exchange first.")

    creds = Credentials(
        token=raw.get("access_token") or raw.get("token") or None,
        refresh_token=raw.get("refresh_token"),
        token_uri=TOKEN_URI,
        client_id=_client_id(args),
        client_secret=_client_secret(args),
        scopes=list(SCOPES),
    )
    if not creds.valid:
        creds.refresh(Request())
        merged = dict(raw)
        merged["access_token"] = creds.token
        TOKEN_PATH.write_text(json.dumps(merged), encoding="utf-8")
        TOKEN_PATH.chmod(0o600)
    return creds


def _creds_from_service_account():
    from google.oauth2 import service_account

    path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")
    blob = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if path:
        return service_account.Credentials.from_service_account_file(path, scopes=list(SCOPES))
    if blob:
        info = json.loads(blob)
        return service_account.Credentials.from_service_account_info(info, scopes=list(SCOPES))
    return None


def _drive_service(args: argparse.Namespace):
    from googleapiclient.discovery import build

    creds = _creds_from_service_account()
    if creds is None:
        creds = _creds_from_oauth(args)
    return build("drive", "v3", credentials=creds), creds


def publish(args: argparse.Namespace) -> None:
    from googleapiclient.http import MediaFileUpload

    if not XLSX_PATH.exists():
        raise SystemExit(f"Missing {XLSX_PATH}. Run build_intake_sheet.py first.")

    drive, creds = _drive_service(args)
    media = MediaFileUpload(
        str(XLSX_PATH),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    created = (
        drive.files()
        .create(
            body={"name": SHEET_TITLE, "mimeType": "application/vnd.google-apps.spreadsheet"},
            media_body=media,
            fields="id,webViewLink,owners",
        )
        .execute()
    )
    file_id = created["id"]
    link = created.get("webViewLink") or f"https://docs.google.com/spreadsheets/d/{file_id}"

    share_with = args.share_with or SHARE_WITH
    is_service_account = getattr(creds, "service_account_email", None)
    if is_service_account and share_with:
        drive.permissions().create(
            fileId=file_id,
            body={"type": "user", "role": "writer", "emailAddress": share_with},
            sendNotificationEmail=True,
            fields="id",
        ).execute()

    print(f"Created: {link}")
    print(f"Spreadsheet ID: {file_id}")


def main() -> None:
    auth = argparse.ArgumentParser(add_help=False)
    auth.add_argument("--client-id", default="")
    auth.add_argument("--client-secret", default="")

    parser = argparse.ArgumentParser(description="Authorize and publish the overseas intake sheet.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("auth-url", help="Print the Google consent URL", parents=[auth])

    exchange = sub.add_parser("exchange", help="Swap the one-time code for tokens", parents=[auth])
    exchange.add_argument("--code", default="")

    create = sub.add_parser("create", help="Upload the xlsx as a Google Sheet", parents=[auth])
    create.add_argument("--share-with", default="")

    args = parser.parse_args()
    commands = {
        "auth-url": print_auth_url,
        "exchange": exchange_code,
        "create": publish,
    }
    commands[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
