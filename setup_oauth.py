#!/usr/bin/env python3
"""
OAuth Setup for Google Sheets - Manual Flow
"""

from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
import gspread
import pickle
import os

CREDENTIALS_FILE = '/home/filepro/credentials/client_secret_2705580970-ttgvft35bj2d9mte3hoqcpk7h9p2t4rm.apps.googleusercontent.com.json'
TOKEN_FILE = '/home/filepro/credentials/token.pickle'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = None

# Load existing token if available
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE, 'rb') as token:
        creds = pickle.load(token)

# If no valid credentials, do OAuth flow
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = Flow.from_client_secrets_file(
            CREDENTIALS_FILE,
            scopes=SCOPES,
            redirect_uri='urn:ietf:wg:oauth:2.0:oob'
        )

        auth_url, _ = flow.authorization_url(prompt='consent')

        print("\n" + "="*60)
        print("Open this URL in your browser:")
        print("="*60)
        print(auth_url)
        print("="*60)
        print("\nAfter authorizing, copy the code and paste it below.\n")

        code = input("Enter authorization code: ")
        flow.fetch_token(code=code)
        creds = flow.credentials

    # Save credentials for next run
    with open(TOKEN_FILE, 'wb') as token:
        pickle.dump(creds, token)

# Authorize gspread
gc = gspread.authorize(creds)

print("\nAuthentication successful!")
print("Testing by creating a sheet...")

sh = gc.create("Test_FilePro_Sheet")
print(f"Created: {sh.url}")
print("\nOAuth setup complete! Token saved for future use.")
