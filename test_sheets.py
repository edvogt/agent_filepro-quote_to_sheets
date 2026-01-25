#!/usr/bin/env python3
"""Quick test to diagnose Google Sheets API issues"""

import gspread
from google.oauth2.service_account import Credentials

SERVICE_ACCOUNT = '/home/filepro/credentials/mystical-timing-483316-b7-a0778db44c09.json'
FOLDER_ID = '1jnWNqdSBy8aigv2c_MclEtv4r-HJvsmB'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT, scopes=scopes)
client = gspread.authorize(creds)

print("Testing Google Sheets API...")

# Test 1: List files in folder
print("\n1. Listing files in shared folder...")
try:
    files = client.list_spreadsheet_files(folder_id=FOLDER_ID)
    print(f"   Found {len(files)} files")
    for f in files[:5]:
        print(f"   - {f['name']}")
except Exception as e:
    print(f"   ERROR: {e}")

# Test 2: Create sheet WITHOUT folder_id (in service account's Drive)
print("\n2. Creating test sheet in service account Drive...")
try:
    sheet = client.create("Test_Sheet_Temp")
    print(f"   Created: {sheet.url}")
    # Clean up
    client.del_spreadsheet(sheet.id)
    print("   Deleted test sheet")
except Exception as e:
    print(f"   ERROR: {e}")

# Test 3: Create sheet WITH folder_id
print("\n3. Creating test sheet in shared folder...")
try:
    sheet = client.create("Test_Sheet_Shared", folder_id=FOLDER_ID)
    print(f"   Created: {sheet.url}")
except Exception as e:
    print(f"   ERROR: {e}")
