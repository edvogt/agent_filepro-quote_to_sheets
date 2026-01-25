#!/usr/bin/env python3
"""
FilePro to Google Sheets Sync Script
=====================================
This script monitors a directory for FilePro quotation exports and 
automatically syncs them to Google Sheets.

Requirements:
    pip install gspread google-auth watchdog pandas --break-system-packages

Setup:
    1. Create Google Service Account at https://console.cloud.google.com
    2. Download credentials JSON file
    3. Share target Google Drive folder with service account email
    4. Update CONFIG section below with your settings
"""

import os
import time
import json
import logging
import pickle
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import gspread
from google.auth.transport.requests import Request
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ============================================================================
# CONFIGURATION
# ============================================================================
CONFIG = {
    # Google Authentication (OAuth)
    'token_file': '/home/filepro/credentials/token.pickle',
    'google_drive_folder_id': '1jnWNqdSBy8aigv2c_MclEtv4r-HJvsmB',
    
    # FilePro Export Settings
    'export_directory': '/home/filepro/exports/quotations',
    'file_pattern': 'QUOTE_*.json',

    # Sync Settings
    'check_interval': 60,  # seconds
    'archive_processed': True,
    'archive_directory': '/home/filepro/exports/archive',
    
    # Google Sheets Settings
    'sheet_prefix': 'Quote',
    'template_sheet_id': None,  # Optional: copy from template
    
    # Logging
    'log_file': 'filepro_sync.log',
    'log_level': 'INFO',

    # Webhook (Google Apps Script Web App URL)
    'webhook_url': None,  # Set to your deployed Apps Script URL
    'webhook_timeout': 30  # seconds
}

# ============================================================================
# LOGGING SETUP
# ============================================================================
logging.basicConfig(
    level=getattr(logging, CONFIG['log_level']),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(CONFIG['log_file']),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('FilePro-Sync')

# ============================================================================
# GOOGLE SHEETS CLIENT
# ============================================================================
class GoogleSheetsClient:
    """Handles all Google Sheets operations"""

    def __init__(self, token_file: str, folder_id: str):
        self.folder_id = folder_id

        # Load OAuth token
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)

        # Refresh if expired
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(token_file, 'wb') as token:
                pickle.dump(creds, token)

        self.client = gspread.authorize(creds)
        logger.info("Google Sheets client initialized successfully")
    
    def find_sheet_by_quote_number(self, quote_number: str) -> Optional[gspread.Spreadsheet]:
        """Find existing sheet by quotation number"""
        try:
            # Search for sheets in the target folder
            spreadsheets = self.client.list_spreadsheet_files(
                folder_id=self.folder_id
            )
            
            target_name = f"{CONFIG['sheet_prefix']} {quote_number}"
            for sheet in spreadsheets:
                if sheet['name'] == target_name:
                    return self.client.open_by_key(sheet['id'])
            
            return None
        except Exception as e:
            logger.error(f"Error searching for sheet {quote_number}: {e}")
            return None
    
    def create_or_update_sheet(self, quote_number: str, data: pd.DataFrame) -> tuple:
        """Create new sheet or update existing one. Returns (success, sheet_url)"""
        try:
            sheet_name = f"{CONFIG['sheet_prefix']} {quote_number}"

            # Check if sheet exists
            existing_sheet = self.find_sheet_by_quote_number(quote_number)

            if existing_sheet:
                logger.info(f"Updating existing sheet: {sheet_name}")
                worksheet = existing_sheet.sheet1
                spreadsheet = existing_sheet

                # Clear existing content
                worksheet.clear()
            else:
                logger.info(f"Creating new sheet: {sheet_name}")

                # Create new spreadsheet
                spreadsheet = self.client.create(
                    sheet_name,
                    folder_id=self.folder_id
                )
                worksheet = spreadsheet.sheet1

            # Write data to sheet
            self._populate_worksheet(worksheet, quote_number, data)

            # Apply formatting
            self._apply_formatting(worksheet)

            logger.info(f"Successfully synced quotation {quote_number}")
            return True, spreadsheet.url

        except Exception as e:
            logger.error(f"Error creating/updating sheet {quote_number}: {e}")
            return False, None
    
    def _populate_worksheet(self, worksheet: gspread.Worksheet,
                           quote_number: str, data: pd.DataFrame):
        """Populate worksheet with quotation data"""

        # Build all rows at once
        all_rows = [
            ["QUOTATION", f"#{quote_number}"],
            ["Date", datetime.now().strftime("%Y-%m-%d")],
            ["", ""],  # Blank row
            data.columns.tolist()  # Column headers
        ]

        # Add data rows (convert numpy types to native Python)
        for _, row in data.iterrows():
            all_rows.append([float(v) if pd.notna(v) and isinstance(v, (int, float)) else (str(v) if pd.notna(v) else "") for v in row.tolist()])

        # Add totals section if numeric columns exist
        numeric_cols = data.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            all_rows.append([""])  # Blank row
            for col in numeric_cols:
                total = float(data[col].sum())  # Convert to native Python float
                all_rows.append([f"Total {col}", total])

        # Batch update all rows at once
        worksheet.update(f'A1:Z{len(all_rows)}', all_rows)
    
    def _apply_formatting(self, worksheet: gspread.Worksheet):
        """Apply professional formatting to worksheet"""
        try:
            # Format header rows (bold, colored background)
            worksheet.format('A1:B2', {
                'backgroundColor': {'red': 0.12, 'green': 0.31, 'blue': 0.47},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}},
                'horizontalAlignment': 'LEFT'
            })
            
            # Format column headers
            worksheet.format('4:4', {
                'backgroundColor': {'red': 0.84, 'green': 0.91, 'blue': 0.94},
                'textFormat': {'bold': True},
                'horizontalAlignment': 'CENTER'
            })
            
            # Auto-resize columns
            worksheet.columns_auto_resize(0, worksheet.col_count)
            
            # Freeze header rows
            worksheet.freeze(rows=4)
            
        except Exception as e:
            logger.warning(f"Error applying formatting: {e}")

# ============================================================================
# WEBHOOK CALLER
# ============================================================================
def call_webhook(quote_number: str, sheet_url: str) -> bool:
    """Call Apps Script webhook after successful sync"""
    if not CONFIG.get('webhook_url'):
        return True  # No webhook configured, skip

    try:
        payload = json.dumps({
            'quote_number': quote_number,
            'sheet_url': sheet_url,
            'timestamp': datetime.now().isoformat()
        }).encode('utf-8')

        req = urllib.request.Request(
            CONFIG['webhook_url'],
            data=payload,
            headers={'Content-Type': 'application/json'},
            method='POST'
        )

        with urllib.request.urlopen(req, timeout=CONFIG['webhook_timeout']) as response:
            logger.info(f"Webhook called successfully for quote {quote_number}")
            return True

    except urllib.error.HTTPError as e:
        logger.error(f"Webhook HTTP error for quote {quote_number}: {e.code} {e.reason}")
        return False
    except urllib.error.URLError as e:
        logger.error(f"Webhook URL error for quote {quote_number}: {e.reason}")
        return False
    except Exception as e:
        logger.error(f"Webhook error for quote {quote_number}: {e}")
        return False

# ============================================================================
# FILE PROCESSOR
# ============================================================================
class QuotationProcessor:
    """Processes FilePro quotation exports"""
    
    def __init__(self, sheets_client: GoogleSheetsClient):
        self.sheets_client = sheets_client
    
    def process_file(self, file_path: Path) -> bool:
        """Process a single quotation file"""
        try:
            logger.info(f"Processing file: {file_path}")
            
            # Extract quotation number from filename
            quote_number = self._extract_quote_number(file_path)
            if not quote_number:
                logger.error(f"Could not extract quote number from {file_path}")
                return False
            
            # Read JSON data
            data = pd.read_json(file_path)
            logger.info(f"Loaded {len(data)} rows for quote {quote_number}")
            
            # Clean and validate data
            data = self._clean_data(data)

            # Sync to Google Sheets
            success, sheet_url = self.sheets_client.create_or_update_sheet(
                quote_number,
                data
            )

            # Call webhook if configured
            if success and sheet_url:
                call_webhook(quote_number, sheet_url)

            # Archive file if successful
            if success and CONFIG['archive_processed']:
                self._archive_file(file_path)

            return success
            
        except Exception as e:
            logger.error(f"Error processing file {file_path}: {e}")
            return False
    
    def _extract_quote_number(self, file_path: Path) -> Optional[str]:
        """Extract quotation number from filename"""
        # Example: QUOTE_12345_20250124.json -> 12345
        try:
            parts = file_path.stem.split('_')
            if len(parts) >= 2 and parts[0] == 'QUOTE':
                return parts[1]
        except:
            pass
        return None
    
    def _clean_data(self, data: pd.DataFrame) -> pd.DataFrame:
        """Clean and validate quotation data"""
        # Remove empty rows
        data = data.dropna(how='all')
        
        # Remove empty columns
        data = data.dropna(axis=1, how='all')
        
        # Clean column names
        data.columns = [col.strip() for col in data.columns]
        
        # Convert numeric columns
        for col in data.columns:
            if 'price' in col.lower() or 'amount' in col.lower() or 'total' in col.lower():
                try:
                    data[col] = pd.to_numeric(data[col], errors='coerce')
                except:
                    pass
        
        return data
    
    def _archive_file(self, file_path: Path):
        """Move processed file to archive directory"""
        try:
            archive_dir = Path(CONFIG['archive_directory'])
            archive_dir.mkdir(parents=True, exist_ok=True)
            
            # Create dated subdirectory
            date_dir = archive_dir / datetime.now().strftime('%Y-%m')
            date_dir.mkdir(exist_ok=True)
            
            # Move file
            archive_path = date_dir / file_path.name
            file_path.rename(archive_path)
            logger.info(f"Archived file to {archive_path}")
            
        except Exception as e:
            logger.warning(f"Could not archive file {file_path}: {e}")

# ============================================================================
# FILE WATCHER
# ============================================================================
class QuotationFileHandler(FileSystemEventHandler):
    """Watches for new quotation files and processes them"""
    
    def __init__(self, processor: QuotationProcessor):
        self.processor = processor
        self.processing = set()
    
    def on_created(self, event):
        """Handle new file creation"""
        if event.is_directory:
            return
        
        file_path = Path(event.src_path)
        
        # Check if it matches our pattern
        if not file_path.match(CONFIG['file_pattern']):
            return
        
        # Avoid duplicate processing
        if str(file_path) in self.processing:
            return
        
        # Wait for file to be fully written
        time.sleep(2)
        
        # Process the file
        self.processing.add(str(file_path))
        try:
            self.processor.process_file(file_path)
        finally:
            self.processing.discard(str(file_path))

# ============================================================================
# MAIN APPLICATION
# ============================================================================
def process_existing_files(processor: QuotationProcessor, directory: Path):
    """Process any existing files in the directory"""
    logger.info(f"Checking for existing files in {directory}")
    
    for file_path in directory.glob(CONFIG['file_pattern']):
        if file_path.is_file():
            processor.process_file(file_path)

def main():
    """Main application entry point"""
    logger.info("=" * 60)
    logger.info("FilePro to Google Sheets Sync - Starting")
    logger.info("=" * 60)
    
    # Validate configuration
    export_dir = Path(CONFIG['export_directory'])
    if not export_dir.exists():
        logger.error(f"Export directory does not exist: {export_dir}")
        return
    
    token_file = Path(CONFIG['token_file'])
    if not token_file.exists():
        logger.error(f"OAuth token file not found: {token_file}")
        logger.error("Run setup_oauth.py first to authenticate")
        return

    # Initialize Google Sheets client
    try:
        sheets_client = GoogleSheetsClient(
            CONFIG['token_file'],
            CONFIG['google_drive_folder_id']
        )
    except Exception as e:
        logger.error(f"Failed to initialize Google Sheets client: {e}")
        return
    
    # Initialize processor
    processor = QuotationProcessor(sheets_client)
    
    # Process existing files
    process_existing_files(processor, export_dir)
    
    # Set up file watcher
    event_handler = QuotationFileHandler(processor)
    observer = Observer()
    observer.schedule(event_handler, str(export_dir), recursive=False)
    observer.start()
    
    logger.info(f"Monitoring directory: {export_dir}")
    logger.info("Press Ctrl+C to stop")
    
    try:
        while True:
            time.sleep(CONFIG['check_interval'])
    except KeyboardInterrupt:
        logger.info("Stopping file watcher...")
        observer.stop()
    
    observer.join()
    logger.info("Sync application stopped")

if __name__ == '__main__':
    main()
