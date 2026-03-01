#!/usr/bin/env python3
"""
FilePro to Google Sheets Sync Script
=====================================
Version: 1.0.1
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
import re
import csv
import time
import json
import logging
import pickle
import subprocess
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

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
    'google_drive_folder_id': '1fRcg-tMAOkt81KVbI4h56zZFI7Hu6him',  # CLIENT_QUOTES folder
    
    # FilePro Export Settings
    # $SPOOL = /appl/spool
    'export_directory': '/home/filepro/agent_filepro-quote_to_sheets/exports',
    'file_pattern': 'QUOTE_*.tsv',

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
    'url_log_file': '/home/filepro/quote_urls.log',

    # Webhook (Google Apps Script Web App URL)
    'webhook_url': 'https://script.google.com/macros/s/AKfycbyr34ZZ5h7kZpUtJwu7Pn_O2XxUq3FfW3Wb027PCbNSUaav8jnEgbxU-YbpAMJJlGcK/exec',
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
# TSV COLUMN LAYOUT (stquote/prc.tabexport field order)
# ============================================================================
# Columns 0-22: quote header and customer info
TSV_HEADER = {
    'quote_number':  0,
    'quote_date':    1,
    'cust_po':       2,
    'terms':         3,
    'ship_via':      4,
    'salesperson':   5,
    'subtotal':      6,
    'tax':           7,
    'total':         8,
    'company_name':  9,
    'bill_contact': 10,
    'bill_addr1':   11,
    'bill_addr2':   12,
    'bill_city':    13,
    'bill_state':   14,
    'bill_zip':     15,
    'ship_company': 16,
    'ship_contact': 17,
    'ship_addr1':   18,
    'ship_addr2':   19,
    'ship_city':    20,
    'ship_state':   21,
    'ship_zip':     22,
}
# Columns 23+: 10 line item slots, 6 fields each
# Slot N (0-indexed): base = 23 + (N * 6)
# +0 Item Number, +1 Qty, +2 Price, +3 Extension, +4 Description, +5 New Inv Description
TSV_ITEM_BASE   = 23
TSV_ITEM_FIELDS = 6
TSV_ITEM_COUNT  = 10

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
    
    def _get_next_version(self, quote_number: str) -> int:
        """Scan folder for Quote-{quote_number}-N sheets and return the next version number."""
        try:
            spreadsheets = self.client.list_spreadsheet_files(folder_id=self.folder_id)
            prefix = f"Quote-{quote_number}-"
            max_version = 0
            for sheet in spreadsheets:
                name = sheet['name']
                if name.startswith(prefix):
                    try:
                        version = int(name[len(prefix):])
                        if version > max_version:
                            max_version = version
                    except ValueError:
                        pass
            return max_version + 1
        except Exception as e:
            logger.error(f"Error checking versions for quote {quote_number}: {e}")
            return 1

    def create_or_update_sheet(self, quote_number: str, data: pd.DataFrame, metadata: dict = None) -> tuple:
        """Create a new versioned sheet (Quote-{number}-{version}). Returns (success, sheet_url)"""
        try:
            version = self._get_next_version(quote_number)
            sheet_name = f"Quote-{quote_number}-{version}"

            logger.info(f"Creating new sheet: {sheet_name}")
            spreadsheet = self.client.create(sheet_name, folder_id=self.folder_id)
            worksheet = spreadsheet.sheet1

            self._populate_worksheet(worksheet, quote_number, data, metadata)
            self._apply_formatting(worksheet)

            logger.info(f"Successfully synced quotation {quote_number}")
            return True, spreadsheet.url

        except Exception as e:
            logger.error(f"Error creating sheet {quote_number}: {e}")
            return False, None
    
    def _populate_worksheet(self, worksheet: gspread.Worksheet,
                           quote_number: str, data: pd.DataFrame, metadata: dict = None):
        """Populate worksheet with quotation data"""

        all_rows = []

        # Add header info from metadata if available
        if metadata and metadata.get('quote_info'):
            qi = metadata['quote_info']
            all_rows.append(["QUOTATION", f"#{qi.get('quote_number', quote_number)}"])
            all_rows.append(["Date", qi.get('date', datetime.now().strftime("%Y-%m-%d"))])
            all_rows.append(["PO Reference", qi.get('purchase_order_ref', '')])
            all_rows.append(["Terms", qi.get('terms', '')])
            all_rows.append(["Ship Via", qi.get('ship_via', '')])
        else:
            all_rows.append(["QUOTATION", f"#{quote_number}"])
            all_rows.append(["Date", datetime.now().strftime("%Y-%m-%d")])

        # Add customer info if available
        if metadata and metadata.get('customer'):
            cust = metadata['customer']
            bill_to = cust.get('bill_to', {})
            all_rows.append(["", ""])
            all_rows.append(["BILL TO", bill_to.get('name', '')])
            all_rows.append(["", bill_to.get('organization', '')])
            all_rows.append(["", bill_to.get('address', '')])

        all_rows.append(["", ""])  # Blank row
        all_rows.append(data.columns.tolist())  # Column headers

        # Add data rows (convert numpy types to native Python)
        for _, row in data.iterrows():
            all_rows.append([float(v) if pd.notna(v) and isinstance(v, (int, float)) else (str(v) if pd.notna(v) else "") for v in row.tolist()])

        # Add financial summary if available
        if metadata and metadata.get('financial_summary'):
            fs = metadata['financial_summary']
            all_rows.append([""])
            all_rows.append(["Sub Total", float(fs.get('sub_total', 0))])
            all_rows.append(["Tax", float(fs.get('tax_amount', 0))])
            all_rows.append(["Shipping", float(fs.get('shipping', 0))])
            all_rows.append(["TOTAL", float(fs.get('total_amount', 0))])
        else:
            # Add totals section if numeric columns exist
            numeric_cols = data.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                all_rows.append([""])  # Blank row
                for col in numeric_cols:
                    total = float(data[col].sum())
                    all_rows.append([f"Total {col}", total])

        # Batch update all rows at once
        worksheet.update(all_rows, f'A1:Z{len(all_rows)}')
    
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
        self._quote_metadata = None

    def _parse_tsv_file(self, file_path: Path) -> tuple[List[Dict], Dict[str, Any]]:
        """Parse a FilePro tab-delimited quote export into (line_items, metadata)."""
        with open(file_path, 'r') as f:
            reader = csv.reader(f, delimiter='\t')
            rows = [row for row in reader if any(field.strip() for field in row)]

        if not rows:
            return [], {}

        row = rows[0]

        # Ensure row is long enough
        expected = TSV_ITEM_BASE + (TSV_ITEM_COUNT * TSV_ITEM_FIELDS)
        while len(row) < expected:
            row.append('')

        def col(i):
            return row[i].strip() if i < len(row) else ''

        def to_float(val):
            try:
                return float(val) if val else 0
            except ValueError:
                return 0

        metadata = {
            'quote_info': {
                'quote_number':      col(TSV_HEADER['quote_number']),
                'date':              col(TSV_HEADER['quote_date']),
                'purchase_order_ref': col(TSV_HEADER['cust_po']),
                'terms':             col(TSV_HEADER['terms']),
                'ship_via':          col(TSV_HEADER['ship_via']),
                'sold_by':           col(TSV_HEADER['salesperson']),
            },
            'customer': {
                'bill_to': {
                    'name':         col(TSV_HEADER['company_name']),
                    'organization': col(TSV_HEADER['bill_contact']),
                    'address':      ', '.join(filter(None, [
                        col(TSV_HEADER['bill_addr1']),
                        col(TSV_HEADER['bill_addr2']),
                        col(TSV_HEADER['bill_city']),
                        col(TSV_HEADER['bill_state']),
                        col(TSV_HEADER['bill_zip']),
                    ])),
                }
            },
            'financial_summary': {
                'sub_total':    to_float(col(TSV_HEADER['subtotal'])),
                'tax_amount':   to_float(col(TSV_HEADER['tax'])),
                'shipping':     0,
                'total_amount': to_float(col(TSV_HEADER['total'])),
            },
        }

        line_items = []
        for i in range(TSV_ITEM_COUNT):
            base = TSV_ITEM_BASE + (i * TSV_ITEM_FIELDS)
            item_num  = col(base)
            qty       = col(base + 1)
            price     = col(base + 2)
            ext       = col(base + 3)
            desc      = col(base + 4)
            new_desc  = col(base + 5)

            if not item_num:
                continue

            line_items.append({
                'qty':             qty,
                'part_id':         item_num,
                'description':     new_desc or desc,
                'price_each':      price,
                'price_extended':  ext,
            })

        return line_items, metadata

    def _fix_json_file(self, file_path: Path) -> bool:
        """
        Run fix_filepro_json.sh to fix malformed JSON before processing.
        Returns True if fix was successful or script not configured.
        """
        fix_script = CONFIG.get('json_fix_script')
        if not fix_script or not Path(fix_script).exists():
            logger.debug("JSON fix script not configured or not found, skipping")
            return True

        try:
            logger.info(f"Running JSON fix script on: {file_path}")
            result = subprocess.run(
                [fix_script, str(file_path)],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode != 0:
                logger.error(f"JSON fix script failed: {result.stderr}")
                return False

            logger.info(f"JSON fix script completed: {result.stdout.strip()}")
            return True

        except subprocess.TimeoutExpired:
            logger.error(f"JSON fix script timed out for {file_path}")
            return False
        except Exception as e:
            logger.error(f"Error running JSON fix script: {e}")
            return False

    def _parse_filepro_json(self, file_path: Path) -> tuple[List[Dict], Dict[str, Any]]:
        """
        Parse FilePro JSON format which has malformed structure:
        - Line items are loose objects between entry_details and totals (not in array)
        - Missing commas in totals section
        Returns (line_items, metadata_dict)
        """
        with open(file_path, 'r') as f:
            content = f.read()

        # Extract the structured sections using regex
        metadata = {}

        # Parse meta section
        meta_match = re.search(r'"meta"\s*:\s*(\{[^}]+\})', content, re.DOTALL)
        if meta_match:
            try:
                metadata['meta'] = json.loads(meta_match.group(1))
            except json.JSONDecodeError:
                metadata['meta'] = {}

        # Parse invoiced_to section
        invoiced_match = re.search(r'"invoiced_to"\s*:\s*(\{[^}]+\})', content, re.DOTALL)
        if invoiced_match:
            try:
                metadata['invoiced_to'] = json.loads(invoiced_match.group(1))
            except json.JSONDecodeError:
                metadata['invoiced_to'] = {}

        # Parse ship_to section
        ship_match = re.search(r'"ship_to"\s*:\s*(\{[^}]+\})', content, re.DOTALL)
        if ship_match:
            try:
                metadata['ship_to'] = json.loads(ship_match.group(1))
            except json.JSONDecodeError:
                metadata['ship_to'] = {}

        # Parse quote_details section
        quote_details_match = re.search(r'"quote_details"\s*:\s*(\{[^}]+\})', content, re.DOTALL)
        if quote_details_match:
            try:
                metadata['quote_details'] = json.loads(quote_details_match.group(1))
            except json.JSONDecodeError:
                metadata['quote_details'] = {}

        # Parse entry_details section
        entry_match = re.search(r'"entry_details"\s*:\s*(\{[^}]+\})', content, re.DOTALL)
        if entry_match:
            try:
                metadata['entry_details'] = json.loads(entry_match.group(1))
            except json.JSONDecodeError:
                metadata['entry_details'] = {}

        # Parse totals section (fix malformed JSON - missing commas, empty values)
        totals_match = re.search(r'"totals"\s*:\s*\{([^}]+)\}', content, re.DOTALL)
        if totals_match:
            totals_content = totals_match.group(1)
            metadata['totals'] = {}
            # Match key-value pairs, handling keys with colons and numeric/null values
            # Pattern: "Key Name:" or "Key Name" followed by optional whitespace and value
            for match in re.finditer(r'"([^"]+)"\s*:\s*([-\d.]+|null)?', totals_content):
                key = match.group(1).strip().rstrip(':')
                val = match.group(2)
                if val is None or val == 'null' or val.strip() == '':
                    metadata['totals'][key] = None
                else:
                    try:
                        metadata['totals'][key] = float(val.strip())
                    except ValueError:
                        metadata['totals'][key] = None

        # Extract line items - they're loose objects with "type": "line"
        line_items = []
        # Match all objects that have qty, part_id, description pattern
        item_pattern = re.compile(
            r'\{\s*"qty"\s*:\s*"([^"]*)"\s*,\s*'
            r'"part_id"\s*:\s*"([^"]*)"\s*,\s*'
            r'"description"\s*:\s*"([^"]*)"\s*,\s*'
            r'"price_each"\s*:\s*"([^"]*)"\s*,\s*'
            r'"price_extended"\s*:\s*"([^"]*)"\s*,\s*'
            r'"type"\s*:\s*"([^"]*)"\s*\}',
            re.DOTALL
        )

        for match in item_pattern.finditer(content):
            qty_str = match.group(1).strip()
            price_each_str = match.group(4).strip()
            price_ext_str = match.group(5).strip()

            line_items.append({
                'qty': int(qty_str) if qty_str.isdigit() else (float(qty_str) if qty_str else ''),
                'part_id': match.group(2).strip(),
                'description': match.group(3).strip(),
                'price_each': float(price_each_str) if price_each_str and price_each_str.replace('-', '').replace('.', '').isdigit() else '',
                'price_extended': float(price_ext_str) if price_ext_str and price_ext_str.replace('-', '').replace('.', '').isdigit() else '',
                'type': match.group(6).strip()
            })

        return line_items, metadata

    def _convert_filepro_metadata(self, fp_metadata: Dict[str, Any]) -> Dict[str, Any]:
        """Convert FilePro metadata structure to the format expected by GoogleSheetsClient"""
        meta = fp_metadata.get('meta', {})
        quote_details = fp_metadata.get('quote_details', {})
        entry_details = fp_metadata.get('entry_details', {})
        invoiced_to = fp_metadata.get('invoiced_to', {})
        ship_to = fp_metadata.get('ship_to', {})
        totals = fp_metadata.get('totals', {})

        # Build quote_info from meta and quote_details
        quote_info = {
            'quote_number': meta.get('quote_number') or quote_details.get('QUOTE#', ''),
            'date': meta.get('quote_date') or quote_details.get('DATE', ''),
            'purchase_order_ref': quote_details.get('PURCHASE ORDER #', ''),
            'terms': quote_details.get('TERMS OF SALE', ''),
            'ship_via': entry_details.get('SHIP VIA', ''),
            'order_number': quote_details.get('ORDER #', ''),
            'sold_by': entry_details.get('SOLD BY:', '')
        }

        # Build customer info from invoiced_to
        invoiced_lines = invoiced_to.get('lines', [])
        customer = {
            'bill_to': {
                'name': invoiced_lines[0] if len(invoiced_lines) > 0 else '',
                'organization': invoiced_lines[1] if len(invoiced_lines) > 1 else '',
                'address': ', '.join(invoiced_lines[2:]) if len(invoiced_lines) > 2 else ''
            }
        }

        # Build financial summary from totals
        financial_summary = {
            'sub_total': totals.get('Sub Total') or totals.get('Sub Total:') or 0,
            'tax_amount': totals.get('Tax Amount (T)') or totals.get('Tax') or 0,
            'shipping': totals.get('Shipping') or 0,
            'total_amount': totals.get('Total') or totals.get('Total:') or 0
        }

        return {
            'quote_info': quote_info,
            'customer': customer,
            'vendor': {},
            'financial_summary': financial_summary
        }

    def process_file(self, file_path: Path) -> bool:
        """Process a single quotation TSV file"""
        try:
            logger.info(f"Processing file: {file_path}")

            # Parse TSV
            line_items, metadata = self._parse_tsv_file(file_path)

            if not line_items:
                logger.error(f"No line items found in {file_path}")
                return False

            # Quote number: from filename first, then from TSV data
            quote_number = self._extract_quote_number(file_path)
            if not quote_number:
                quote_number = metadata.get('quote_info', {}).get('quote_number')
            if not quote_number:
                logger.error(f"Could not determine quote number from {file_path}")
                return False

            data = pd.DataFrame(line_items)
            data = self._clean_data(data)

            logger.info(f"Loaded {len(data)} rows for quote {quote_number}")

            # Convert TSV data to JSON and write alongside TSV for archiving
            json_data = {
                'quote_info':       metadata.get('quote_info', {}),
                'customer':         metadata.get('customer', {}),
                'financial_summary': metadata.get('financial_summary', {}),
                'line_items':       line_items,
            }
            json_path = file_path.with_suffix('.json')
            with open(json_path, 'w') as jf:
                json.dump(json_data, jf, indent=2)
            logger.info(f"Wrote JSON: {json_path}")

            # Sync to Google Sheets
            success, sheet_url = self.sheets_client.create_or_update_sheet(
                quote_number,
                data,
                metadata=metadata
            )

            if success and sheet_url:
                logger.info(f"SYNCED | Quote {quote_number} | {sheet_url}")
                url_log = Path(CONFIG.get('url_log_file', '/home/filepro/quote_urls.log'))
                with open(url_log, 'a') as f:
                    f.write(f"{datetime.now().isoformat()} | Quote {quote_number} | {sheet_url}\n")
                print(f"  Quote {quote_number} synced successfully", flush=True)
                print(f"    {sheet_url}", flush=True)
                call_webhook(quote_number, sheet_url)

            # Archive JSON, remove TSV
            if success and CONFIG['archive_processed']:
                self._archive_file(json_path)
                file_path.unlink()
                logger.info(f"Removed TSV: {file_path}")

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
        
        # Avoid duplicate processing — add to set before sleep so concurrent
        # events for the same file are blocked while we wait for the write to finish
        if str(file_path) in self.processing:
            return
        self.processing.add(str(file_path))

        try:
            # Wait for file to be fully written
            time.sleep(2)

            # Check file still exists (may have been processed by another event)
            if not file_path.exists():
                return

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
