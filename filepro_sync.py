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
    'webhook_url': 'https://script.google.com/macros/s/AKfycbxkC58mywofLZ0jqsPAawrVv_SCf0thQqKNZswSfpZ4TVRAjQFWQyxMgHyWhUlFe4DMhw/exec',
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
def call_webhook(quote_number: str, json_data: dict) -> Optional[str]:
    """Call Apps Script webhook — creates new sheet, returns sheet_url or None on failure"""
    if not CONFIG.get('webhook_url'):
        logger.warning("No webhook_url configured — skipping webhook")
        return None

    try:
        full_payload = dict(json_data)  # copy quote_info, customer, financial_summary, line_items
        full_payload['quote_number'] = quote_number
        full_payload['timestamp']    = datetime.now().isoformat()
        payload = json.dumps(full_payload).encode('utf-8')

        req = urllib.request.Request(
            CONFIG['webhook_url'],
            data=payload,
            headers={'Content-Type': 'application/json'},
            method='POST'
        )

        with urllib.request.urlopen(req, timeout=CONFIG['webhook_timeout']) as response:
            body = response.read().decode('utf-8')
            resp_json = json.loads(body)
            sheet_url = resp_json.get('sheet_url', '')
            if sheet_url:
                logger.info(f"Webhook created new sheet for quote {quote_number}: {sheet_url}")
            else:
                logger.warning(f"Webhook succeeded but no sheet_url returned for quote {quote_number}")
            return sheet_url or None

    except urllib.error.HTTPError as e:
        logger.error(f"Webhook HTTP error for quote {quote_number}: {e.code} {e.reason}")
        return None
    except urllib.error.URLError as e:
        logger.error(f"Webhook URL error for quote {quote_number}: {e.reason}")
        return None
    except Exception as e:
        logger.error(f"Webhook error for quote {quote_number}: {e}")
        return None

# ============================================================================
# FILE PROCESSOR
# ============================================================================
class QuotationProcessor:
    """Processes FilePro quotation exports"""

    def __init__(self, sheets_client: GoogleSheetsClient):
        self.sheets_client = sheets_client
        self._quote_metadata = None

    def _parse_tsv_file(self, file_path: Path) -> tuple[List[Dict], Dict[str, Any]]:
        """Parse FilePro TSV export using DELIMITED-ITEM-SECTION sentinels.

        TSV format (one row per quote page):
          Header section  - cols 0-71, fixed field positions
          Item sections   - groups of 11 cols each:
                            col+0  DELIMITED-ITEM-SECTION  (sentinel)
                            col+1  item_num
                            col+2  qty_ordered
                            col+3  price
                            col+4  total
                            col+5  shipped
                            col+6  qty_to_ship
                            col+7  description
                            col+8  cost
                            col+9  extension
                            col+10 new_inv_desc

        Multiple rows may exist for the same quote (one per page).
        Header data is taken from the first (lowest page number) row.
        Line items are collected across all pages.

        Service items have item_num starting with '#' - these are
        preserved with type='service'. Items with blank item_num are skipped.
        """

        DELIM         = 'DELIMITED-ITEM-SECTION'
        ITEM_NUM      = 1
        ITEM_QTY      = 2
        ITEM_PRICE    = 3
        ITEM_TOTAL    = 4
        ITEM_SHIPPED  = 5
        ITEM_QSHIP    = 6
        ITEM_DESC     = 7
        ITEM_COST     = 8
        ITEM_EXT      = 9
        ITEM_NEWDESC  = 10

        # Header column positions (0-based)
        H_QUOTE_NUM    =  0
        H_PAGE_NUM     =  1
        H_INVOICE_NUM  =  2
        H_CUST_PO      =  3
        H_PROFIT_CTR   =  4
        H_TERMS        =  7
        H_QUOTE_DATE   = 10
        H_SUBTOTAL     = 17
        H_FREIGHT_IN   = 18
        H_FREIGHT_OUT  = 19
        H_TOTAL        = 20
        H_TAX          = 22
        H_SALESPERSON  = 24
        H_SHIP_VIA     = 29
        H_COMPANY      = 39
        H_BILL_1       = 40
        H_BILL_2       = 41
        H_BILL_3       = 42
        H_BILL_4       = 43
        H_BILL_5       = 44
        H_SHIP_1       = 45
        H_SHIP_2       = 46
        H_SHIP_3       = 47
        H_BILL_CONTACT = 57
        H_TAX_CODE     = 58
        H_STATUS       = 59
        H_SHIP_COMPANY = 63
        H_SHIP_CONTACT = 64
        H_ECOM_EMAIL   = 65

        def col(row, idx, default=''):
            """Safe column accessor."""
            try:
                return row[idx].strip() if row[idx] else default
            except IndexError:
                return default

        def to_float(val, default=0.0):
            """Convert string to float safely."""
            if not val:
                return default
            try:
                return float(val.replace(',', ''))
            except (ValueError, AttributeError):
                return default

        # Read and group all rows by quote number
        rows_by_quote = {}
        try:
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                for raw_line in f:
                    row = raw_line.rstrip('\n').split('\t')
                    if not row or not row[0]:
                        continue
                    qnum = row[0].strip()
                    if not qnum:
                        continue
                    if qnum not in rows_by_quote:
                        rows_by_quote[qnum] = []
                    rows_by_quote[qnum].append(row)
        except Exception as e:
            logger.error(f"Error reading TSV file {file_path}: {e}")
            return [], {}

        if not rows_by_quote:
            logger.warning(f"No valid rows found in {file_path}")
            return [], {}

        # Use the first quote number found (file contains one quote, possibly multi-page)
        quote_number = next(iter(rows_by_quote))
        all_rows = rows_by_quote[quote_number]

        # Sort by page number, use page 1 row for header data
        def page_num(row):
            try:
                return int(row[H_PAGE_NUM]) if len(row) > H_PAGE_NUM and row[H_PAGE_NUM].isdigit() else 999
            except (ValueError, IndexError):
                return 999

        all_rows.sort(key=page_num)
        hdr = all_rows[0]

        # Build bill-to address from bill_1 through bill_5
        bill_parts = [col(hdr, H_BILL_1), col(hdr, H_BILL_2),
                      col(hdr, H_BILL_3), col(hdr, H_BILL_4),
                      col(hdr, H_BILL_5)]
        bill_address = ', '.join(p for p in bill_parts if p)

        # Build ship-to address from ship_1 through ship_3
        ship_parts = [col(hdr, H_SHIP_1), col(hdr, H_SHIP_2), col(hdr, H_SHIP_3)]
        ship_address = ', '.join(p for p in ship_parts if p)

        metadata = {
            'quote_info': {
                'quote_number':    col(hdr, H_QUOTE_NUM),
                'date':            col(hdr, H_QUOTE_DATE),
                'purchase_order_ref': col(hdr, H_CUST_PO),
                'terms':           col(hdr, H_TERMS),
                'ship_via':        col(hdr, H_SHIP_VIA),
                'sold_by':         col(hdr, H_SALESPERSON),
                'profit_center':   col(hdr, H_PROFIT_CTR),
                'invoice_number':  col(hdr, H_INVOICE_NUM),
                'status':          col(hdr, H_STATUS),
                'tax_code':        col(hdr, H_TAX_CODE),
            },
            'customer': {
                'bill_to': {
                    'name':         col(hdr, H_BILL_CONTACT),
                    'organization': col(hdr, H_COMPANY),
                    'address':      bill_address,
                    'email':        col(hdr, H_ECOM_EMAIL),
                },
                'ship_to': {
                    'name':         col(hdr, H_SHIP_CONTACT),
                    'organization': col(hdr, H_SHIP_COMPANY),
                    'address':      ship_address,
                },
            },
            'financial_summary': {
                'sub_total':    to_float(col(hdr, H_SUBTOTAL)),
                'tax_amount':   to_float(col(hdr, H_TAX)),
                'freight_in':   to_float(col(hdr, H_FREIGHT_IN)),
                'freight_out':  to_float(col(hdr, H_FREIGHT_OUT)),
                'total_amount': to_float(col(hdr, H_TOTAL)),
            },
        }

        # Collect line items across all pages
        line_items = []
        seen_items = set()  # deduplicate across pages

        for row in all_rows:
            # Find all delimiter positions in this row
            delim_positions = [j for j, v in enumerate(row) if v == DELIM]
            if not delim_positions:
                continue

            # Determine block size from spacing between delimiters
            if len(delim_positions) > 1:
                block_size = delim_positions[1] - delim_positions[0]
            else:
                block_size = 11  # default

            for pos in delim_positions:
                fields = row[pos:pos + block_size]

                def f(offset, default=''):
                    try:
                        return fields[offset].strip() if offset < len(fields) and fields[offset] else default
                    except IndexError:
                        return default

                item_num = f(ITEM_NUM)

                # Skip empty blocks
                if not item_num:
                    continue

                # Deduplicate (same item can appear on multiple pages)
                dedup_key = f"{item_num}|{f(ITEM_QTY)}|{f(ITEM_PRICE)}"
                if dedup_key in seen_items:
                    continue
                seen_items.add(dedup_key)

                # Determine item type
                item_type = 'service' if item_num.startswith('#') else 'product'

                # Description: prefer new_inv_desc, fall back to description
                new_inv = f(ITEM_NEWDESC)
                desc    = f(ITEM_DESC)
                display_desc = new_inv if new_inv else desc

                line_items.append({
                    'part_id':        item_num,
                    'type':           item_type,
                    'qty':            f(ITEM_QTY),
                    'price_each':     to_float(f(ITEM_PRICE)),
                    'total':          to_float(f(ITEM_TOTAL)),
                    'shipped':        f(ITEM_SHIPPED),
                    'qty_to_ship':    f(ITEM_QSHIP),
                    'discount':       to_float(f(ITEM_QSHIP)),
                    'description':    display_desc,
                    'short_desc':     desc,
                    'cost':           to_float(f(ITEM_COST)),
                    'price_extended':      to_float(f(ITEM_EXT)),
                })

        logger.info(f"Parsed quote {quote_number}: {len(line_items)} line items across {len(all_rows)} page(s)")
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

            # Safety check: TSV content must match the filename quote number.
            # A mismatch means FilePro exported the wrong record (stale tmp selection).
            tsv_quote_number = metadata.get('quote_info', {}).get('quote_number', '').strip()
            if tsv_quote_number and tsv_quote_number != quote_number:
                logger.error(
                    f"Quote number mismatch: filename={quote_number}, "
                    f"TSV content={tsv_quote_number} — FilePro exported the wrong record. "
                    f"Skipping {file_path.name}."
                )
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
            # Apps Script creates the new sheet and returns its URL
            sheet_url = call_webhook(quote_number, json_data)
            success = bool(sheet_url)

            if success and sheet_url:
                logger.info(f"SYNCED | Quote {quote_number} | {sheet_url}")
                url_log = Path(CONFIG.get('url_log_file', '/home/filepro/quote_urls.log'))
                with open(url_log, 'a') as f:
                    f.write(f"{datetime.now().isoformat()} | Quote {quote_number} | {sheet_url}\n")
                print(f"  Quote {quote_number} synced successfully", flush=True)
                print(f"    {sheet_url}", flush=True)

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
