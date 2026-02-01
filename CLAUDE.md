# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python application that monitors a directory for FilePro accounting system quotation exports (JSON) and automatically syncs them to Google Sheets. Runs as a background service on Ubuntu using watchdog for filesystem monitoring.

## Key Commands

```bash
# Install dependencies
pip install --user gspread google-auth google-auth-oauthlib watchdog pandas

# First-time OAuth setup (interactive - requires browser)
python setup_oauth.py

# Run the sync service
python filepro_sync.py

# Test Google Sheets API connection (uses service account, not OAuth)
python test_sheets.py
```

## Architecture

**Data flow**: FilePro exports JSON → Watchdog detects file → JSON fix script runs → QuotationProcessor parses → GoogleSheetsClient creates/updates Sheet → Webhook formats sheet → URL logged → File archived

Three main components in `filepro_sync.py`:

1. **GoogleSheetsClient** (line 89): OAuth token loading/refresh, Sheets CRUD operations, formatting. Uses `gspread` library.
   - `find_sheet_by_quote_number()` (line 108): Searches existing sheets in target folder
   - `create_or_update_sheet()` (line 126): Main entry point for sheet creation/update
   - `_populate_worksheet()` (line 164): Writes header, line items, and financial summary
   - `_apply_formatting()` (line 218): Applies colors, bold headers, auto-resize, freeze rows

2. **QuotationProcessor** (line 283): JSON parsing, quote number extraction from filename, data cleaning with pandas, coordinates sync workflow.
   - `_fix_json_file()` (line 290): Calls `fix_filepro_json.sh` to fix malformed JSON before parsing
   - `_parse_filepro_json()` (line 323): Regex-based parser for malformed FilePro JSON (fallback)
   - `_convert_filepro_metadata()` (line 423): Transforms FilePro structure to standard format
   - `process_file()` (line 468): Main entry point - fixes JSON, handles trigger files, all formats

3. **QuotationFileHandler** (line 647): Watchdog `FileSystemEventHandler` subclass. Monitors export directory, debounces file creation events (2-second delay), prevents duplicate processing via `self.processing` set.

Supporting files:
- `setup_oauth.py`: One-time OAuth flow generating `/home/filepro/credentials/token.pickle`
- `test_sheets.py`: Diagnostic script using service account auth (different auth path than main app)
- `call_webhook()` (line 247): POST to Apps Script after successful sync
- `format_quote_sheet.js`: Google Apps Script for sheet formatting (deployed as web app)
- `quote_url_html.sh`: Output sheet URLs in HTML link format

## Authentication

Two authentication methods exist in this codebase:

1. **OAuth (main app)**: `filepro_sync.py` uses OAuth tokens stored in pickle file. Run `setup_oauth.py` once to generate the token interactively.

2. **Service Account (test script)**: `test_sheets.py` uses a service account JSON key file. This is a separate auth path for diagnostics only.

## JSON Fix Script

`fix_filepro_json.sh` automatically fixes malformed FilePro JSON before processing. Called by `_fix_json_file()` (line 290) in `QuotationProcessor`.

**Fixes applied:**
- Wraps loose line item objects in `"line_items": [...]` array
- Adds missing commas in totals section
- Fixes empty `"Tax":` value → `"Tax": null`
- Removes trailing commas before `]` or `}`
- Normalizes spacing in numeric values

**Manual usage:**
```bash
# Fix in place (creates .orig backup)
./fix_filepro_json.sh QUOTE_12345.json

# Fix to new file
./fix_filepro_json.sh input.json output.json
```

## Configuration

`CONFIG` dict at `filepro_sync.py:40`:
- `token_file`: OAuth token pickle path (default: `/home/filepro/credentials/token.pickle`)
- `google_drive_folder_id`: Target Drive folder ID (CLIENT-QUOTES: `1SG2iyJ1ej_MUyu4WEJyImWG8iz78A-j0`)
- `export_directory`: Watch directory (default: `/appl/spool/QUOTES-SHEETS`)
- `file_pattern`: Glob pattern (default: `QUOTE_*.json`)
- `json_fix_script`: Path to bash script that fixes malformed JSON (default: `fix_filepro_json.sh`)
- `url_log_file`: Log file for sheet URLs (default: `/home/filepro/quote_urls.log`)
- `archive_directory`: Where processed files move (default: `/home/filepro/exports/archive`)
- `webhook_url`: Apps Script URL for sheet formatting (deployed web app URL)

## JSON Format

Filename pattern: `QUOTE_[NUMBER]_[TIMESTAMP].json` (e.g., `QUOTE_12345_20250124_143022.json`) - quote number extracted from second underscore segment via `_extract_quote_number()` (line 594).

**Nested format** (code expects `line_items` key):
```json
{
  "line_items": [...],
  "quote_info": {...},
  "customer": {...},
  "financial_summary": {...}
}
```

**Flat format**: Simple array of line item objects.

**Trigger file format**: Spool directory contains small trigger files with `html_path` pointing to actual quote JSON:
```json
{
  "quote_number": "91697",
  "html_path": "/appl/fileprow/quotes/91697.json"
}
```

**FilePro format**: Actual quote files use `meta`, `invoiced_to`, `ship_to`, `quote_details`, `entry_details`, `line_items`, and `totals` sections. This format has malformed JSON (missing commas, loose objects) that requires regex-based parsing in `_parse_filepro_json()`.

## Webhook Formatting

After successful sync, the webhook calls `format_quote_sheet.js` (deployed as Google Apps Script web app) to apply professional formatting:
- Dark blue header with quote number
- Light blue column headers
- Alternating row colors for line items
- Green totals section with currency formatting
- Auto-resized columns and frozen header rows

Deploy the script: Apps Script → Deploy → New deployment → Web app → Anyone

## URL Logging

Sheet URLs are logged to `/home/filepro/quote_urls.log` in format:
```
2026-01-31T18:53:09.143790 | Quote 86016 | https://docs.google.com/spreadsheets/d/...
```

**Get URLs as HTML links:**
```bash
./quote_url_html.sh              # Latest link
./quote_url_html.sh 86016        # Specific quote
./quote_url_html.sh all          # All links as <ul> list
./quote_url_html.sh last 5       # Last 5 links
```

## Startup Behavior

On startup, `process_existing_files()` (line 682) processes any files already in the watch directory before starting the watchdog observer. The observer then monitors for new files in real-time.

## Note on README.md

The README.md describes an older Windows/CSV-based workflow with service account auth. The actual implementation uses Ubuntu/JSON with OAuth tokens. Refer to this CLAUDE.md and the code for current behavior.
