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

**Data flow**: FilePro exports JSON → Watchdog detects file → QuotationProcessor parses → GoogleSheetsClient creates/updates Sheet → Webhook called (optional) → File archived

Three main components in `filepro_sync.py`:

1. **GoogleSheetsClient** (line 83): OAuth token loading/refresh, Sheets CRUD operations, formatting. Uses `gspread` library.
2. **QuotationProcessor** (line 277): JSON parsing, quote number extraction from filename, data cleaning with pandas, coordinates sync workflow.
3. **QuotationFileHandler** (line 395): Watchdog `FileSystemEventHandler` subclass. Monitors export directory, debounces file creation events (2-second delay), prevents duplicate processing.

Supporting files:
- `setup_oauth.py`: One-time OAuth flow generating `/home/filepro/credentials/token.pickle`
- `test_sheets.py`: Diagnostic script using service account auth (different auth path than main app)
- `call_webhook()` (line 241): POST to Apps Script after successful sync

## Configuration

`CONFIG` dict at `filepro_sync.py:39`:
- `token_file`: OAuth token pickle path (default: `/home/filepro/credentials/token.pickle`)
- `google_drive_folder_id`: Target Drive folder ID (CLIENT-QUOTES: `1SG2iyJ1ej_MUyu4WEJyImWG8iz78A-j0`)
- `export_directory`: Watch directory (default: `/appl/spool/QUOTES-SHEETS`)
- `file_pattern`: Glob pattern (default: `QUOTE_*.json`)
- `archive_directory`: Where processed files move (default: `/home/filepro/exports/archive`)
- `webhook_url`: Optional Apps Script URL for notifications

## JSON Format

Filename pattern: `QUOTE_[NUMBER]_[TIMESTAMP].json` (e.g., `QUOTE_12345_20250124_143022.json`) - quote number extracted from second underscore segment.

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

**FilePro format**: Actual quote files use `meta`, `invoiced_to`, `ship_to`, `quote_details`, `entry_details`, `line_items`, and `totals` sections - fully supported by the processor.
