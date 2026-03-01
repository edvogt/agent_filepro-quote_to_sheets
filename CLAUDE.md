# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python application that monitors a directory for FilePro quotation exports (TSV) and automatically syncs them to Google Sheets. Runs as a background service on Ubuntu using watchdog for filesystem monitoring.

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

**Data flow**: CGI (`quot-edit-sheets-acct`) → `rreport` exports TSV → dropped in `exports/` → Watchdog detects file → `_parse_tsv_file()` parses → JSON written to disk → GoogleSheetsClient creates versioned Sheet → Webhook formats sheet → URL logged → JSON archived → TSV removed → CGI redirects browser to sheet

Three main components in `filepro_sync.py`:

1. **GoogleSheetsClient** (line 123): OAuth token loading/refresh, Sheets CRUD operations, formatting. Uses `gspread` library.
   - `_get_next_version()`: Scans folder for `Quote-{number}-N` sheets and returns the next version number
   - `create_or_update_sheet()`: Always creates a new versioned sheet named `Quote-{number}-{version}` (e.g. `Quote-88960-3`); never updates an existing sheet
   - `_populate_worksheet()`: Writes header, line items, and financial summary
   - `_apply_formatting()`: Rough Python-side formatting pass (colors, bold headers, auto-resize, freeze row 4). The Apps Script webhook runs afterward as the final formatting pass.

2. **QuotationProcessor** (line 317): TSV parsing, quote number extraction, data cleaning with pandas, coordinates sync workflow.
   - `_parse_tsv_file()`: Reads tab-delimited export, maps 83 columns to metadata and up to 10 line item slots
   - `process_file()`: Main entry point — parse TSV → write JSON → sync to Sheets → archive JSON → delete TSV
   - `_parse_filepro_json()`, `_convert_filepro_metadata()`, `_fix_json_file()`: **Dead code** — legacy JSON pipeline methods never called in the current TSV flow; kept for reference only

3. **QuotationFileHandler**: Watchdog `FileSystemEventHandler` subclass. Monitors export directory, debounces file creation events (2-second delay), prevents duplicate processing via `self.processing` set. File is added to the set **before** the debounce sleep to block concurrent duplicate watchdog events.

Supporting files:
- `stquote_to_spool.sh`: Called by FilePro watchfolder — sets `TERM=ansi`, `FP=/appl/fp`, `PFSKIPLOCKED=1`, runs `rreport` against `stquote` with `prc.tabexport -R <quote_number>`, writes to `/appl/fpmerge/quote_export.tsv`, then moves to `exports/QUOTE_<N>_<TIMESTAMP>.tsv`
- `setup_oauth.py`: One-time OAuth flow generating `/home/filepro/credentials/token.pickle`
- `test_sheets.py`: Diagnostic script using service account auth (different auth path than main app)
- `call_webhook()` (module-level function): POSTs `{quote_number, sheet_url, timestamp}` to Apps Script. The JS side extracts the sheet ID from the URL via regex (`extractSheetId()`).
- `format_quote_sheet.js`: Google Apps Script web app for final sheet formatting. `findDataStartRow()` locates the column-header row dynamically; `findTotalsStartRow()` scans from the bottom for "Sub Total".
- `fix_filepro_json.sh`: Legacy script for malformed JSON (older pipeline, unused).
- `quote_url_html.sh`: Output sheet URLs in HTML link format

## CGI Script

`/var/www/html/secure/cgi-bin/quot-edit-sheets-acct` — Apache CGI that triggers a quote export and streams progress to the browser.

**Flow:**
1. Outputs "Creating filepro export file..." immediately (with 4 KB buffer flush so browser renders before rreport runs)
2. Runs `/appl/bin/rlf` to remove FilePro lockfiles
3. Runs `rreport stquote -f export-tsv -s tmp` → copies output to `exports/QUOTE_{N}_{TS}.tsv`
4. Outputs "Success: FilePro generated export File..." + "queued for google sheets processing, stand by..." + spinner gif (125×100 px)
5. Sleeps 15 seconds for sync + webhook to complete
6. Greps `filepro_sync.log` for `SYNCED | Quote {N} |` and extracts the URL
7. Redirects browser to the Google Sheet via `window.location.href`

**Key variables:**
- `EXPORTS`: `/home/filepro/agent_filepro-quote_to_sheets/exports`
- `SYNCLOG`: `/home/filepro/agent_filepro-quote_to_sheets/filepro_sync.log`
- Spinner: `https://acct.ear.net/filepro/gif/circle-spinning-blue-to-black.gif`

**Not in the repo** — lives only at `/var/www/html/secure/cgi-bin/quot-edit-sheets-acct` on the web server.

## Authentication

Two authentication methods exist in this codebase:

1. **OAuth (main app)**: `filepro_sync.py` uses OAuth tokens stored in pickle file. Run `setup_oauth.py` once to generate the token interactively.

2. **Service Account (test script)**: `test_sheets.py` uses a service account JSON key file. This is a separate auth path for diagnostics only.

## TSV Input Format

FilePro exports tab-delimited files named `QUOTE_[NUMBER]_[TIMESTAMP].tsv` (e.g., `QUOTE_96036_20260227_120000.tsv`) to the watch directory. One record per file (one row of data).

**83-column layout** defined by `TSV_HEADER`, `TSV_ITEM_BASE`, `TSV_ITEM_FIELDS`, `TSV_ITEM_COUNT` constants:

| Columns | Content |
|---|---|
| 0–8 | Quote#, Date, Cust PO#, Terms, Ship Via, Salesperson, Subtotal, Tax, Total |
| 9–15 | Company Name, Bill Contact, Bill Addr1/2, City, State, Zip |
| 16–22 | Ship Company, Ship Contact, Ship Addr1/2, City, State, Zip |
| 23–82 | 10 line item slots × 6 fields (Item#, Qty, Price, Extension, Description, New Inv Description) |

Line item slots with an empty Item# field are skipped. Items prefixed with `#` are FilePro memo/comment lines (no qty or price).

The longer `New Inv Description` (field +5 per slot) is used as the primary description; falls back to `Description` (field +4) if empty.

## JSON Output Format

After parsing, `process_file()` writes a structured JSON file (`QUOTE_[NUMBER]_[TIMESTAMP].json`) alongside the TSV before syncing. This is the file that gets archived. Format:

```json
{
  "quote_info":        { "quote_number", "date", "purchase_order_ref", "terms", "ship_via", "sold_by" },
  "customer":          { "bill_to": { "name", "organization", "address" } },
  "financial_summary": { "sub_total", "tax_amount", "shipping", "total_amount" },
  "line_items":        [ { "qty", "part_id", "description", "price_each", "price_extended" }, ... ]
}
```

## Configuration

`CONFIG` dict at `filepro_sync.py:42`:
- `token_file`: OAuth token pickle path (default: `/home/filepro/credentials/token.pickle`)
- `google_drive_folder_id`: Target Drive folder ID (CLIENT_QUOTES: `1fRcg-tMAOkt81KVbI4h56zZFI7Hu6him`)
- `export_directory`: Watch directory (default: `/home/filepro/agent_filepro-quote_to_sheets/exports`)
- `file_pattern`: Glob pattern (default: `QUOTE_*.tsv`)
- `log_file`: Relative path — `filepro_sync.log` created in the working directory where `filepro_sync.py` is launched
- `url_log_file`: Log file for sheet URLs (default: `/home/filepro/quote_urls.log`)
- `archive_directory`: Where processed JSON files move (default: `/home/filepro/exports/archive`); archived into `YYYY-MM/` subdirectory
- `webhook_url`: Apps Script URL for sheet formatting (deployed web app URL)
- `webhook_timeout`: HTTP timeout in seconds for Apps Script call (default: 30)

## Webhook Formatting

After successful sync, the webhook calls `format_quote_sheet.js` (deployed as Google Apps Script web app) to apply final professional formatting:
- Dark blue header with quote number
- Light blue column headers
- Alternating row colors for line items
- Green totals section with currency formatting (`$#,##0.00`)
- Auto-resized columns and frozen header rows

**Deploy the script**: Apps Script → Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone

**Manual testing**: The web app also accepts GET requests with `?sheet_id=SHEET_ID` for triggering format without a Python sync. The `testFormat()` function in `format_quote_sheet.js` hardcodes a specific sheet ID for in-editor testing.

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

On startup, `process_existing_files()` processes any `QUOTE_*.tsv` files already in the watch directory before starting the watchdog observer. The observer then monitors for new files in real-time.

## FilePro Source Database

Quote data comes from `/appl/filepro/stquote/`. The TSV export is driven by `/appl/filepro/stquote/prc.tabexport` which maps stquote fields to the 83-column layout above.

## Service Management

The sync runs as a systemd service (`filepro-sync.service`), enabled and auto-restarting.

```bash
# After code changes
sudo systemctl restart filepro-sync

# Check status
sudo systemctl status filepro-sync

# Live log via journald
sudo journalctl -u filepro-sync -f
```

## Debugging & Operations

**Tail the service log:**
```bash
tail -f /home/filepro/agent_filepro-quote_to_sheets/filepro_sync.log
```

**Manually trigger a sync** (drop a TSV in the watch dir to simulate FilePro export):
```bash
./stquote_to_spool.sh <quote_number>
```
Or copy an existing TSV into `exports/` — the watcher will pick it up within 2 seconds.

**Check the webhook** is reachable (GET returns `{"status":"ready",...}`):
```bash
curl 'https://script.google.com/macros/s/<deploy_id>/exec'
```

**Re-run OAuth if token expires:**
```bash
python setup_oauth.py
```

**Verify service account path (test_sheets.py only):** Separate from OAuth — uses a service account JSON key. Do not confuse with the main app auth.

**Log grep patterns:**
```bash
grep "SYNCED"    filepro_sync.log   # Successful syncs with sheet URL
grep "ERROR"     filepro_sync.log   # Failures
grep "Webhook"   filepro_sync.log   # Webhook call results
```
