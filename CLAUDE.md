# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python service that monitors a directory for FilePro quotation TSV exports and automatically creates versioned Google Sheets. Runs as a systemd service (`filepro-sync.service`) using watchdog for filesystem monitoring.

## Key Commands

```bash
# Install dependencies
pip install --user gspread google-auth google-auth-oauthlib watchdog pandas

# First-time OAuth setup (interactive - requires browser)
python setup_oauth.py

# Run the sync service (normally runs via systemd)
python filepro_sync.py

# Service management (after code changes)
sudo systemctl restart filepro-sync
sudo systemctl status filepro-sync
sudo journalctl -u filepro-sync -f
# Unit file: /etc/systemd/system/filepro-sync.service

# Manually trigger a sync
./stquote_to_spool.sh <quote_number>

# Tail the log
tail -f filepro_sync.log

# Log grep patterns
grep "SYNCED"  filepro_sync.log   # Successful syncs with sheet URL
grep "ERROR"   filepro_sync.log   # Failures
grep "Webhook" filepro_sync.log   # Webhook call results

# Sheet URLs as HTML links
./quote_url_html.sh              # Latest link
./quote_url_html.sh 86016        # Specific quote
./quote_url_html.sh all          # All links as <ul> list
./quote_url_html.sh last 5       # Last 5 links
```

## Architecture

**Data flow:**
```
Browser → CGI (quot-edit-sheets-acct)
            ↓  rreport stquote → TSV
        exports/QUOTE_<N>_<TS>.tsv
            ↓  watchdog (filepro_sync.py)
        _parse_tsv_file() → JSON written to disk
            ↓
        GoogleSheetsClient.create_or_update_sheet()
            ↓  new versioned sheet: Quote-<N>-<version>
        _apply_formatting() (Python/gspread) → call_webhook() (Apps Script, optional)
            ↓
        URL logged → JSON archived → TSV deleted
            ↓
        CGI redirects browser to sheet
```

### Three classes in `filepro_sync.py`

1. **GoogleSheetsClient**: OAuth token loading/refresh, Sheets CRUD via `gspread`.
   - `create_or_update_sheet()`: Always creates a **new** versioned sheet `Quote-{number}-{version}` — never updates existing sheets
   - `_get_next_version()`: Scans Drive folder for existing `Quote-{number}-N` sheets, returns next version
   - `_populate_worksheet()`: Writes header rows, line items, and financial summary in one batch update
   - `_apply_formatting()`: Python-side formatting pass (dark blue header, light blue column headers, freeze row 4, auto-resize). The Apps Script webhook runs afterward as the final formatting pass.

2. **QuotationProcessor**: TSV parsing, quote number extraction, data cleaning with pandas, coordinates sync workflow.
   - `_parse_tsv_file()`: Reads tab-delimited export using DELIMITED-ITEM-SECTION sentinels. Supports multi-row per quote (one row per page). Collects header from first row, line items across all rows.
   - `process_file()`: Main entry point — parse TSV → write JSON → sync to Sheets → archive JSON → delete TSV. Includes a **quote number mismatch guard**: if the TSV content's quote number doesn't match the filename's quote number, the file is skipped with an error (prevents wrong-record syncs from stale FilePro tmp files).
   - `_parse_filepro_json()`, `_convert_filepro_metadata()`, `_fix_json_file()`: **Dead code** — legacy JSON pipeline methods never called in the current TSV flow; kept for reference only. Do not delete without asking.

3. **QuotationFileHandler**: Watchdog `FileSystemEventHandler`. Debounces file creation events (2-second delay). File is added to `self.processing` set **before** the sleep to block concurrent duplicate watchdog events for the same file.

### Module-level function

- `call_webhook()`: POSTs full `json_data` dict (quote_info, customer, financial_summary, line_items) plus `sheet_url` and `timestamp` to Apps Script. Skips silently if `webhook_url` is empty string.

### Supporting files

| File | Purpose |
|---|---|
| `stquote_to_spool.sh` | CLI export: runs `rreport stquote -f tabexport -R <N> -A` → drops TSV in `exports/` |
| `quot-edit-sheets-acct` | Apache CGI: runs `rreport stquote -f export-tsv -R <N> -V selquote -A` → drops TSV, waits 15s, redirects to sheet |
| `createBoundScript.js` | `createBoundScript_(ssId)` — called from doPost to attach a bound script (Publish menu + Save as PDF) to each new sheet via Apps Script API |
| `format_quote_sheet.js` | Google Apps Script web app (v1.1.0) — retired, kept for reference |
| `setup_oauth.py` | One-time OAuth flow → `/home/filepro/credentials/token.pickle` |
| `test_sheets.py` | Diagnostic script using service account auth (separate from OAuth) |
| `quote_url_html.sh` | Output sheet URLs as HTML links |
| `fix_filepro_json.sh` | Legacy — unused in current TSV pipeline |

**Note on rreport format names:** `stquote_to_spool.sh` uses `-f tabexport` (prc.tabexport), while the CGI uses `-f export-tsv` (prc.export-tsv with `-V selquote`). Both produce the same 83-column TSV layout but use different FilePro process files and output paths.

## CGI Script

`quot-edit-sheets-acct` — Apache CGI at `/var/www/html/secure/cgi-bin/`. After edits: `cp quot-edit-sheets-acct /var/www/html/secure/cgi-bin/`

**Flow:**
1. Outputs "Creating filepro export file..." + 4 KB HTML comment buffer flush (so browser renders before rreport runs)
2. Runs `/appl/bin/rlf` to remove FilePro lockfiles; removes stale `/appl/spool/QUOTES-SHEETS-tmp.tsv`
3. Runs `rreport stquote -f export-tsv -R "$STRING1" -V selquote -A` (stderr → `/tmp/rreport-cgi.log`) → writes to `/appl/spool/QUOTES-SHEETS-tmp.tsv`
4. Copies tmp TSV → `$EXPORTS/QUOTE_{N}_{TS}.tsv`
5. Outputs success message + spinner gif
6. Sleeps 15 seconds for sync + webhook to complete
7. Greps `filepro_sync.log` for `SYNCED | Quote {N} |` and extracts URL
8. Redirects browser to the Google Sheet via `window.location.href`

**Key variables:** `EXPORTS=/home/filepro/agent_filepro-quote_to_sheets/exports`, `SYNCLOG=filepro_sync.log`

## Authentication

1. **OAuth (main app)**: `filepro_sync.py` uses token at `/home/filepro/credentials/token.pickle`. Auto-refreshes on expiry. Re-run `setup_oauth.py` if refresh fails.
2. **Service Account (test_sheets.py only)**: Uses a separate service account JSON key. Do not confuse with the main app auth.

## TSV Input Format

FilePro exports one-row tab-delimited files named `QUOTE_[NUMBER]_[TIMESTAMP].tsv` to the watch directory.

**DELIMITED-ITEM-SECTION format** (sentinel: `DELIMITED-ITEM-SECTION`):

FilePro exports one or more tab-delimited rows per quote (one row per page). The first row (lowest page number) contains header data. Line items are collected across all rows.

| Col offset | Content |
|----|----|---|
| 0 | `DELIMITED-ITEM-SECTION` sentinel |
| +1 | item_num |
| +2 | qty_ordered |
| +3 | price |
| +4 | total |
| +5 | shipped |
| +6 | qty_to_ship |
| +7 | description |
| +8 | cost |
| +9 | extension |
| +10 | new_inv_desc |

Header constants: H_QUOTE_NUM=0, H_PAGE_NUM=1, H_INVOICE_NUM=2, H_CUST_PO=3, etc. Service items have item_num starting with `#`. Items with blank item_num are skipped.

## JSON Output Format

`process_file()` writes a structured JSON alongside the TSV before syncing. This JSON is what gets archived to `YYYY-MM/` subdirectories.

```json
{
  "quote_info":        { "quote_number", "date", "purchase_order_ref", "terms", "ship_via", "sold_by" },
  "customer":          { "bill_to": { "name", "organization", "address" } },
  "financial_summary": { "sub_total", "tax_amount", "shipping", "total_amount" },
  "line_items":        [ { "qty", "part_id", "description", "price_each", "price_extended" }, ... ]
}
```

## Configuration

`CONFIG` dict at top of `filepro_sync.py`:

| Key | Default | Notes |
|---|---|---|
| `token_file` | `/home/filepro/credentials/token.pickle` | OAuth token |
| `google_drive_folder_id` | `1fRcg-tMAOkt81KVbI4h56zZFI7Hu6him` | CLIENT_QUOTES folder |
| `export_directory` | `exports/` (absolute path in code) | Watch directory for TSV files |
| `file_pattern` | `QUOTE_*.tsv` | Glob pattern for watchdog |
| `log_file` | `filepro_sync.log` | Relative to working directory |
| `url_log_file` | `/home/filepro/quote_urls.log` | One URL per line: `timestamp \| Quote N \| URL` |
| `archive_directory` | `/home/filepro/exports/archive` | Processed JSON files; `YYYY-MM/` subdirectories |
| `webhook_url` | Apps Script deployment URL | Set to `''` to disable. Update after each Apps Script redeployment. |
| `webhook_timeout` | `30` | Seconds |

Current webhook URL: `https://script.google.com/macros/s/AKfycbzlBVpK45Ezvhfmg0_zPnsLT5PhWAxErV54eCElc5JiBZ2RCVnnoRDS1-VXDtgl9_9g_g/exec`

## Webhook / Apps Script

**Current endpoint**: "Filepro Quote Publish 01" v8.5.0, deployed as Version 11 ("add publish to pdf option").
- Script ID: `1LR2sMsnUebMVG0VL63rbRwwGO3Fc6rDyJiPN0I2_8SSs7407qy6QbWp4`
- Deployment ID: `AKfycbzlBVpK45Ezvhfmg0_zPnsLT5PhWAxErV54eCElc5JiBZ2RCVnnoRDS1-VXDtgl9_9g_g`
- Execute as: `dev@ear.net` | Access: Anyone

**The endpoint requires a full quote payload** — there is no format-only mode. Expected POST body:
```json
{
  "quote_info":        { "quote_number", "date", "purchase_order_ref", "terms", "ship_via", "sold_by" },
  "customer":          { "bill_to": { "name", "organization", "address" } },
  "financial_summary": { "sub_total", "tax_amount", "shipping", "total_amount" },
  "line_items":        [ { "qty", "part_id", "description", "price_each", "price_extended" }, ... ]
}
```

**Webhook payload**: `call_webhook()` sends the full `json_data` dict (quote_info, customer, financial_summary, line_items) plus `sheet_url` and `timestamp`. The Apps Script uses this to create and format the sheet.

**Formatting applied by the Apps Script**:
- EAR blue (`#1a73e8`) header + column headers, Roboto font
- Alternating row colors for line items
- Green `TOTAL` row with currency formatting (`$#,##0.00`)
- Hidden gridlines, auto-resized columns, frozen header rows

**Bound script — Publish menu** (v8.5.0): `doPost` calls `createBoundScript_(ssId)` after creating each new sheet. This uses the Apps Script API (`script.googleapis.com/v1/projects`) to create a container-bound script project on the spreadsheet containing `onOpen()` (adds Publish > Save as PDF menu) and `publishAsPDF()` (removes images, exports as PDF via `/export?format=pdf` + OAuth token, saves to same Drive folder with `<sheet name>.pdf`, replaces existing, confirms via alert). Requires Apps Script API enabled in the Cloud project and `https://www.googleapis.com/auth/script.projects` in the standalone script's `oauthScopes`.

**Old endpoint** (`format_quote_sheet.js` v1.1.0): Retired. Was a format-only webhook that accepted `{quote_number, sheet_url, timestamp}`.

**Deploy**: Apps Script → Deploy → Manage deployments → Edit → set Version to "New version" → Deploy

**Manual test** (GET): `curl -sL 'https://script.google.com/macros/s/AKfycbzlBVpK45Ezvhfmg0_zPnsLT5PhWAxErV54eCElc5JiBZ2RCVnnoRDS1-VXDtgl9_9g_g/exec'`

## Startup Behavior

On startup, `process_existing_files()` processes any `QUOTE_*.tsv` files already in the watch directory before starting the watchdog observer. The observer then monitors for new files in real-time.

## FilePro Source Database

Quote data comes from `/appl/filepro/stquote/`. The TSV layout is defined by `/appl/filepro/stquote/prc.tabexport` (CLI) and `prc.export-tsv` (CGI).

## API Notes

- `gspread` `CellFormat.borders` only supports `top`/`bottom`/`left`/`right` — `innerHorizontal`/`innerVertical` are invalid there (they belong in `UpdateBordersRequest`)
- Sheet naming is always versioned: `Quote-{number}-{version}`. Never updates an existing sheet.

## Testing

There are no automated tests. `test_sheets.py` is a manual diagnostic script (uses service account auth, separate from OAuth). To verify changes, trigger a manual sync and check the log:
```bash
./stquote_to_spool.sh <quote_number>
grep "SYNCED\|ERROR" filepro_sync.log | tail -5
```
