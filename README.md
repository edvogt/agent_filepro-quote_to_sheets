# FilePro Quote → Google Sheets Sync

Monitors a spool directory for FilePro quotation TSV exports and automatically creates versioned Google Sheets in a shared Drive folder. Runs as a systemd service on Linux.

## Data Flow

```
Browser → CGI (quot-edit-sheets-acct)
             ↓  rreport stquote -f export-tsv -R <quote#>
         exports/QUOTE_<N>_<TS>.tsv
             ↓  watchdog (filepro_sync.py)
         _parse_tsv_file() → JSON written
             ↓
         GoogleSheetsClient.create_or_update_sheet()
             ↓  new versioned sheet: Quote-<N>-<version>
         _apply_formatting() + call_webhook()
             ↓
         URL logged → JSON archived → TSV deleted
             ↓
         CGI redirects browser to sheet
```

## Setup

```bash
# Install dependencies
pip install gspread google-auth google-auth-oauthlib watchdog pandas --break-system-packages

# First-time OAuth setup (interactive, requires browser)
python setup_oauth.py

# Run the service
python filepro_sync.py
```

OAuth token is stored at `/home/filepro/credentials/token.pickle`. Re-run `setup_oauth.py` if the token expires.

## Service Management

```bash
sudo systemctl restart filepro-sync
sudo systemctl status filepro-sync
sudo journalctl -u filepro-sync -f
```

## Configuration

`CONFIG` dict at the top of `filepro_sync.py`:

| Key | Default | Notes |
|-----|---------|-------|
| `token_file` | `/home/filepro/credentials/token.pickle` | OAuth token |
| `google_drive_folder_id` | `1fRcg-tMAOkt81KVbI4h56zZFI7Hu6him` | CLIENT_QUOTES folder |
| `export_directory` | `exports/` | Watch directory for TSV files |
| `archive_directory` | `/home/filepro/exports/archive` | Processed JSON files; archived into `YYYY-MM/` |
| `url_log_file` | `/home/filepro/quote_urls.log` | One URL per line |
| `webhook_url` | Apps Script deployment URL | Set to `''` to disable |

## Manually Triggering a Sync

```bash
# Export a quote from FilePro and drop it in the watch directory
./stquote_to_spool.sh <quote_number>
```

Or copy any `QUOTE_*.tsv` into `exports/` — the watcher picks it up within 2 seconds.

## Logs & URLs

```bash
# Live service log
tail -f filepro_sync.log

# Successful syncs
grep "SYNCED" filepro_sync.log

# Sheet URLs as HTML links
./quote_url_html.sh              # Latest
./quote_url_html.sh 88960        # Specific quote
./quote_url_html.sh last 5       # Last 5
./quote_url_html.sh all          # All as <ul>
```

## CGI Script

`quot-edit-sheets-acct` is the Apache CGI entry point (deployed to `/var/www/html/secure/cgi-bin/`). It accepts a quote number, runs `rreport`, drops the TSV into `exports/`, waits 15 seconds for the sync to complete, then redirects the browser to the new Google Sheet.

After editing: `cp quot-edit-sheets-acct /var/www/html/secure/cgi-bin/`

## Apps Script Webhook

`format_quote_sheet.js` is deployed as a Google Apps Script web app. After each sync, `call_webhook()` POSTs the full `json_data` dict (quote_info, customer, financial_summary, line_items) plus `sheet_url` and `timestamp` to it. The Apps Script creates and formats the sheet.

**Deploy**: Apps Script → Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
**Update `webhook_url`** in CONFIG after each new deployment.

**Manual test** (GET):
```bash
curl 'https://script.google.com/macros/s/<deploy_id>/exec'
# Returns: {"status":"ready",...}
```
