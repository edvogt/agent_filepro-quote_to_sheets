# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python application that monitors a directory for FilePro accounting system quotation exports (JSON files) and automatically syncs them to Google Sheets. It runs as a background service on Ubuntu, watching for new JSON exports and creating/updating corresponding Google Sheets.

## Key Commands

```bash
# Install dependencies
pip install --user gspread google-auth watchdog pandas

# Run the sync service
python filepro_sync.py
```

## Architecture

The application consists of three main components in `filepro_sync.py`:

1. **GoogleSheetsClient** (line 74): Handles Google API authentication and all Sheets operations (create, update, find, format)
2. **QuotationProcessor** (line 213): Parses JSON files, extracts quote numbers from filenames, cleans data, and coordinates syncing
3. **QuotationFileHandler** (line 306): Watchdog event handler that monitors the export directory for new JSON files

**Data flow**: FilePro exports JSON -> Watchdog detects new file -> Processor parses JSON -> GoogleSheetsClient creates/updates Sheet -> File archived

## Configuration

All settings are in the `CONFIG` dictionary at the top of `filepro_sync.py` (line 35). Key settings:
- `service_account_file`: Path to Google service account JSON credentials
- `google_drive_folder_id`: Target Google Drive folder ID
- `export_directory`: Directory to watch for JSON exports
- `file_pattern`: Expected filename pattern (default: `QUOTE_*.json`)

## Expected JSON Format

- Filename: `QUOTE_[NUMBER]_[TIMESTAMP].json` (e.g., `QUOTE_12345_20250124_143022.json`)
- Quote number is extracted from the second underscore-delimited segment of the filename
