#!/bin/bash
# stquote_to_spool.sh - Export a FilePro quote to the Google Sheets watch directory
#
# Usage: stquote_to_spool.sh <quote_number>
#
# Called by FilePro watchfolder process when a quote is flagged for export.
# Runs rreport against stquote, writes TSV to fpmerge, then renames and
# moves to the spool directory for pickup by filepro_sync.py.

set -euo pipefail

if [[ -z "${1:-}" ]]; then
    echo "Usage: $0 <quote_number>" >&2
    exit 1
fi

QUOTE_NUM="$1"
FPMERGE_FILE="/appl/fpmerge/quote_export.tsv"
SPOOL_DIR="/home/filepro/agent_filepro-quote_to_sheets/exports"
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
DEST="${SPOOL_DIR}/QUOTE_${QUOTE_NUM}_${TIMESTAMP}.tsv"

export TERM=ansi
export FPTERM=ansi
export FP=/appl/fp
export PFSKIPLOCKED=1

# Remove any leftover file from a previous run
rm -f "$FPMERGE_FILE"

# Run FilePro export for the specific quote record
/appl/fp/rreport stquote -f tabexport -R "$QUOTE_NUM" -A >/dev/null 2>&1

if [[ ! -f "$FPMERGE_FILE" ]]; then
    echo "Error: rreport did not produce $FPMERGE_FILE for quote $QUOTE_NUM" >&2
    exit 1
fi

# Move to spool with QUOTE_* naming for pickup by filepro_sync.py
mv "$FPMERGE_FILE" "$DEST"
echo "Exported quote $QUOTE_NUM -> $DEST"
