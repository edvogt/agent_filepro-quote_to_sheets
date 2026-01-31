#!/bin/bash
# fix_filepro_json.sh - Fixes malformed FilePro JSON exports
# Usage: ./fix_filepro_json.sh input.json [output.json]
#        If output.json is omitted, fixes in place (backup saved as .orig)
#
# Fixes:
#   1. Wraps loose line item objects in "line_items": [...] array
#   2. Adds missing commas in totals section
#   3. Fixes empty "Tax:" value
#   4. Normalizes spacing in numeric values

set -e

INPUT="$1"
OUTPUT="${2:-}"

if [[ -z "$INPUT" ]]; then
    echo "Usage: $0 input.json [output.json]" >&2
    exit 1
fi

if [[ ! -f "$INPUT" ]]; then
    echo "Error: File not found: $INPUT" >&2
    exit 1
fi

# Check if file is already valid JSON - skip fixing if so
if command -v jq &> /dev/null; then
    if jq empty "$INPUT" 2>/dev/null; then
        # Valid JSON - check if it has line_items array
        if jq -e '.line_items' "$INPUT" &>/dev/null; then
            echo "Skipped: $INPUT (already valid JSON with line_items)"
            exit 0
        fi
    fi
fi

# Use temp file for processing
TMPFILE=$(mktemp)
trap "rm -f $TMPFILE" EXIT

# AWK script handles the structural fix (wrapping line items in array)
# Then sed handles the simpler comma/value fixes
awk '
BEGIN { in_loose_items = 0; buffer = "" }

# Detect end of entry_details section
/^  "entry_details":/ { in_entry = 1 }

in_entry && /^  \},?$/ {
    print
    in_entry = 0
    in_loose_items = 1
    print ""
    print "  \"line_items\": ["
    next
}

# Detect start of totals section - close the line_items array
in_loose_items && /^  "totals":/ {
    print "  ],"
    print ""
    in_loose_items = 0
}

# Skip the orphaned ], line (we handle array closing above)
in_loose_items && /^\],?$/ {
    next
}

# Print everything else
{ print }
' "$INPUT" | \
sed -E '
    # Fix totals section - add commas after values missing them
    /"Sub Total[:]?"[:]?\s*[0-9.-]+$/ s/$/,/

    # Fix empty "Tax:" line - add null value and comma
    /"Tax":\s*$/ s/$/ null,/

    # Normalize spacing in numeric values (remove extra leading spaces)
    s/:\s{2,}([0-9])/: \1/g

    # Remove trailing whitespace
    s/[[:space:]]+$//
' | \
awk '
    # Remove trailing commas before array/object close brackets
    # Handle blank lines between comma and closing bracket
    {
        lines[NR] = $0
    }
    END {
        for (i = 1; i <= NR; i++) {
            line = lines[i]
            # If this line ends with comma, look ahead for closing bracket
            if (match(line, /,\s*$/)) {
                # Look ahead past blank lines for ] or }
                for (j = i + 1; j <= NR; j++) {
                    if (lines[j] ~ /^\s*$/) continue  # skip blank lines
                    if (lines[j] ~ /^\s*[\]\}]/) {
                        sub(/,\s*$/, "", line)  # remove trailing comma
                    }
                    break
                }
            }
            print line
        }
    }
' > "$TMPFILE"

# Validate the result is valid JSON (if jq is available)
if command -v jq &> /dev/null; then
    if ! jq empty "$TMPFILE" 2>/dev/null; then
        echo "Warning: Output may still have JSON issues. Manual review recommended." >&2
    fi
fi

# Output or replace
if [[ -z "$OUTPUT" ]]; then
    cp "$INPUT" "${INPUT}.orig"
    mv "$TMPFILE" "$INPUT"
    echo "Fixed: $INPUT (backup: ${INPUT}.orig)"
else
    mv "$TMPFILE" "$OUTPUT"
    echo "Fixed: $OUTPUT"
fi
