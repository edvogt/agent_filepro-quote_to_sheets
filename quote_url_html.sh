#!/bin/bash
# quote_url_html.sh - Output Google Sheets links in HTML format
#
# Usage:
#   ./quote_url_html.sh              # Latest link
#   ./quote_url_html.sh 86016        # Specific quote number
#   ./quote_url_html.sh all          # All links as HTML list
#   ./quote_url_html.sh last 5       # Last 5 links

URL_LOG="/home/filepro/quote_urls.log"

if [[ ! -f "$URL_LOG" ]]; then
    echo "Error: URL log not found: $URL_LOG" >&2
    exit 1
fi

case "$1" in
    all)
        echo "<ul>"
        while IFS='|' read -r ts quote url; do
            url="${url## }"  # trim leading space
            quote="${quote## }"
            echo "  <li><a href=\"${url}\">${quote}</a></li>"
        done < "$URL_LOG"
        echo "</ul>"
        ;;
    last)
        count="${2:-5}"
        echo "<ul>"
        tail -n "$count" "$URL_LOG" | while IFS='|' read -r ts quote url; do
            url="${url## }"
            quote="${quote## }"
            echo "  <li><a href=\"${url}\">${quote}</a></li>"
        done
        echo "</ul>"
        ;;
    ""|latest)
        tail -1 "$URL_LOG" | awk -F' \\| ' '{
            gsub(/^ +/, "", $2);
            gsub(/^ +/, "", $3);
            print "<a href=\""$3"\">"$2"</a>"
        }'
        ;;
    *)
        # Assume it's a quote number
        grep -i "Quote $1" "$URL_LOG" | tail -1 | awk -F' \\| ' '{
            gsub(/^ +/, "", $2);
            gsub(/^ +/, "", $3);
            if ($3 != "") {
                print "<a href=\""$3"\">"$2"</a>"
            } else {
                print "Quote not found" > "/dev/stderr"
                exit 1
            }
        }'
        ;;
esac
