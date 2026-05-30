#!/bin/bash
#
# Ad-hoc ticker analysis — double-click from Finder, get a prompt for
# ticker(s), populate them, open the first one.
#

set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
source .venv/bin/activate

# Use AppleScript so the prompt is a proper GUI dialog instead of a
# scary-looking Terminal prompt. Fall back to read if osascript fails.
if command -v osascript >/dev/null 2>&1; then
    TICKERS=$(osascript -e 'tell application "System Events" to display dialog "Enter one or more tickers (space-separated):" default answer "NVDA" with title "Ticker Analysis"' \
        -e 'return text returned of result' 2>/dev/null || true)
else
    read -p "Ticker(s) [NVDA]: " TICKERS
    TICKERS=${TICKERS:-NVDA}
fi

if [ -z "$TICKERS" ]; then
    echo "No ticker entered, aborting."
    exit 0
fi

echo "==> Populating: $TICKERS"
python populate.py $TICKERS

# Open the first ticker file
FIRST=$(echo "$TICKERS" | awk '{print toupper($1)}')
TICKERS_DIR=$(python -c "import config; print(config.get('tickers_dir'))" 2>/dev/null || echo "")
if [ -n "$TICKERS_DIR" ] && [ -f "$TICKERS_DIR/Ticker_$FIRST.xlsx" ]; then
    open "$TICKERS_DIR/Ticker_$FIRST.xlsx"
fi

echo "==> Done."
read -p "Press Enter to close this window..."
