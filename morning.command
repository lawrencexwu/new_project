#!/bin/bash
#
# Morning routine — double-click from Finder to run.
#
# What this does:
#   1. cd into the project directory
#   2. Activate the Python venv (creates it on first run)
#   3. Pull latest code from the branch (only if you're using git)
#   4. Run `python populate.py --routine`
#      — refreshes Market_Daily.xlsx + every Ticker_*.xlsx in your Drive
#   5. Open Market_Daily.xlsx in whatever the system uses for .xlsx
#
# First-time setup:
#   chmod +x morning.command
#   Right-click → Open (one-time macOS Gatekeeper trust prompt)
#

set -e
cd "$(dirname "${BASH_SOURCE[0]}")"

echo "==> Morning routine starting at $(date '+%Y-%m-%d %H:%M')"

# Activate venv (create if missing)
if [ ! -d ".venv" ]; then
    echo "==> First run — creating Python venv..."
    python3 -m venv .venv
    .venv/bin/pip install -q -r requirements.txt
fi
source .venv/bin/activate

# Pull latest code if this is a git checkout (silent if no remote)
if [ -d ".git" ]; then
    git pull --quiet 2>/dev/null || true
fi

# Run the routine
python populate.py --routine

# Open the file
MARKET_PATH=$(python -c "import config; print(config.get('market_daily_path'))" 2>/dev/null || echo "")
if [ -n "$MARKET_PATH" ] && [ -f "$MARKET_PATH" ]; then
    echo "==> Opening Market_Daily.xlsx..."
    open "$MARKET_PATH"
else
    echo "==> Warning: market_daily_path not found in settings or file missing"
fi

echo "==> Done at $(date '+%H:%M:%S')"
echo ""
read -p "Press Enter to close this window..."
