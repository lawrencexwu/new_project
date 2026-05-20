"""Populate workbooks with live data (headless via openpyxl).

Usage:
    python populate.py --market                  # fills Market_Daily.xlsx
    python populate.py NVDA                      # clones template → Ticker_NVDA.xlsx, fills it
    python populate.py NVDA --out path.xlsx      # custom output path
    python populate.py NVDA --force              # bypass cache
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import config
from xl import populator


def _build_dir() -> Path:
    return Path(__file__).resolve().parent / "build"


def _market_path() -> Path:
    configured = config.get("market_daily_path")
    if configured and configured.exists():
        return configured
    return _build_dir() / "Market_Daily.xlsx"


def _template_path() -> Path:
    configured = config.get("template_path")
    if configured and configured.exists():
        return configured
    return _build_dir() / "Ticker_TEMPLATE.xlsx"


def _tickers_dir() -> Path:
    configured = config.get("tickers_dir")
    if configured and configured.exists():
        return configured
    return _build_dir()


def main():
    p = argparse.ArgumentParser(description="Populate workbooks with live data")
    p.add_argument("ticker", nargs="?", help="Ticker to populate")
    p.add_argument("--market", action="store_true",
                   help="Populate Market_Daily.xlsx")
    p.add_argument("--out", help="Output path (ticker mode)")
    p.add_argument("--force", action="store_true",
                   help="Bypass cache and force fresh fetches")
    args = p.parse_args()

    if args.market:
        path = _market_path()
        if not path.exists():
            print(f"ERROR: {path} does not exist. Run `python -m xl.build_workbooks` first.")
            sys.exit(1)
        out = populator.populate_market_daily(path, force=args.force)
        print(f"Populated {out}")
        return

    if not args.ticker:
        p.print_help()
        sys.exit(1)

    template = _template_path()
    if not template.exists():
        print(f"ERROR: template {template} does not exist. Run `python -m xl.build_workbooks` first.")
        sys.exit(1)

    out_path = Path(args.out) if args.out else _tickers_dir() / f"Ticker_{args.ticker.upper()}.xlsx"
    result = populator.populate_ticker(template, out_path, args.ticker, force=args.force)
    print(f"Populated {result}")


if __name__ == "__main__":
    main()
