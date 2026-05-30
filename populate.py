"""Populate workbooks with live data (headless via openpyxl).

Usage:
    python populate.py --market                  fills Market_Daily.xlsx
    python populate.py NVDA                      clones template → Ticker_NVDA.xlsx
    python populate.py NVDA AAPL MSFT            batch (one call per ticker)
    python populate.py --routine                 morning: --market + refresh every
                                                 existing Ticker_*.xlsx file
    python populate.py --watchlist               refresh every existing
                                                 Ticker_*.xlsx (no --market)
    python populate.py NVDA --force              bypass cache
    python populate.py --doctor                  setup diagnostic
    python populate.py --list-archive            list snapshot files
    python populate.py --snapshot NVDA           snapshot a ticker workbook
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
    p.add_argument("tickers", nargs="*", help="One or more tickers to populate")
    p.add_argument("--market", action="store_true",
                   help="Populate Market_Daily.xlsx")
    p.add_argument("--routine", action="store_true",
                   help="Morning routine: --market + refresh every existing Ticker_*.xlsx")
    p.add_argument("--watchlist", action="store_true",
                   help="Refresh every existing Ticker_*.xlsx file (no --market)")
    p.add_argument("--out", help="Output path (single-ticker mode only)")
    p.add_argument("--force", action="store_true",
                   help="Bypass cache and force fresh fetches")
    p.add_argument("--snapshot", help="Snapshot the named ticker workbook into Archive/")
    p.add_argument("--list-archive", action="store_true",
                   help="List the contents of the archive directory")
    p.add_argument("--doctor", action="store_true",
                   help="Diagnose setup: check paths, API keys, network")
    args = p.parse_args()

    if args.doctor:
        _doctor()
        return

    if args.list_archive:
        archive = config.get("archive_dir")
        if archive is None or not Path(archive).exists():
            archive = Path(__file__).resolve().parent / "build" / "Archive"
        archive = Path(archive)
        if not archive.exists():
            print(f"Archive dir {archive} is empty or does not exist.")
            return
        snapshots = sorted(archive.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not snapshots:
            print(f"No snapshots found in {archive}.")
            return
        print(f"Snapshots in {archive}:")
        for s in snapshots:
            from datetime import datetime
            ts = datetime.fromtimestamp(s.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            print(f"  {ts}  {s.name}")
        return

    if args.snapshot:
        wb_path = _tickers_dir() / f"Ticker_{args.snapshot.upper()}.xlsx"
        if not wb_path.exists():
            print(f"ERROR: {wb_path} does not exist.")
            sys.exit(1)
        out = populator.snapshot_to_archive(wb_path)
        print(f"Snapshot: {out}")
        return

    if args.routine:
        # Refresh any per-ticker deep-dive workbooks the user has created,
        # then the market dashboard (which scans the full watchlist.txt for
        # the Watchlist tab + Screener — no per-ticker workbook needed).
        _refresh_watchlist(args.force)
        _refresh_market(args.force)
        return

    if args.watchlist:
        _refresh_watchlist(args.force)
        return

    if args.market:
        _refresh_market(args.force)
        return

    if not args.tickers:
        p.print_help()
        sys.exit(1)

    template = _template_path()
    if not template.exists():
        print(f"ERROR: template {template} does not exist. Run `python -m xl.build_workbooks` first.")
        sys.exit(1)

    if args.out and len(args.tickers) > 1:
        print("ERROR: --out can only be used with a single ticker.")
        sys.exit(1)

    for tkr in args.tickers:
        out_path = (Path(args.out) if args.out
                    else _tickers_dir() / f"Ticker_{tkr.upper()}.xlsx")
        result = populator.populate_ticker(template, out_path, tkr,
                                            force=args.force)
        print(f"Populated {result}")


def _refresh_market(force: bool):
    path = _market_path()
    if not path.exists():
        print(f"ERROR: {path} does not exist. Run `python -m xl.build_workbooks` first.")
        sys.exit(1)
    out = populator.populate_market_daily(path, force=force)
    print(f"Populated {out}")


def _refresh_watchlist(force: bool):
    tdir = _tickers_dir()
    if not tdir.exists():
        print(f"ERROR: tickers dir {tdir} does not exist.")
        sys.exit(1)
    files = sorted(tdir.glob("Ticker_*.xlsx"))
    files = [f for f in files if f.name != "Ticker_TEMPLATE.xlsx"]
    if not files:
        print(f"No Ticker_*.xlsx files in {tdir}. Run `python populate.py NVDA` first.")
        return
    template = _template_path()
    print(f"Refreshing {len(files)} ticker workbooks...")
    for f in files:
        # Extract ticker from filename
        ticker = f.stem.replace("Ticker_", "")
        try:
            populator.populate_ticker(template, f, ticker, force=force)
            print(f"  {ticker}: OK")
        except Exception as e:
            print(f"  {ticker}: FAILED ({type(e).__name__}: {e})")


def _doctor():
    print("=" * 60)
    print("Markets & Stock Analysis Dashboard — setup check")
    print("=" * 60)
    settings = config.load()

    print("\n[Paths]")
    for key in ("drive_root", "cache_dir", "tickers_dir", "archive_dir",
                "market_daily_path", "template_path"):
        path = settings.get(key)
        if path is None:
            print(f"  {key:<22} (unset)")
            continue
        path = Path(path)
        exists = "OK " if path.exists() else "MISSING"
        print(f"  {key:<22} {exists}  {path}")

    print("\n[Credentials]")
    fred = bool(settings.get("fred_api_key"))
    print(f"  FRED API key          {'set' if fred else 'NOT set (Macro tab will be empty)'}")
    sec = settings.get("sec_user_agent") or ""
    if sec and "example.com" not in sec:
        print(f"  SEC user-agent        set ({sec})")
    else:
        print(f"  SEC user-agent        NOT set (EDGAR will be denied — required by SEC)")

    print("\n[Templates]")
    build_dir = Path(__file__).resolve().parent / "build"
    checks = [
        ("Market_Daily.xlsx", settings.get("market_daily_path"),
         build_dir / "Market_Daily.xlsx"),
        ("Ticker_TEMPLATE.xlsx", settings.get("template_path"),
         build_dir / "Ticker_TEMPLATE.xlsx"),
    ]
    for name, drive_path, build_path in checks:
        if drive_path is not None and Path(drive_path).exists():
            print(f"  {name:<22} found in Drive  {drive_path}")
        elif build_path.exists():
            print(f"  {name:<22} found in build/ (move to Drive)  {build_path}")
        else:
            print(f"  {name:<22} missing (run `python -m xl.build_workbooks`)")

    print("\n[Network checks]")
    try:
        from data import fred_client as fc
        v = fc.latest("DGS10") if fred else None
        print(f"  FRED                  {'OK' if v is not None else 'no data (key issue)'}")
    except Exception as e:
        print(f"  FRED                  ERROR: {e}")

    try:
        from data import edgar_client as ec
        cik = ec.ticker_to_cik("AAPL")
        print(f"  SEC EDGAR             {'OK' if cik else 'no response'}")
    except Exception as e:
        print(f"  SEC EDGAR             ERROR: {e}")

    try:
        from data import yfinance_client as yc
        px = yc.prices("SPY", period="5d")
        print(f"  yfinance              {'OK' if not px.empty else 'no data'}")
    except Exception as e:
        print(f"  yfinance              ERROR: {e}")

    print()


if __name__ == "__main__":
    main()
