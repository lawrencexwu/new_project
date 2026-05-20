"""CLI entry point: python refresh.py --market | python refresh.py NVDA"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import config
from compute import altman, capm, dcf, hillegeist, kmv, merton, ratios, signals
from data import edgar_client as edgar
from data import fred_client as fred
from data import yfinance_client as yfc


def refresh_market():
    print("Refreshing market-wide data...")
    universe = config.get("universe", {})
    for section, tickers in universe.items():
        if section == "macro_risk":
            continue
        for t in tickers:
            df = yfc.prices(t, period="2y", force=True)
            ok = "ok" if df is not None and not df.empty else "—"
            print(f"  {section}/{t}: {ok}")

    print("Refreshing Treasury curve...")
    curve = fred.treasury_curve()
    print(curve.to_string(index=False))


def refresh_ticker(ticker: str):
    print(f"Refreshing {ticker}...")
    ticker = ticker.upper()

    info = yfc.info(ticker, force=True)
    print(f"  name: {info.get('shortName', '?')}, industry: {info.get('industry', '?')}")

    px = yfc.prices(ticker, period="10y", force=True)
    if px is not None and not px.empty:
        print(f"  prices: {len(px)} rows, last close {px['Close'].iloc[-1]:.2f}")

    is_q = yfc.income_statement(ticker, quarterly=True, force=True)
    bs_q = yfc.balance_sheet(ticker, quarterly=True, force=True)
    cf_q = yfc.cashflow(ticker, quarterly=True, force=True)
    print(f"  IS quarters: {is_q.shape[1] - 1 if not is_q.empty else 0}")
    print(f"  BS quarters: {bs_q.shape[1] - 1 if not bs_q.empty else 0}")
    print(f"  CF quarters: {cf_q.shape[1] - 1 if not cf_q.empty else 0}")

    eps_est = yfc.earnings_estimate(ticker, force=True)
    rev_est = yfc.revenue_estimate(ticker, force=True)
    print(f"  EPS estimate rows: {len(eps_est)}, Rev estimate rows: {len(rev_est)}")

    rf = fred.risk_free_10y()
    b = capm.beta(ticker)
    print(f"  beta(60mo vs SPY): {b}")
    if rf is not None and b is not None:
        re = capm.cost_of_equity(rf, b, mrp=0.05)
        print(f"  Re (CAPM): {re:.3%}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("ticker", nargs="?", help="Ticker to refresh (e.g. NVDA)")
    p.add_argument("--market", action="store_true", help="Refresh market-wide data")
    args = p.parse_args()

    if args.market:
        refresh_market()
    elif args.ticker:
        refresh_ticker(args.ticker)
    else:
        p.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
