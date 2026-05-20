from __future__ import annotations

from datetime import datetime, timedelta
from typing import Iterable

import pandas as pd
import yfinance as yf

from data import cache


def prices(ticker: str, period: str = "10y", interval: str = "1d",
           force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_{period}_{interval}"

    def fetch():
        df = yf.Ticker(ticker).history(period=period, interval=interval,
                                       auto_adjust=False)
        if df.empty:
            return df
        df = df.reset_index().rename(columns={"Date": "date"})
        df["ticker"] = ticker
        return df

    return cache.get_or_fetch("prices", key, fetch, force=force)


def batch_prices(tickers: Iterable[str], period: str = "1y",
                 force: bool = False) -> dict[str, pd.DataFrame]:
    return {t: prices(t, period=period, force=force) for t in tickers}


def info(ticker: str, force: bool = False) -> dict:
    key = f"{ticker}_info"

    def fetch():
        d = yf.Ticker(ticker).info or {}
        return pd.DataFrame([{"k": k, "v": str(v)} for k, v in d.items()])

    df = cache.get_or_fetch("prices", key, fetch, force=force)
    if df is None or df.empty:
        return {}
    return dict(zip(df["k"], df["v"]))


def income_statement(ticker: str, quarterly: bool = True,
                     force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_is_{'q' if quarterly else 'a'}"

    def fetch():
        t = yf.Ticker(ticker)
        df = t.quarterly_income_stmt if quarterly else t.income_stmt
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index().rename(columns={"index": "line"})

    return cache.get_or_fetch("statements", key, fetch, force=force)


def balance_sheet(ticker: str, quarterly: bool = True,
                  force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_bs_{'q' if quarterly else 'a'}"

    def fetch():
        t = yf.Ticker(ticker)
        df = t.quarterly_balance_sheet if quarterly else t.balance_sheet
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index().rename(columns={"index": "line"})

    return cache.get_or_fetch("statements", key, fetch, force=force)


def cashflow(ticker: str, quarterly: bool = True,
             force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_cf_{'q' if quarterly else 'a'}"

    def fetch():
        t = yf.Ticker(ticker)
        df = t.quarterly_cashflow if quarterly else t.cashflow
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index().rename(columns={"index": "line"})

    return cache.get_or_fetch("statements", key, fetch, force=force)


def earnings_estimate(ticker: str, force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_eps_est"

    def fetch():
        try:
            df = yf.Ticker(ticker).earnings_estimate
        except Exception:
            return pd.DataFrame()
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index()

    return cache.get_or_fetch("consensus", key, fetch, force=force)


def revenue_estimate(ticker: str, force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_rev_est"

    def fetch():
        try:
            df = yf.Ticker(ticker).revenue_estimate
        except Exception:
            return pd.DataFrame()
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index()

    return cache.get_or_fetch("consensus", key, fetch, force=force)


def eps_revisions(ticker: str, force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_eps_rev"

    def fetch():
        try:
            df = yf.Ticker(ticker).eps_revisions
        except Exception:
            return pd.DataFrame()
        if df is None or df.empty:
            return pd.DataFrame()
        return df.reset_index()

    return cache.get_or_fetch("consensus", key, fetch, force=force)


def calendar(ticker: str, force: bool = False) -> dict:
    key = f"{ticker}_calendar"

    def fetch():
        try:
            c = yf.Ticker(ticker).calendar
        except Exception:
            return pd.DataFrame()
        if c is None:
            return pd.DataFrame()
        if isinstance(c, dict):
            return pd.DataFrame([{"k": k, "v": str(v)} for k, v in c.items()])
        return c.reset_index() if hasattr(c, "reset_index") else pd.DataFrame()

    df = cache.get_or_fetch("consensus", key, fetch, force=force)
    if df is None or df.empty:
        return {}
    if "k" in df.columns and "v" in df.columns:
        return dict(zip(df["k"], df["v"]))
    return df.iloc[0].to_dict() if len(df) else {}


def insider_transactions(ticker: str, force: bool = False) -> pd.DataFrame:
    key = f"{ticker}_insider"

    def fetch():
        try:
            df = yf.Ticker(ticker).insider_transactions
        except Exception:
            return pd.DataFrame()
        return df if df is not None else pd.DataFrame()

    return cache.get_or_fetch("insider", key, fetch, force=force)


def options_atm_iv(ticker: str, force: bool = False) -> float | None:
    """ATM implied vol for the nearest-month expiry; None if unavailable."""
    try:
        t = yf.Ticker(ticker)
        expiries = t.options
        if not expiries:
            return None
        chain = t.option_chain(expiries[0])
        spot = t.history(period="1d")["Close"].iloc[-1]
        calls = chain.calls
        if calls is None or calls.empty:
            return None
        calls = calls.assign(dist=(calls["strike"] - spot).abs())
        return float(calls.sort_values("dist").iloc[0]["impliedVolatility"])
    except Exception:
        return None


def historical_vol(ticker: str, window_days: int = 30) -> float | None:
    """Annualized historical vol of daily log returns over a trailing window."""
    df = prices(ticker, period="1y")
    if df is None or df.empty:
        return None
    close = df["Close"].astype(float)
    rets = (close / close.shift(1)).apply(lambda x: None if x is None else __import__("math").log(x))
    rets = rets.dropna().tail(window_days)
    if len(rets) < 5:
        return None
    import numpy as np
    return float(rets.std() * np.sqrt(252))
