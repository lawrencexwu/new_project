from __future__ import annotations

from datetime import datetime, timedelta

import pandas as pd
import requests

import config
from data import cache

BASE = "https://api.stlouisfed.org/fred"

SERIES = {
    "DGS3MO": "3M Treasury",
    "DGS2": "2Y Treasury",
    "DGS5": "5Y Treasury",
    "DGS10": "10Y Treasury",
    "DGS30": "30Y Treasury",
    "T10Y2Y": "10Y-2Y Spread",
    "T10Y3M": "10Y-3M Spread",
    "DFF": "Fed Funds Effective",
    "CPIAUCSL": "CPI",
    "PPIACO": "PPI",
    "PAYEMS": "Nonfarm Payrolls",
    "UNRATE": "Unemployment",
    "GDPNOW": "GDP Nowcast (Atlanta Fed)",
}


def _api_key() -> str:
    k = config.get("fred_api_key", "")
    return k or ""


def series(code: str, force: bool = False) -> pd.DataFrame:
    key = f"fred_{code}"

    def fetch():
        api_key = _api_key()
        if not api_key:
            return pd.DataFrame()
        url = f"{BASE}/series/observations"
        params = {
            "series_id": code,
            "api_key": api_key,
            "file_type": "json",
            "observation_start": (datetime.utcnow() - timedelta(days=365 * 5)).strftime("%Y-%m-%d"),
        }
        r = requests.get(url, params=params, timeout=20)
        if r.status_code != 200:
            return pd.DataFrame()
        obs = r.json().get("observations", [])
        if not obs:
            return pd.DataFrame()
        df = pd.DataFrame(obs)
        df["date"] = pd.to_datetime(df["date"])
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        df["series"] = code
        return df[["date", "series", "value"]].dropna()

    return cache.get_or_fetch("treasury", key, fetch, force=force)


def latest(code: str) -> float | None:
    df = series(code)
    if df is None or df.empty:
        return None
    return float(df.iloc[-1]["value"])


def treasury_curve() -> pd.DataFrame:
    rows = []
    for code in ("DGS3MO", "DGS2", "DGS5", "DGS10", "DGS30"):
        rows.append({"series": code, "label": SERIES[code], "yield_pct": latest(code)})
    return pd.DataFrame(rows)


def risk_free_10y() -> float | None:
    """10Y Treasury yield in decimal (0.043 = 4.3%)."""
    v = latest("DGS10")
    return None if v is None else v / 100.0


def value_n_ago(code: str, n: int) -> float | None:
    """The observation value n steps back from the latest (n=1 → prior obs)."""
    df = series(code)
    if df is None or df.empty:
        return None
    df = df.dropna()
    if len(df) <= n:
        return None
    return float(df.iloc[-1 - n]["value"])


def latest_value(code: str) -> float | None:
    """Latest non-NaN observation."""
    df = series(code)
    if df is None or df.empty:
        return None
    df = df.dropna()
    if df.empty:
        return None
    return float(df.iloc[-1]["value"])


def yoy_change(code: str) -> float | None:
    """Year-over-year % change for a monthly index series (CPI, PPI).
    Returns a decimal (0.032 = 3.2%)."""
    df = series(code)
    if df is None or df.empty:
        return None
    df = df.dropna()
    if len(df) < 13:
        return None
    latest_v = float(df.iloc[-1]["value"])
    year_ago = float(df.iloc[-13]["value"])
    if year_ago == 0:
        return None
    return latest_v / year_ago - 1


def yoy_change_prev(code: str) -> float | None:
    """The YoY % change as of one month ago — for the 'previous' column."""
    df = series(code)
    if df is None or df.empty:
        return None
    df = df.dropna()
    if len(df) < 14:
        return None
    prev = float(df.iloc[-2]["value"])
    prev_year_ago = float(df.iloc[-14]["value"])
    if prev_year_ago == 0:
        return None
    return prev / prev_year_ago - 1


def monthly_change(code: str, n_back: int = 0) -> float | None:
    """Month-over-month change in level (for Nonfarm Payrolls). n_back=0 is
    the latest month's change, n_back=1 is the prior month's change."""
    df = series(code)
    if df is None or df.empty:
        return None
    df = df.dropna()
    if len(df) < 2 + n_back:
        return None
    cur = float(df.iloc[-1 - n_back]["value"])
    prev = float(df.iloc[-2 - n_back]["value"])
    return cur - prev


# FOMC meeting decision dates (the second/final day of each meeting).
# The Fed publishes these ~2 years ahead; update annually.
FOMC_DECISION_DATES = [
    "2026-01-28", "2026-03-18", "2026-04-29", "2026-06-17",
    "2026-07-29", "2026-09-16", "2026-10-28", "2026-12-09",
    "2027-01-27", "2027-03-17",
]


def next_fomc_date(today=None) -> str | None:
    """Return the next FOMC decision date (YYYY-MM-DD) on or after today."""
    import datetime as _dt
    if today is None:
        today = _dt.date.today()
    elif isinstance(today, str):
        today = _dt.date.fromisoformat(today)
    for d in FOMC_DECISION_DATES:
        if _dt.date.fromisoformat(d) >= today:
            return d
    return None
