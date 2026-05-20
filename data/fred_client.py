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
