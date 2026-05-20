from __future__ import annotations

import pandas as pd
import requests

from data import cache

SHORT_INTEREST_BASE = "https://api.finra.org/data/group/otcMarket/name/consolidatedShortInterest"


def short_interest(ticker: str, force: bool = False) -> pd.DataFrame:
    """Latest published short interest. Returns empty if unavailable."""
    key = f"{ticker}_short_int"

    def fetch():
        try:
            params = {
                "limit": 24,
                "compareFilters": f"symbolCode:{ticker.upper()}",
            }
            r = requests.get(SHORT_INTEREST_BASE, params=params, timeout=20)
            if r.status_code != 200:
                return pd.DataFrame()
            data = r.json()
            if not data:
                return pd.DataFrame()
            df = pd.DataFrame(data)
            if "settlementDate" in df.columns:
                df["settlementDate"] = pd.to_datetime(df["settlementDate"])
                df = df.sort_values("settlementDate")
            return df
        except Exception:
            return pd.DataFrame()

    return cache.get_or_fetch("shortinterest", key, fetch, force=force)


def short_interest_pct_float(ticker: str, shares_float: float | None) -> float | None:
    df = short_interest(ticker)
    if df is None or df.empty or shares_float in (None, 0):
        return None
    col = "currentShortPositionQuantity" if "currentShortPositionQuantity" in df.columns else None
    if not col:
        return None
    latest = float(df.iloc[-1][col])
    return latest / float(shares_float)
