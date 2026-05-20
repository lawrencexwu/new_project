from __future__ import annotations

import requests
import pandas as pd

import config
from data import cache

BASE = "https://data.sec.gov"
TICKER_MAP_URL = "https://www.sec.gov/files/company_tickers.json"


def _headers() -> dict:
    ua = config.get("sec_user_agent", "Personal Markets Dashboard contact@example.com")
    return {"User-Agent": ua, "Accept-Encoding": "gzip, deflate"}


def ticker_to_cik(ticker: str) -> str | None:
    df = cache.read("statements", "sec_ticker_map")
    if df is None or not cache.is_fresh("statements", "sec_ticker_map"):
        r = requests.get(TICKER_MAP_URL, headers=_headers(), timeout=20)
        r.raise_for_status()
        records = r.json()
        df = pd.DataFrame.from_dict(records, orient="index")
        df["ticker"] = df["ticker"].str.upper()
        df["cik_str"] = df["cik_str"].astype(int).astype(str).str.zfill(10)
        cache.write("statements", "sec_ticker_map", df)

    row = df[df["ticker"] == ticker.upper()]
    if row.empty:
        return None
    return row.iloc[0]["cik_str"]


def company_facts(ticker: str, force: bool = False) -> dict | None:
    cik = ticker_to_cik(ticker)
    if cik is None:
        return None
    key = f"{ticker}_facts_meta"

    def fetch():
        url = f"{BASE}/api/xbrl/companyfacts/CIK{cik}.json"
        r = requests.get(url, headers=_headers(), timeout=30)
        if r.status_code != 200:
            return pd.DataFrame()
        j = r.json()
        cache.write("statements", f"{ticker}_facts", pd.DataFrame([{"j": __import__("json").dumps(j)}]))
        return pd.DataFrame([{"cik": cik, "entityName": j.get("entityName", "")}])

    cache.get_or_fetch("statements", key, fetch, force=force)
    blob = cache.read("statements", f"{ticker}_facts")
    if blob is None or blob.empty:
        return None
    import json as _json
    return _json.loads(blob.iloc[0]["j"])


def concept_series(ticker: str, concept: str, unit: str = "USD") -> pd.DataFrame:
    """Pull a single XBRL concept time-series (e.g. 'Revenues', 'NetIncomeLoss')."""
    facts = company_facts(ticker)
    if not facts:
        return pd.DataFrame()
    try:
        units = facts["facts"]["us-gaap"][concept]["units"]
        if unit not in units:
            unit = next(iter(units.keys()))
        records = units[unit]
        df = pd.DataFrame(records)
        if df.empty:
            return df
        df["end"] = pd.to_datetime(df["end"])
        df["concept"] = concept
        return df.sort_values("end")
    except KeyError:
        return pd.DataFrame()
