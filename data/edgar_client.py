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


# -----------------------------------------------------------
# yfinance-shaped statement assembly
# -----------------------------------------------------------

# Each line label maps to the list of US-GAAP concepts to try, in order.
# First non-empty wins.
_IS_CONCEPTS = {
    "Total Revenue": ["Revenues", "SalesRevenueNet",
                      "RevenueFromContractWithCustomerExcludingAssessedTax",
                      "SalesRevenueGoodsNet"],
    "Cost Of Revenue": ["CostOfRevenue", "CostOfGoodsAndServicesSold",
                        "CostOfGoodsSold"],
    "Gross Profit": ["GrossProfit"],
    "Selling General And Administration": [
        "SellingGeneralAndAdministrativeExpense",
        "GeneralAndAdministrativeExpense"],
    "Research And Development": ["ResearchAndDevelopmentExpense"],
    "Reconciled Depreciation": ["DepreciationDepletionAndAmortization",
                                "DepreciationAndAmortization",
                                "Depreciation"],
    "Operating Income": ["OperatingIncomeLoss"],
    "Interest Expense": ["InterestExpense"],
    "Pretax Income": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                      "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "Tax Provision": ["IncomeTaxExpenseBenefit"],
    "Net Income": ["NetIncomeLoss"],
}

_BS_CONCEPTS = {
    "Cash And Cash Equivalents": ["CashAndCashEquivalentsAtCarryingValue",
                                  "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
    "Accounts Receivable": ["AccountsReceivableNetCurrent",
                            "ReceivablesNetCurrent"],
    "Inventory": ["InventoryNet"],
    "Current Assets": ["AssetsCurrent"],
    "Total Assets": ["Assets"],
    "Current Liabilities": ["LiabilitiesCurrent"],
    "Long Term Debt": ["LongTermDebtNoncurrent", "LongTermDebt"],
    "Total Liabilities Net Minority Interest": ["Liabilities"],
    "Retained Earnings": ["RetainedEarningsAccumulatedDeficit"],
    "Common Stock Equity": ["StockholdersEquity",
                            "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
}

_CF_CONCEPTS = {
    "Operating Cash Flow": ["NetCashProvidedByUsedInOperatingActivities",
                            "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "Capital Expenditure": ["PaymentsToAcquirePropertyPlantAndEquipment",
                            "PaymentsToAcquireProductiveAssets"],
    "Cash Dividends Paid": ["PaymentsOfDividends",
                            "PaymentsOfDividendsCommonStock"],
    "Repurchase Of Capital Stock": ["PaymentsForRepurchaseOfCommonStock"],
}


def _quarterly_concept_values(facts: dict, concept_candidates: list[str]) -> dict:
    """For each unique quarter-end date, return the latest reported value
    from the first concept candidate that has any matching rows."""
    if not facts:
        return {}
    gaap = facts.get("facts", {}).get("us-gaap", {})
    for c in concept_candidates:
        if c not in gaap:
            continue
        unit_dict = gaap[c].get("units", {})
        unit = "USD" if "USD" in unit_dict else next(iter(unit_dict.keys()), None)
        if unit is None:
            continue
        records = unit_dict[unit]
        per_quarter: dict = {}
        for r in records:
            form = (r.get("form") or "").upper()
            if not form.startswith("10-Q") and not form.startswith("10-K"):
                continue
            end = pd.to_datetime(r.get("end"), errors="coerce")
            if pd.isna(end):
                continue
            # Skip annual rollups for income/cash-flow concepts (fp == "FY")
            fp = (r.get("fp") or "").upper()
            if c in _IS_CONCEPTS.get("Net Income", []) + _CF_CONCEPTS.get("Operating Cash Flow", []):
                if fp == "FY":
                    continue
            val = r.get("val")
            if val is None:
                continue
            # Newer filings win on identical dates
            existing = per_quarter.get(end)
            filed = pd.to_datetime(r.get("filed"), errors="coerce")
            if existing is None or filed > existing.get("filed", pd.Timestamp.min):
                per_quarter[end] = {"val": float(val), "filed": filed}
        if per_quarter:
            return {dt: v["val"] for dt, v in per_quarter.items()}
    return {}


def _statement_from_concepts(ticker: str, concept_map: dict,
                             force: bool = False) -> pd.DataFrame:
    facts = company_facts(ticker, force=force)
    if not facts:
        return pd.DataFrame()

    all_quarters: set = set()
    rows: dict[str, dict] = {}
    for label, concepts in concept_map.items():
        q_values = _quarterly_concept_values(facts, concepts)
        if q_values:
            rows[label] = q_values
            all_quarters.update(q_values.keys())

    if not rows:
        return pd.DataFrame()

    quarters = sorted(all_quarters)
    df = pd.DataFrame(index=list(rows.keys()), columns=quarters, dtype=float)
    for label, q_values in rows.items():
        for q, v in q_values.items():
            df.loc[label, q] = v
    return df.reset_index().rename(columns={"index": "line"})


def income_statement(ticker: str, force: bool = False) -> pd.DataFrame:
    return _statement_from_concepts(ticker, _IS_CONCEPTS, force=force)


def balance_sheet(ticker: str, force: bool = False) -> pd.DataFrame:
    return _statement_from_concepts(ticker, _BS_CONCEPTS, force=force)


def cashflow(ticker: str, force: bool = False) -> pd.DataFrame:
    return _statement_from_concepts(ticker, _CF_CONCEPTS, force=force)
