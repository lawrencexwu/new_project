from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass
class MultipleBand:
    name: str
    current: float | None
    median_3y: float | None
    median_5y: float | None
    median_10y: float | None
    min_5y: float | None
    max_5y: float | None


def _safe_div(num, denom):
    if num is None or denom in (None, 0):
        return None
    return num / denom


def pe(price: float, eps: float) -> float | None:
    return _safe_div(price, eps)


def ev_ebitda(ev: float, ebitda: float) -> float | None:
    return _safe_div(ev, ebitda)


def ev_sales(ev: float, sales: float) -> float | None:
    return _safe_div(ev, sales)


def p_book(price: float, book_per_share: float) -> float | None:
    return _safe_div(price, book_per_share)


def fcf_yield(fcf: float, market_cap: float) -> float | None:
    return _safe_div(fcf, market_cap)


def band_from_series(name: str, current: float | None,
                     historical: pd.Series) -> MultipleBand:
    if historical is None or historical.empty:
        return MultipleBand(name=name, current=current,
                            median_3y=None, median_5y=None, median_10y=None,
                            min_5y=None, max_5y=None)
    s = historical.dropna()
    tail3 = s.tail(12)   # quarterly
    tail5 = s.tail(20)
    tail10 = s.tail(40)
    return MultipleBand(
        name=name,
        current=current,
        median_3y=float(tail3.median()) if len(tail3) else None,
        median_5y=float(tail5.median()) if len(tail5) else None,
        median_10y=float(tail10.median()) if len(tail10) else None,
        min_5y=float(tail5.min()) if len(tail5) else None,
        max_5y=float(tail5.max()) if len(tail5) else None,
    )
