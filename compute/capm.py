from __future__ import annotations

import math

import numpy as np
import pandas as pd

from data import yfinance_client as yfc


def monthly_returns(series: pd.Series) -> pd.Series:
    s = series.dropna()
    monthly = s.resample("ME").last()
    return monthly.pct_change().dropna()


def beta(ticker: str, market: str = "SPY", months: int = 60) -> float | None:
    px = yfc.prices(ticker, period=f"{max(months // 12 + 1, 6)}y")
    mx = yfc.prices(market, period=f"{max(months // 12 + 1, 6)}y")
    if px is None or px.empty or mx is None or mx.empty:
        return None
    px = px.set_index("date")["Close"].astype(float)
    mx = mx.set_index("date")["Close"].astype(float)
    rt = monthly_returns(px).tail(months)
    rm = monthly_returns(mx).tail(months)
    df = pd.concat([rt, rm], axis=1, join="inner").dropna()
    if len(df) < 24:
        return None
    cov = np.cov(df.iloc[:, 0], df.iloc[:, 1])[0, 1]
    var = np.var(df.iloc[:, 1])
    if var == 0:
        return None
    return float(cov / var)


def cost_of_equity(risk_free: float, beta_val: float, mrp: float,
                   country_risk: float = 0.0) -> float:
    return risk_free + beta_val * mrp + country_risk


def cost_of_debt(interest_expense: float, total_debt: float,
                 tax_rate: float = 0.21) -> float | None:
    if total_debt in (None, 0):
        return None
    pretax = interest_expense / total_debt
    return pretax * (1 - tax_rate)


def wacc(re: float, rd: float, equity_value: float, debt_value: float,
         tax_rate: float) -> float | None:
    total = (equity_value or 0) + (debt_value or 0)
    if total == 0:
        return None
    we = equity_value / total
    wd = debt_value / total
    return we * re + wd * rd * (1 - tax_rate)
