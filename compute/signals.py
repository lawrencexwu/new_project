from __future__ import annotations

import pandas as pd


def _ma(close: pd.Series, n: int) -> pd.Series:
    return close.rolling(n).mean()


def daily_buy_signal(close: pd.Series) -> bool:
    if len(close) < 22:
        return False
    ma10 = _ma(close, 10).iloc[-1]
    ma20 = _ma(close, 20).iloc[-1]
    last = close.iloc[-1]
    return bool(last > ma10 and last > ma20 and ma10 > ma20)


def weekly_buy_signal(close_daily: pd.Series) -> bool:
    if len(close_daily) < 110:
        return False
    weekly = close_daily.resample("W-FRI").last().dropna()
    if len(weekly) < 22:
        return False
    ma10 = _ma(weekly, 10).iloc[-1]
    ma20 = _ma(weekly, 20).iloc[-1]
    last = weekly.iloc[-1]
    return bool(last > ma10 and last > ma20 and ma10 > ma20)


def above_rising_5dma(close: pd.Series) -> bool:
    if len(close) < 8:
        return False
    ma5 = _ma(close, 5)
    return bool(close.iloc[-1] > ma5.iloc[-1] and ma5.iloc[-1] > ma5.iloc[-3])


def market_regime(close: pd.Series) -> str:
    if len(close) < 50:
        return "Unknown"
    ma20 = _ma(close, 20)
    ma50 = _ma(close, 50)
    last_20 = ma20.iloc[-20:]
    slope = (last_20.iloc[-1] - last_20.iloc[0]) / max(abs(last_20.iloc[0]), 1e-9)
    above_20 = close.iloc[-1] > ma20.iloc[-1]
    above_50 = close.iloc[-1] > ma50.iloc[-1]

    if slope > 0.02 and above_20 and above_50:
        return "Trend Up"
    if slope < -0.02 and not above_20 and not above_50:
        return "Trend Down"
    return "Chop"


def pct_change(close: pd.Series, lookback: int) -> float | None:
    if close is None or len(close) <= lookback:
        return None
    return float(close.iloc[-1] / close.iloc[-1 - lookback] - 1)


def pct_off_52w_high(close: pd.Series) -> float | None:
    if close is None or close.empty:
        return None
    window = close.tail(252)
    high = window.max()
    if high == 0:
        return None
    return float(close.iloc[-1] / high - 1)


def ytd_change(close_with_dates: pd.DataFrame) -> float | None:
    if close_with_dates is None or close_with_dates.empty:
        return None
    df = close_with_dates.copy()
    df["date"] = pd.to_datetime(df["date"])
    year = df["date"].dt.year.iloc[-1]
    yr = df[df["date"].dt.year == year]
    if yr.empty:
        return None
    return float(yr["Close"].iloc[-1] / yr["Close"].iloc[0] - 1)
