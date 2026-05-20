"""xlwings macro entry points. Called from buttons in the workbooks."""
from __future__ import annotations

from datetime import datetime

import xlwings as xw

import config
from data import yfinance_client as yfc
from data import fred_client as fc
from compute import signals


def _open_wb_for_macro():
    return xw.Book.caller()


def refresh_market_sectors():
    """Populate sector heatmaps on the Daily Plan sheet."""
    wb = _open_wb_for_macro()
    ws = wb.sheets["Daily Plan"]
    universe = config.get("universe", {})

    rows_by_section = {
        "S&P 500 Sub-Market Performance": universe.get("sub_market", []),
        "S&P 500 Sub-Sector Performance (Cap-Weighted)": universe.get("sub_sector_cw", []),
        "S&P 500 Sub-Sector Performance (Equal-Weighted)": universe.get("sub_sector_ew", []),
        "Top Thematic Sectors": universe.get("thematic", []),
    }

    spy_5d = signals.pct_change(_close(yfc.prices("SPY", period="6mo")), 5) or 0.0

    for section, tickers in rows_by_section.items():
        anchor = _find_section_anchor(ws, section)
        if anchor is None:
            continue
        for i, tkr in enumerate(tickers):
            row = anchor + 2 + i  # +2 for section header + col header
            close = _close(yfc.prices(tkr, period="2y"))
            ws.cells(row, 11).value = signals.pct_change(close, 1)
            ws.cells(row, 12).value = signals.pct_change(close, 5)
            ws.cells(row, 13).value = signals.pct_off_52w_high(close)
            ws.cells(row, 14).value = signals.ytd_change(yfc.prices(tkr, period="1y"))
            five_d = signals.pct_change(close, 5)
            ws.cells(row, 15).value = (five_d - spy_5d) if (five_d is not None) else None


def refresh_market_signals():
    """Populate booleans + regime in Section 1 of Daily Plan."""
    wb = _open_wb_for_macro()
    ws = wb.sheets["Daily Plan"]
    close = _close(yfc.prices("SPY", period="2y"))
    if close is None or close.empty:
        return
    daily = signals.daily_buy_signal(close)
    weekly = signals.weekly_buy_signal(close)
    above5 = signals.above_rising_5dma(close)
    regime = signals.market_regime(close)

    # Find rows under "1. Market Trend"
    anchor = _find_section_anchor(ws, "1. Market Trend")
    if anchor is None:
        return
    ws.cells(anchor + 1, 2).value = "YES" if daily else "NO"
    ws.cells(anchor + 2, 2).value = "YES" if weekly else "NO"
    ws.cells(anchor + 3, 2).value = "YES" if above5 else "NO"
    # +4 / +5 are user-entered ("recent trade traction", "52w trending")
    ws.cells(anchor + 6, 2).value = regime


def refresh_macro():
    """Populate Macro tab from FRED."""
    wb = _open_wb_for_macro()
    ws = wb.sheets["Macro"]
    mapping = {"3M": "DGS3MO", "2Y": "DGS2", "5Y": "DGS5",
               "10Y": "DGS10", "30Y": "DGS30"}
    for label, code in mapping.items():
        row = _find_label_row(ws, label)
        if row:
            v = fc.latest(code)
            ws.cells(row, 2).value = (v / 100.0) if v is not None else None

    for label, code in (("2s10s", "T10Y2Y"), ("3M10Y", "T10Y3M"),
                        ("Fed Funds Effective", "DFF"),
                        ("Unemployment", "UNRATE")):
        row = _find_label_row(ws, label)
        if row:
            v = fc.latest(code)
            ws.cells(row, 2).value = (v / 100.0) if v is not None else None


def snapshot_to_journal():
    """Append current Daily Plan state to Journal Log."""
    wb = _open_wb_for_macro()
    plan = wb.sheets["Daily Plan"]
    log = wb.sheets["Journal Log"]

    next_row = log.range("A" + str(log.cells.last_cell.row)).end("up").row + 1
    today = datetime.utcnow().strftime("%Y-%m-%d")

    regime_row = _find_section_anchor(plan, "1. Market Trend")
    regime = plan.cells(regime_row + 6, 2).value if regime_row else ""

    log.cells(next_row, 1).value = today
    log.cells(next_row, 2).value = regime
    log.cells(next_row, 3).value = ""
    sit_row = _find_section_anchor(plan, "3. Situational Awareness")
    emo_row = _find_section_anchor(plan, "4. Emotional Analysis")
    log.cells(next_row, 4).value = plan.cells(sit_row + 1, 1).value if sit_row else ""
    log.cells(next_row, 5).value = plan.cells(emo_row + 1, 1).value if emo_row else ""


def _close(df):
    if df is None or df.empty:
        return None
    df = df.copy()
    df["date"] = __import__("pandas").to_datetime(df["date"])
    return df.set_index("date")["Close"].astype(float)


def _find_section_anchor(ws, section_label: str) -> int | None:
    for cell in ws.range("A1:A300"):
        if cell.value == section_label:
            return cell.row
    for cell in ws.range("I1:I300"):
        if cell.value == section_label:
            return cell.row
    return None


def _find_label_row(ws, label: str) -> int | None:
    for cell in ws.range("A1:A100"):
        if cell.value == label:
            return cell.row
    return None
