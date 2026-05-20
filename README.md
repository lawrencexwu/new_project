# Personal Markets & Stock Analysis Dashboard

Two-workbook system that drives a daily market workflow and per-ticker
fundamental + valuation + credit analysis, backed by a Python data layer
(yfinance, SEC EDGAR, FRED, FINRA) and rendered into Excel via openpyxl
and xlwings.

## Layout

```
Market_Daily.xlsx        ← opened every morning
  Daily Plan
  Breadth
  Macro
  Screener
  Journal Log
  Settings

Ticker_TEMPLATE.xlsx     ← cloned per ticker
  Cover
  Market
  Fin Stat
  Analysis
  Valuation
  Credit
  Trade Notes
```

## Setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp settings.example.json settings.local.json
# edit settings.local.json: drive_root, FRED API key, SEC user-agent
```

Build the templates:

```bash
python -m xl.build_workbooks
# produces build/Market_Daily.xlsx and build/Ticker_TEMPLATE.xlsx
```

Copy them to your Drive path and start working.

## CLI usage

```bash
python refresh.py --market      # warm cache for the sector universe + Treasury curve
python refresh.py NVDA          # pull prices / statements / consensus for one ticker
```

## xlwings buttons

`xl/macros.py` exposes entry points that can be wired to ActiveX buttons:

- `refresh_market_sectors` — populates sector heatmaps on Daily Plan
- `refresh_market_signals` — fills Section 1 trend booleans + Market Regime
- `refresh_macro` — fills the Macro tab from FRED
- `snapshot_to_journal` — appends current Daily Plan state to Journal Log

## Tests

```bash
python -m pytest tests/
```
