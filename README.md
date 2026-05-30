# Macro Dashboard

Two-workbook system that drives a daily market workflow and per-ticker
fundamental + valuation + credit analysis, backed by a Python data layer
(yfinance, SEC EDGAR, FRED, FINRA) and rendered into Excel via openpyxl
and xlwings.

## What you get

```
Market_Daily.xlsx        ← opened every morning
  Daily Plan      market-trend signals, watchlist & themes,
                  situational + emotional journal, sector heatmaps
                  (sub-market, cap-weighted, equal-weighted, thematic,
                  custom watchlist), macro risk strip
  Breadth         market internals
  Macro           Treasury curve, inflation, jobs, GDP nowcast (FRED)
  Screener        auto-flag custom-watchlist names with double-buy signal
  Journal Log     append-only daily entries
  Lessons         append-only post-mortems (Date · Ticker · Action ·
                  Reasoning · Outcome · Lesson)
  Positions       holdings + auto-computed P&L + pairwise correlation
                  matrix (90d daily returns, red/white/blue scale)
  Summary         aggregates every Ticker_*.xlsx Cover sheet into one
                  ranked table (Upside · MoS · Quality · EDF · Altman ·
                  Next Earnings · 3M EPS Revision) with color scales
  Earnings Calendar  next 60 days for watchlist + positions + summary
  Settings

Ticker_<SYM>.xlsx        ← cloned per ticker
  Cover           headline tiles (price · DCF FV · upside · MoS ·
                  quality score · EDF · Altman Z · next earnings ·
                  insider net · short interest · 3M EPS revision ·
                  Form 4 filing count) +
                  auto-flagged red flags + thesis/risks/pre-mortem
  Market          live quote, vol (HV30/90/365 + ATM IV), CAPM block,
                  debt market block
  Fin Stat        IS / BS / CFS — quarterly, 11 quarters
  Analysis        per-quarter ratios + 0-10 scores + DuPont 5-step +
                  Owner Earnings + Capital Allocation scorecard
                  (CAPEX/Buybacks/Dividends/M&A/Debt Paydown % of OCF) +
                  multiples band (P/E · EV/EBITDA · EV/Sales · P/B ·
                  FCF Yield: current vs 3y/5y/10y) + 4 native charts
                  (EBIT · Revenue · Net Income · Pretax)
  Valuation       5-year UFCF forecast (from consensus revenue +
                  historical ratios), EV bridge, sensitivity grids
                  (discount-rate × terminal-multiple), reverse DCF,
                  margin of safety
  Credit          KMV (primary, scipy-solver) + Altman Z (cross-check)
                  + collapsed Merton + Hillegeist
  Trade Notes     entry/exit, stop, pre-mortem, insider, short int
```

## Setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp settings.example.json settings.local.json
# edit settings.local.json:
#   drive_root, tickers_dir, archive_dir, market_daily_path, template_path
#   fred_api_key (free at https://fredaccount.stlouisfed.org/apikey)
#   sec_user_agent ("Your Name your@email.com" — required by SEC)
```

Generate the templates:

```bash
python -m xl.build_workbooks
# produces build/Market_Daily.xlsx and build/Ticker_TEMPLATE.xlsx
# move them to your Drive paths if not already pointed there
```

## CLI usage

```bash
# Morning routine — refresh Market_Daily + every existing Ticker_*.xlsx
python populate.py --routine

# Just refresh the market dashboard (sector heatmaps + signals + macro)
python populate.py --market

# Refresh just your watchlist (every Ticker_*.xlsx in Drive)
python populate.py --watchlist

# Populate one or more tickers
python populate.py NVDA
python populate.py NVDA AAPL MSFT GOOGL    # batch

# Bypass cache and force a fresh fetch
python populate.py NVDA --force

# Custom output path (single ticker only)
python populate.py NVDA --out ~/Desktop/nvda-snapshot.xlsx

# Take a timestamped snapshot of an existing ticker workbook
python populate.py --snapshot NVDA
# → Archive/Ticker_NVDA_20260520_1430.xlsx

# List all snapshots in the archive (newest first)
python populate.py --list-archive

# Diagnose setup: check paths, API keys, network reachability
python populate.py --doctor
```

## xlwings macros (in-place refresh from inside Excel)

`xl/macros.py` exposes entry points you can wire to buttons via Excel's
Form Controls (Developer → Insert → Button → Assign Macro):

| Macro                  | Where         | What it does                          |
|------------------------|---------------|---------------------------------------|
| `refresh_market`       | Market_Daily  | save → reload from yfinance + FRED    |
| `refresh_ticker`       | Ticker_<SYM>  | reads ticker from `Cover!B3`, refreshes |
| `snapshot_to_journal`  | Market_Daily  | appends Daily Plan state to Journal Log |

To set up:
1. Save the workbook as `.xlsm` (macro-enabled) if you want VBA buttons,
   or keep `.xlsx` and call macros via `xlwings runpython`.
2. In Excel: Developer → Macro Security → enable trusted access to VBA
   project model.
3. From a terminal in the project root, run `xlwings addin install`
   once.
4. In Excel: Add-Ins → xlwings → Configure → set "Interpreter" and
   "PYTHONPATH" to point at this repo's `.venv` and `code/` directory.
5. Add buttons (Developer → Insert → Form Control button) and assign
   them to `xl.macros.refresh_market`, etc.

## Daily morning workflow — ergonomic options

In increasing order of automation:

### 1. Terminal one-liner

```bash
python populate.py --routine
```

### 2. Double-click `morning.command` (Mac) / `morning.bat` (Windows)

Both ship in the repo root. They:
1. Activate the venv (creating it on first run)
2. `git pull` for any updates
3. Run `--routine`
4. Open Market_Daily.xlsx automatically when done

**First-time setup (Mac only):**
```bash
chmod +x morning.command analyze.command
```
Then right-click `morning.command` in Finder → Open. macOS will ask
once whether to trust an unsigned script — say yes. From then on,
double-click works normally.

### 3. Set a watchlist in settings, double-click does everything

Edit `settings.local.json`:

```json
"watchlist": ["NVDA", "AAPL", "MSFT", "GOOGL", "AMZN"],
```

Now `python populate.py --routine` (or the morning.command double-click):
- Auto-creates a `Ticker_<SYM>.xlsx` for any watchlist ticker that
  doesn't have one yet
- Refreshes every existing Ticker file
- Auto-fills the Daily Plan → Custom Watchlist heatmap rows with your
  watchlist tickers and their %D / %5D / %YTD / Off-52w / RS-SPY / sparkline
- Refreshes Market_Daily so Summary tab ranks your watchlist

You never type a ticker again.

### 4. Scheduled — runs even before you open the laptop

Copy `scripts/com.investing.morning-routine.plist.example` to
`~/Library/LaunchAgents/com.investing.morning-routine.plist`, edit the
two paths near the top, then:
```bash
launchctl load ~/Library/LaunchAgents/com.investing.morning-routine.plist
```

Default schedule: weekdays at 7:00 AM. Logs to `/tmp/morning-routine.log`.

### Ad-hoc ticker analysis

Double-click `analyze.command` → enter ticker(s) in the dialog →
workbook opens when ready. Or `python populate.py NVDA AAPL MSFT`.

---

After refresh, open Market_Daily, glance at Market Regime + sector
heatmaps + Summary (your ranked watchlist), read your prior journal
entry, write today's situational + emotional notes.

**Ticker analysis (per name):**
```bash
python populate.py NEW_TICKER
```
Open the file. Read the Cover red flags first, then dig into Analysis
quarters for trends, Valuation for DCF/reverse DCF/MoS, Credit for KMV
EDF + Altman Z.

**Earnings:** `python populate.py NVDA --force` to bypass cache. Charts
+ ratios + DCF auto-recompute against new filings + new consensus.

**Decision artifact:** `python populate.py --snapshot NVDA` writes a
timestamped copy to Archive/. Useful for post-mortems.

## Architecture

```
data/        free-source clients with parquet cache
  yfinance_client     prices, statements, options, consensus, insiders
  edgar_client        SEC company-facts XBRL (companyfacts.json)
  fred_client         Treasury curve, macro (CPI/PPI/NFP/GDP nowcast)
  finra_client        short interest (TRACE bond prices via NA-fallback)
  cache               parquet TTL store

compute/     pure-Python math (no I/O)
  ratios              all margins, leverage, liquidity, efficiency
  scoring             0-10 band lookups (CURRENT_RATIO_BANDS, etc.)
  capm                beta regression, Re, Rd, WACC
  dcf                 PV, EV, sensitivity grid, reverse DCF, MoS
  kmv                 fsolve(A, σ_A) → Distance-to-Default → EDF
  merton, altman, hillegeist
  multiples           historical-band computation
  signals             daily/weekly buy signals, market regime

xl/          Excel I/O
  styles              font, fill, border, number-format library
  build_workbooks     openpyxl template builder (idempotent)
  populator           headless cell-write engine (label→row mapping)
  macros              xlwings entry points

tests/       31 pytest cases incl. end-to-end populator with synth data

refresh.py   cache-warming CLI (`--market` or `<TICKER>`)
populate.py  workbook-filling CLI (`--market`, `<TICKER>`, `--snapshot`)
```

## Tests

```bash
python -m pytest tests/ -v
# 44 passing — ratios, scoring, DCF, KMV, Merton, Altman, Hillegeist,
# signals, multiples, owner earnings, capital allocation, EDGAR concept
# extraction + Form 4 filtering, populator helpers, end-to-end populator
# across every sheet, per-quarter ratios, multiples band, asset-light
# industry gating, positions correlation, screener flagging, summary
# aggregation, snapshot-to-archive
```

CI runs the suite on every PR (see `.github/workflows/test.yml`) and
uploads the built workbook templates as artifacts.

## Cuts vs the original workbook

- Hillegeist demoted to alternate (KMV is primary)
- Country-risk-premium hidden by default (0 for US listings)
- Personal required return removed as a separate discount rate
- Implied perpetuity growth is a derived cell, not a sensitivity grid
- DSO + Inventory Turnover hidden for asset-light industries
- Implied P/E removed from DCF outputs (FCFY kept)

## Additions vs the original

- Reverse DCF + Margin of Safety + implied-growth cell
- Multiples band (P/E, EV/EBITDA, EV/Sales, P/B, FCF Yield) with
  current vs 3y/5y/10y medians and 5y min/max
- Macro Risk Strip on Daily Plan (VIX, MOVE, HYG-LQD, DXY, 2s10s)
- RS-vs-SPY column on every sector heatmap
- 0-10 scoring with red/yellow/green color scale
- Per-quarter scoring (not just latest)
- Industry-tag gating (DSO/Inventory hidden for software/services)
- Owner Earnings (Buffett-style: NI + D&A − maintenance CAPEX)
- Capital Allocation scorecard (% of OCF to each deployment)
- Portfolio tab with pairwise correlation matrix (concentration risk
  that doesn't show in $ weights)
- Lessons Learned tab for post-mortems
- Auto-screener flagging double-buy-signal names
- Watchlist Summary tab aggregating every Ticker_*.xlsx into one
  ranked dashboard (Upside · MoS · Quality · EDF · Altman · Next Earnings)
- Earnings Calendar tab (next 60 days, sorted)
- Form 4 insider filing count (last 90 days) from SEC EDGAR submissions
- Timestamped snapshot to Archive + `--list-archive` CLI
- `--doctor` CLI to diagnose paths, API keys, and network reachability
- SEC EDGAR fallback for statements when yfinance returns empty
