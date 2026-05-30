@echo off
REM Morning routine for Windows — double-click from File Explorer.
REM
REM First-time setup: just double-click. The script will create a venv if
REM one doesn't exist. Subsequent runs reuse the existing venv.

cd /d "%~dp0"

if not exist ".venv\" (
    echo ==^> First run, creating Python venv...
    python -m venv .venv
    .venv\Scripts\pip install -q -r requirements.txt
)
call .venv\Scripts\activate.bat

if exist ".git\" (
    git pull --quiet 2>nul
)

python populate.py --routine

REM Open Market_Daily.xlsx using the configured path
for /f "delims=" %%P in ('python -c "import config; print(config.get('market_daily_path'))"') do set MARKET=%%P
if exist "%MARKET%" (
    echo ==^> Opening Market_Daily.xlsx...
    start "" "%MARKET%"
)

echo.
pause
