import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent
_SETTINGS = None


def _expand(path: str) -> Path:
    return Path(os.path.expanduser(path)).resolve()


def load() -> dict:
    global _SETTINGS
    if _SETTINGS is not None:
        return _SETTINGS

    local = ROOT / "settings.local.json"
    default = ROOT / "settings.example.json"
    target = local if local.exists() else default

    with open(target) as f:
        s = json.load(f)

    for key in ("drive_root", "cache_dir", "tickers_dir", "archive_dir",
                "market_daily_path", "template_path"):
        if key in s:
            s[key] = _expand(s[key])

    _SETTINGS = s
    return s


def get(key, default=None):
    return load().get(key, default)


def get_watchlist() -> list[str]:
    """Return the user's watchlist tickers.

    Prefers a plain-text `watchlist.txt` in the repo root (one ticker per
    line, '#' starts a comment) — easy to edit without JSON syntax. Falls
    back to the 'watchlist' key in settings if the file is absent/empty.
    """
    wl_file = ROOT / "watchlist.txt"
    if wl_file.exists():
        tickers: list[str] = []
        for line in wl_file.read_text().splitlines():
            line = line.split("#", 1)[0].strip()
            if line:
                tickers.append(line.upper())
        if tickers:
            return tickers
    return load().get("watchlist", []) or []

