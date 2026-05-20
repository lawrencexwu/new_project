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
