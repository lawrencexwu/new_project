from __future__ import annotations

import time
from pathlib import Path

import pandas as pd

import config


def _cache_root() -> Path:
    root = config.get("cache_dir")
    root.mkdir(parents=True, exist_ok=True)
    return root


def _path(namespace: str, key: str) -> Path:
    ns = _cache_root() / namespace
    ns.mkdir(parents=True, exist_ok=True)
    return ns / f"{key}.parquet"


def _ttl_seconds(namespace: str) -> float:
    ttls = config.get("cache_ttl_days", {})
    days = ttls.get(namespace, 1)
    return float(days) * 86400.0


def is_fresh(namespace: str, key: str) -> bool:
    p = _path(namespace, key)
    if not p.exists():
        return False
    age = time.time() - p.stat().st_mtime
    return age < _ttl_seconds(namespace)


def read(namespace: str, key: str) -> pd.DataFrame | None:
    p = _path(namespace, key)
    if not p.exists():
        return None
    try:
        return pd.read_parquet(p)
    except Exception:
        return None


def write(namespace: str, key: str, df: pd.DataFrame) -> None:
    if df is None or len(df) == 0:
        return
    p = _path(namespace, key)
    df.to_parquet(p)


def get_or_fetch(namespace: str, key: str, fetcher, force: bool = False):
    if not force and is_fresh(namespace, key):
        cached = read(namespace, key)
        if cached is not None:
            return cached
    df = fetcher()
    write(namespace, key, df)
    return df
