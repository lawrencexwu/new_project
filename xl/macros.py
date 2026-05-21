"""xlwings macro entry points. Called from buttons in the workbooks.

These delegate to xl.populator which uses openpyxl underneath, then reload
the workbook in Excel so the user sees the fresh values.
"""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

from xl import populator


def _book() -> xw.Book:
    return xw.Book.caller()


def _save_close_repopulate_reopen(book: xw.Book, fn, *args, **kwargs):
    """Save current state, close, run headless populator, reopen."""
    path = Path(book.fullname)
    app = book.app
    book.save()
    book.close()
    fn(path, *args, **kwargs)
    app.books.open(str(path))


def refresh_market(force: bool = False):
    """Populate Market_Daily.xlsx: heatmaps, signals, macro."""
    book = _book()
    _save_close_repopulate_reopen(book, populator.populate_market_daily, force=force)


def refresh_ticker(force: bool = False):
    """Populate this ticker workbook. Looks up the ticker from the Cover sheet."""
    book = _book()
    ws = book.sheets["Cover"]
    ticker = ws.range("B3").value
    if not ticker:
        raise ValueError("Enter a ticker in cell B3 of the Cover sheet first.")

    path = Path(book.fullname)
    app = book.app
    book.save()
    book.close()
    populator.populate_ticker(path, path, ticker, force=force)
    app.books.open(str(path))
