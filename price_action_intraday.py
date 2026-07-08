"""
Price Action Script — Intraday cadence
Spec: 00_Portfolio_Automation_Spec_V3.8.md, Section 7a. Sprint 2 script
re-engineering: replaces the 1H/4H scope of price_action_4h.py.

Runs: 1H/4H cadence (11am / 3pm / 7pm, or hourly -- schedule is external to
this script; every run behaves identically).

Scope each run:
    M_Eliminated != "No Touch"   (universal kill switch, checked first)
    -- all rows, all sleeves. No M_Fetch_Cadence filter (§7a: cadence is
    now governed by the run schedule, not a per-row column check).

S_ columns fed (9 attributes only):
    S_Daily_High, S_Daily_Low, S_Daily_Open, S_Last_Price, S_Current_Price,
    S_LastTradedTime, S_1D_Change_%, S_2D_Change_%, S_3D_Change_%

No D_ columns. No CSV ingestion. No cadence filter.

Source: one yfinance Ticker.history() call per row, no `.info[]` calls.

Notes on scope decisions not fully specified in §7a itself:
  - §7a's literal text says "history(1d) only" as the lightest possible YF
    call. A single history(period="1d") pull returns exactly one daily bar,
    which cannot support S_1D/2D/3D_Change_% (each needs a prior close N
    trading days back). Raised with the user directly (see conversation) --
    resolved as: pull history(period="5d") instead (still one Ticker call,
    still zero .info[] calls, so the "lightest possible / no .info" intent
    is preserved) so the three change-% columns actually compute. This is a
    deliberate deviation from the literal "1d" period string, not an
    oversight.
  - S_Current_Price / S_Last_Price: the Data Dictionary (§20) sources both
    from `info['currentPrice']`, but §7a bans `.info[]` calls entirely for
    this script. Both are set from the latest daily bar's Close instead
    (same value price_action_4h.py already used for S_Last_Price) -- the
    two columns are aliases of each other per the Data Dictionary note, so
    this preserves that relationship even without a live quote. Flagged,
    not solved: this makes both fields "last close", not a true
    intraday tick price to yfinance's fastinfo, an option if that's
    revisited later.
  - S_LastTradedTime: Data Dictionary just says "YF", normally
    `info['regularMarketTime']` (see fetch_engine_monthly.py), unavailable
    here without `.info[]`. Uses the timestamp of the latest returned bar
    (hist.index[-1]) instead, formatted "YYYY-MM-DD HH:MM:SS".
  - Ticker symbol resolution: same S_YF_Ticker -> M_Ticker fallback as
    price_action_4h.py / fetch_engine_monthly.py.

openpyxl constraint (§1): surgical cell edits only, never a pandas
rewrite of the sheet -- column dtypes/formatting must survive.
"""

import numpy as np
import yfinance as yf
from pathlib import Path

import openpyxl

from workbook_io import (
    find_workbook,
    save_workbook_with_increment,
    write_cell,
    NUMBER_FORMAT_NUMBER,
    NUMBER_FORMAT_PERCENT_SCALED,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SHEET_NAME = "Universe"

NO_TOUCH = "No Touch"

S_COLUMNS_INTRADAY = [
    "S_Daily_High", "S_Daily_Low", "S_Daily_Open",
    "S_Last_Price", "S_Current_Price", "S_LastTradedTime",
    "S_1D_Change_%", "S_2D_Change_%", "S_3D_Change_%",
]

FIELD_TYPES = {
    "S_Daily_High": "number",
    "S_Daily_Low": "number",
    "S_Daily_Open": "number",
    "S_Last_Price": "number",
    "S_Current_Price": "number",
    "S_LastTradedTime": "text",
    "S_1D_Change_%": "percent_scaled",
    "S_2D_Change_%": "percent_scaled",
    "S_3D_Change_%": "percent_scaled",
}

_FIELD_TYPE_FORMATS = {
    "number": NUMBER_FORMAT_NUMBER,
    "percent_scaled": NUMBER_FORMAT_PERCENT_SCALED,
    "text": None,
}


# ---------------------------------------------------------------------------
# Sheet helpers (same pattern as price_action_4h.py)
# ---------------------------------------------------------------------------

def header_map(ws):
    """{header_name: column_index} from row 1."""
    return {cell.value: cell.column for cell in ws[1] if cell.value is not None}


def get_cell(ws, headers, row_idx, col_name, default=None):
    col = headers.get(col_name)
    if col is None:
        return default
    val = ws.cell(row=row_idx, column=col).value
    return val if val is not None else default


def set_cell(ws, headers, row_idx, col_name, value):
    col = headers.get(col_name)
    if col is None:
        return  # column not present in sheet -- skip silently, don't crash a run
    field_type = FIELD_TYPES.get(col_name, "text")
    write_cell(ws, row_idx, col, value, number_format=_FIELD_TYPE_FORMATS.get(field_type))


# ---------------------------------------------------------------------------
# Scope filter -- §7a
# ---------------------------------------------------------------------------

def in_scope(ws, headers, row_idx):
    """M_Eliminated != "No Touch" -- all rows, all sleeves, no cadence filter."""
    return get_cell(ws, headers, row_idx, "M_Eliminated") != NO_TOUCH


def resolve_ticker(ws, headers, row_idx):
    yf_ticker = get_cell(ws, headers, row_idx, "S_YF_Ticker")
    if yf_ticker:
        return str(yf_ticker)
    m_ticker = get_cell(ws, headers, row_idx, "M_Ticker")
    return str(m_ticker) if m_ticker else None


# ---------------------------------------------------------------------------
# yfinance fetch -- S_ columns
# ---------------------------------------------------------------------------

def fetch_price_action(symbol):
    """
    One yfinance Ticker object per row, one history() call, no .info[].
    period="5d" (see module docstring re: deviation from literal "1d") so
    S_1D/2D/3D_Change_% can be computed. Returns a dict of the 9 S_
    columns fed by §7a, or None if the pull failed / no data.
    """
    try:
        t = yf.Ticker(symbol)
        hist = t.history(period="5d")
    except Exception as exc:
        print(f"[SKIP] {symbol}: yfinance error: {exc}")
        return None

    if hist.empty:
        print(f"[SKIP] {symbol}: empty history")
        return None

    last = hist.iloc[-1]
    closes = hist["Close"]
    last_price = float(last["Close"])

    def pct_change(n):
        if len(closes) <= n:
            return None
        prior = closes.iloc[-1 - n]
        if prior in (0, None) or (isinstance(prior, float) and np.isnan(prior)):
            return None
        return (last_price - prior) / prior * 100

    return {
        "S_Daily_High": float(last["High"]),
        "S_Daily_Low": float(last["Low"]),
        "S_Daily_Open": float(last["Open"]),
        "S_Last_Price": last_price,
        "S_Current_Price": last_price,
        "S_LastTradedTime": hist.index[-1].strftime("%Y-%m-%d %H:%M:%S"),
        "S_1D_Change_%": pct_change(1),
        "S_2D_Change_%": pct_change(2),
        "S_3D_Change_%": pct_change(3),
    }


# ---------------------------------------------------------------------------
# Main run
# ---------------------------------------------------------------------------

def run(workbook_path=None):
    """
    workbook_path: if omitted, the live .xlsm is resolved by globbing the
    base path (aborts on 0 matches) and the save at the end rotates the
    patch-digit filename, moving the old one to Archive/Workbook/. If an
    explicit path is passed (e.g. by the test harness), it's used as-is
    and saved in place.
    """
    resolved_by_glob = workbook_path is None
    workbook_path = find_workbook() if resolved_by_glob else Path(workbook_path)

    wb = openpyxl.load_workbook(workbook_path, keep_vba=str(workbook_path).endswith(".xlsm"))
    ws = wb[SHEET_NAME]
    headers = header_map(ws)

    processed, skipped = 0, 0
    for row_idx in range(2, ws.max_row + 1):
        if not in_scope(ws, headers, row_idx):
            continue

        symbol = resolve_ticker(ws, headers, row_idx)
        if not symbol:
            skipped += 1
            continue

        s_values = fetch_price_action(symbol)
        if s_values is None:
            skipped += 1
            continue

        for col_name in S_COLUMNS_INTRADAY:
            set_cell(ws, headers, row_idx, col_name, s_values[col_name])

        processed += 1

    if resolved_by_glob:
        workbook_path = save_workbook_with_increment(wb, workbook_path)
    else:
        wb.save(workbook_path)

    print(f"Run complete (intraday). {processed} row(s) updated, {skipped} skipped. Workbook: {workbook_path.name}")
    return wb


if __name__ == "__main__":
    run()
