"""
PORTFOLIO / CSV Ingestion Script
Spec: 00_Portfolio_Automation_Spec_V3.7.md, Section 10.

Handles all file-based inputs from Cowork input folders. Checks each of the
6 defined filenames on every run, processes any present, skips any absent
independently (no all-or-nothing dependency).

    HL inputs      (Inputs/HL/):
        account-summary.csv     -> PORTFOLIO tab (open positions)
        portfolio-summary.csv   -> PORTFOLIO_TRANSACTIONS tab (trade history)
        income-transactions.csv -> DIVIDENDS tab (dividend payments)

    Barchart inputs (Inputs/BC/):
        Custom.csv       -> Universe tab, S_BC_* (KR view)
        Performance.csv  -> Universe tab, S_BC_* (momentum)
        Fundamental.csv  -> Universe tab, S_BC_* (TTM/fundamentals)

Join keys: Code -> M_Ticker (HL) | M_BC_Ticker, falling back to M_Ticker
when M_BC_Ticker is blank (Barchart). Barchart CSV source key column is
"Symbol" (not "Ticker") -- confirmed against the actual exported files.

HL files are HL's own Windows export encoding (cp1252) -- they contain a
literal £ sign that decodes incorrectly under utf-8-sig. Barchart files are
plain utf-8-sig.

On success: archive file with `_YYYYMMDD_HHMMSS` suffix appended to the
original filename, moved into the matching Archive/<source>/ folder.

HL append dedup: before appending a row to DIVIDENDS or
PORTFOLIO_TRANSACTIONS, a composite key is checked against rows already in
the sheet (and against rows appended earlier in the same run). Matches are
skipped silently -- re-ingesting the same export doesn't double the ledger.
    DIVIDENDS key:               date + ticker + amount
    PORTFOLIO_TRANSACTIONS key:  date + ticker + transaction type + amount
Neither CSV carries a ticker column, so "ticker" in these keys is whatever
the sheet's Ticker column holds for that row (blank for a fresh append,
same as historical rows the script has always left blank) -- this still
catches the intended case, an identical file being re-run.
PORTFOLIO_TRANSACTIONS has no explicit "type" column; "transaction type" is
derived from the Reference field (HL's own B.../S... trade-reference
prefix resolves to BUY/SELL, non-trade references such as "MANAGE FEE" or
"Card Web" are used as-is).

Gap detection (Barchart): tickers flagged M_BC_Watchlist = 'Y' in the
Universe tab are "expected" in each Barchart file.
    Type A gap = expected file missing entirely for this run.
    Type B gap = file present, but an expected ticker is absent from it.
Both are logged as rows in the workbook's Scheduler tab (own header block,
written once then appended to), with a due date read from the Legend tab's
GAP_DUE_DAYS (Sprint 2 Run 5 fix #5 -- was hardcoded). No BARCHART_GAPS
worksheet, no CSV output.

Ticker enrichment (Sprint 2 Run 5 fixes #2/#3): after HL rows are appended,
DIVIDENDS and PORTFOLIO_TRANSACTIONS are swept for blank Ticker cells. Each
blank row's Description is matched against the Lookups tab's Ticker/
Description columns (L/M) -- a row's Description is considered a match for
a Lookups entry if it starts with that entry's Description text (Lookups
descriptions are the "clean" security name; DIVIDENDS/PORTFOLIO_TRANSACTIONS
descriptions carry extra suffix text -- "... Dividend Payment" or
"... <qty> @ <price>"). The longest-matching Lookups description wins, so a
more specific entry beats a shorter prefix collision. PORTFOLIO_TRANSACTIONS
only attempts this for rows the Reference field resolves to BUY/SELL --
non-trade cash movements (fees, transfers, interest) have no ticker by
definition and are left blank without being logged as a gap. Rows that
still don't match anything are logged to the Scheduler tab's action-item
block (same block as the Barchart gaps) as "<SHEET> ticker gap --
[description] -- manual Lookups update required."

Month enrichment (Sprint 2 Run 5 fix #4): DIVIDENDS' Month column is
derived from each row's Date ("YYYY-MM") and written as a literal value at
ingestion time for any row where it's still blank (pre-existing rows use an
Excel TEXT() formula for this and are untouched; only rows the Python side
would otherwise leave blank are backfilled).

Never touches: M_ columns, COVERAGE, Lookups, D_ columns.

Workbook resolution: the live .xlsm is found by globbing the base path for
a single .xlsm (abort with a clear message on 0 or 2+ matches). On save,
the patch digit (third digit) of the "v#_#_#" filename is auto-incremented,
the workbook is saved under the new name, and the old file is moved to
Archive/Workbook/ (Sprint 2 Run 5 fix #7 -- was deleted outright, which the
connected workspace folder could block). This only applies when the
workbook path is resolved automatically -- callers that pass an explicit
workbook_path (e.g. the test harness) get a plain save-in-place, so tests
can reuse one scratch file across calls.
"""

import csv
import shutil
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

from workbook_io import (
    INPUTS_HL,
    INPUTS_BC,
    ARCHIVE_HL,
    ARCHIVE_BC,
    find_workbook,
    save_workbook_with_increment,
    write_cell,
    read_legend_scalars,
    read_legend_lookup_table,
    NUMBER_FORMAT_NUMBER,
    NUMBER_FORMAT_PERCENT_SCALED,
    NUMBER_FORMAT_DATE,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HL_ENCODING = "cp1252"
BC_ENCODING = "utf-8-sig"

# Sprint 2 Run 5 fix #5: GAP_DUE_DAYS is no longer hardcoded here -- it's
# read from the Legend tab's CONFIG THRESHOLDS block at run() start and
# threaded through as a parameter. No module-level default is kept so a
# missing Legend key fails the run clearly (workbook_io.read_legend_scalars)
# instead of silently reverting to an old hardcoded number.

# Fields written to the Scheduler tab's own action-item block.
ACTION_ITEM_FIELDS = ["gap_type", "source_file", "ticker", "note", "detected_date", "due_date"]
SCHEDULER_SHEET = "Scheduler"

# Barchart "N/L" (not-licensed-for-this-symbol) and "N/A" placeholder
# tokens, plus the two free-text footer lines Barchart appends to every
# export ("Downloaded from Barchart.com as of ...", "N/L - Data could not
# be downloaded for licensing reasons."). Confirmed against the actual
# Sprint 2 Run 5 Fundamental.csv/Custom.csv/Performance.csv exports.
NULL_VALUE_TOKENS = {"N/L", "N/A"}


def _is_bc_footer_row(symbol):
    """True for Barchart's trailing free-text footer lines, which land in
    the Symbol column of a DictReader row because the export has no proper
    header for them. These never match a real ticker so they're harmless to
    process, but are filtered out up front to keep found_tickers (gap
    detection) and any match attempts clean."""
    if symbol is None:
        return False
    s = str(symbol).strip()
    return s.startswith("Downloaded from Barchart") or s.startswith("N/L -")


# HL files: matched to PORTFOLIO tab by Code -> Code, column map per §10.
# Sprint 2 Run 3 fix #3: dest_key is "Code" (not "M_Ticker") -- that's the
# actual ticker header on the PORTFOLIO tab.
# Sprint 2 Run 4 fix #3: column_map keys below are the real exported
# account-summary.csv headers (confirmed against the actual file, which
# differ from the prior draft names). "Stock", "Day gain/loss (£)" and
# "Day gain/loss (%)" are intentionally left unmapped/ignored.
HL_MATCH_FILES = {
    "account-summary.csv": {
        "sheet": "PORTFOLIO",
        "source_key": "Code",
        "dest_key": "Code",
        "column_map": {
            "Price (pence)": "S_Current_Price",
            "Value (£)": "S_MarketValue_GBP",
            "Cost (£)": "S_CostBasis",
            "Gain/loss (£)": "S_PnL_GBP",
            "Gain/loss (%)": "S_PnL_Pct",
            "Units held": "S_UnitsHeld",
        },
    },
}

# Sprint 2 Run 4 fix #3: "Price (pence)" is pence, S_Current_Price is pounds
# -- divide by 100 after the normal comma-strip/float clean.
HL_COLUMN_TRANSFORMS = {
    "S_Current_Price": lambda v: (v / 100) if isinstance(v, (int, float)) else v,
}

# HL files: appended as new rows (trade history / dividend payments) - no
# column mapping given in §10 beyond destination tab, so rows are copied
# through matching by header name.
HL_APPEND_FILES = {
    "portfolio-summary.csv": "PORTFOLIO_TRANSACTIONS",
    "income-transactions.csv": "DIVIDENDS",
}

# Barchart files: matched to Universe tab by M_BC_Ticker (fallback M_Ticker),
# column map per §10. Source key is "Symbol" in all three Barchart exports.
#
# Sprint 2 Run 3 fix #1/#2: column maps below use the real exported headers
# (confirmed against the actual Custom.csv / Performance.csv files, which
# differ from the original §10 draft names). S_BC_EBITDA is dropped entirely
# -- there's no EBITDA column in the real Custom.csv export.
#
# Sprint 2 Run 5 fix #1: Fundamental.csv column_map below replaces a set of
# placeholder header names ("PE_TTM", "EPS_TTM", ...) that never matched the
# real export -- confirmed against the actual Fundamental.csv, whose header
# row is: Symbol, Name, "Market Cap", "P/E ttm", "EPS ttm", "Net Income(a)",
# Beta, Dividend(a), "Div Yield(a)", "Earnings Date". Note this differs from
# the spec draft's assumed "Market Cap, $K" header -- the real column is
# just "Market Cap" and its values are already full-scale (not in $K).
# "Market Cap" is listed in fill_if_blank because Custom.csv's own "Market
# Cap" column feeds the same S_BC_MarketCap destination and is processed
# first -- Fundamental.csv should only fill the gap if Custom.csv didn't
# already populate it, never overwrite.
BC_MATCH_FILES = {
    "Custom.csv": {
        "sheet": "Universe",
        "source_key": "Symbol",
        "column_map": {
            "Sector": "S_BC_Sector",
            "Market Cap": "S_BC_MarketCap",
            "Beta": "S_BC_Beta60M",
            "Div Yield(a)": "S_BC_DivYieldFwd",
            "Div Payout%": "S_BC_PayoutRatio",
            "5Y Div%": "S_BC_DivGrowth5Y",
            "P/E fwd": "S_BC_PE_Fwd",
            "ROE%": "S_BC_ROE",
            "Debt/Equity": "S_BC_DebtEquity",
            "Int Cov": "S_BC_IntCoverage",
            "Analyst Rating": "S_BC_AnalystRating",
            "5Y Earn%": "S_BC_EarningsGrowth5Y",
            "Dividend Date": "S_BC_ExDivDate",
            "PEG": "S_BC_PEG",
        },
    },
    "Performance.csv": {
        "sheet": "Universe",
        "source_key": "Symbol",
        "column_map": {
            "Wtd Alpha": "S_BC_WtdAlpha",
            "5D %Chg": "S_BC_Pct5D",
            "1M %Chg": "S_BC_Pct1M",
            "3M %Chg": "S_BC_Pct3M",
            "52W %Chg": "S_BC_Pct52W",
            "YTD %Chg": "S_BC_PctYTD",
        },
    },
    "Fundamental.csv": {
        "sheet": "Universe",
        "source_key": "Symbol",
        "column_map": {
            "Market Cap": "S_BC_MarketCap",
            "P/E ttm": "S_BC_PE_TTM",
            "EPS ttm": "S_BC_EPS_TTM",
            "Net Income(a)": "S_BC_NetIncome",
            "Beta": "S_BC_Beta",
            "Dividend(a)": "S_BC_DivAnnual",
            "Div Yield(a)": "S_BC_DivYield",
            "Earnings Date": "S_BC_EarningsDate",
        },
        "fill_if_blank": {"S_BC_MarketCap"},
    },
}

# Sprint 2 Run 5 fix #5: SECTOR_ETF used to be a hardcoded module dict --
# now read from the Legend tab's "SECTOR ETF LOOKUP" block at run() start
# and passed through as a parameter (see map_sector / process_bc_match_file).


def map_sector(raw_sector, sector_etf):
    if raw_sector is None:
        return None
    return sector_etf.get(raw_sector, "UNMAP")


# Sprint 2 Run 3 fix #6: destination-column -> value type, used to (a) clean
# the raw CSV string (strip commas / "%") and (b) pick the number_format
# applied via write_cell. Not exhaustive over every S_ column in the
# workbook -- only the ones these scripts actually write.
FIELD_TYPES = {
    # BC Custom.csv
    "S_BC_Sector": "text",
    "S_Sector": "text",
    "S_BC_MarketCap": "number",
    "S_BC_Beta60M": "number",
    "S_BC_DivYieldFwd": "percent",
    "S_BC_PayoutRatio": "percent",
    "S_BC_DivGrowth5Y": "percent",
    "S_BC_PE_Fwd": "number",
    "S_BC_ROE": "percent",
    "S_BC_DebtEquity": "number",
    "S_BC_IntCoverage": "number",
    "S_BC_AnalystRating": "number",
    "S_BC_EarningsGrowth5Y": "percent",
    "S_BC_ExDivDate": "date",
    "S_BC_PEG": "number",
    # BC Performance.csv
    "S_BC_WtdAlpha": "number",
    "S_BC_Pct5D": "percent",
    "S_BC_Pct1M": "percent",
    "S_BC_Pct3M": "percent",
    "S_BC_Pct52W": "percent",
    "S_BC_PctYTD": "percent",
    # BC Fundamental.csv
    "S_BC_PE_TTM": "number",
    "S_BC_EPS_TTM": "number",
    "S_BC_NetIncome": "number",
    "S_BC_Beta": "number",
    "S_BC_DivAnnual": "number",
    "S_BC_DivYield": "percent",
    "S_BC_EarningsDate": "date",
    # HL account-summary.csv
    "S_MarketValue_GBP": "number",
    "S_CostBasis": "number",
    "S_UnitsHeld": "number",
    "S_PnL_GBP": "number",
    "S_PnL_Pct": "percent",
    "S_Current_Price": "number",
}

_FIELD_TYPE_FORMATS = {
    "number": NUMBER_FORMAT_NUMBER,
    "percent": NUMBER_FORMAT_PERCENT_SCALED,
    "date": NUMBER_FORMAT_DATE,
    "text": None,
}


def clean_field_value(raw, field_type):
    """
    Clean a raw CSV string per its field_type (see FIELD_TYPES):
      number  -> strip commas, cast to float
      percent -> strip commas and "%", cast to float (percentage-scaled,
                 e.g. "5.09%" -> 5.09, not 0.0509 -- see workbook_io's
                 NUMBER_FORMAT_PERCENT_SCALED note)
      date    -> parse "YYYY-MM-DD" into a real date object so the Date
                 number_format actually applies in Excel
      text / unmapped -> passed through unchanged
    Falls back to the raw string if parsing fails, rather than dropping
    the value or crashing the row.

    Sprint 2 Run 5 fix #1: Barchart's "N/L" (not licensed for this symbol)
    and "N/A" placeholder tokens are treated as blank (None) regardless of
    field_type, rather than being parsed as a number (fails, falls back to
    storing the literal string "N/L"/"N/A" in a numeric column) or stored
    as literal text.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    if s.upper() in NULL_VALUE_TOKENS:
        return None
    if field_type == "number":
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return raw
    if field_type == "percent":
        try:
            return float(s.replace(",", "").replace("%", ""))
        except ValueError:
            return raw
    if field_type == "date":
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError:
            return raw
    return raw


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------

def header_map(ws):
    """{header_name: column_index} from row 1."""
    return {
        cell.value: cell.column
        for cell in ws[1]
        if cell.value is not None
    }


def key_row_map(ws, key_col_idx):
    """{key_value: row_index} for all data rows (row 2+)."""
    rows = {}
    for row_idx in range(2, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=key_col_idx).value
        if val is not None:
            rows[str(val)] = row_idx
    return rows


def read_csv_rows(path, encoding):
    with open(path, newline="", encoding=encoding) as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------------------
# HL: match-and-update files (account-summary.csv)
# ---------------------------------------------------------------------------

def process_hl_match_file(wb, filename, cfg):
    rows = read_csv_rows(INPUTS_HL / filename, HL_ENCODING)
    ws = wb[cfg["sheet"]]
    headers = header_map(ws)
    dest_key_col = headers[cfg["dest_key"]]
    rows_by_key = key_row_map(ws, dest_key_col)

    for row in rows:
        ticker = row.get(cfg["source_key"])
        if not ticker or ticker not in rows_by_key:
            continue
        row_idx = rows_by_key[ticker]
        for src_col, dest_col in cfg["column_map"].items():
            if dest_col in headers and src_col in row:
                field_type = FIELD_TYPES.get(dest_col, "text")
                value = clean_field_value(row[src_col], field_type)
                transform = HL_COLUMN_TRANSFORMS.get(dest_col)
                if transform is not None:
                    value = transform(value)
                write_cell(
                    ws, row_idx, headers[dest_col], value,
                    number_format=_FIELD_TYPE_FORMATS.get(field_type),
                )


# ---------------------------------------------------------------------------
# HL: append files (portfolio-summary.csv, income-transactions.csv)
# ---------------------------------------------------------------------------

def _normalize_date(value):
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s


def _normalize_amount(value):
    if value is None or value == "":
        return None
    try:
        return round(float(str(value).replace(",", "").strip()), 2)
    except ValueError:
        return None


def _normalize_ticker(value):
    return str(value or "").strip().upper()


def transaction_type_from_reference(reference):
    """
    PORTFOLIO_TRANSACTIONS has no explicit "type" column. HL's Reference
    field encodes it: B<digits> = buy, S<digits> = sell, anything else
    (MANAGE FEE, Card Web, REG. SAVER, ...) is a non-trade cash movement
    and is used as its own type label.
    """
    ref = str(reference or "").strip()
    if len(ref) > 1 and ref[0].upper() == "B" and ref[1:].isdigit():
        return "BUY"
    if len(ref) > 1 and ref[0].upper() == "S" and ref[1:].isdigit():
        return "SELL"
    return ref.upper()


def dividend_key(date_val, ticker_val, amount_val):
    return (_normalize_date(date_val), _normalize_ticker(ticker_val), _normalize_amount(amount_val))


def transaction_key(date_val, ticker_val, reference_val, amount_val):
    return (
        _normalize_date(date_val),
        _normalize_ticker(ticker_val),
        transaction_type_from_reference(reference_val),
        _normalize_amount(amount_val),
    )


def existing_append_keys(ws, headers, sheet_name):
    """Composite keys already present in the sheet, per §-defined dedup rule."""
    keys = set()
    if sheet_name == "DIVIDENDS":
        date_c, ticker_c, amount_c = headers.get("Date"), headers.get("Ticker"), headers.get("Amount")
        if not (date_c and amount_c):
            return keys
        for row_idx in range(2, ws.max_row + 1):
            d = ws.cell(row=row_idx, column=date_c).value
            t = ws.cell(row=row_idx, column=ticker_c).value if ticker_c else None
            a = ws.cell(row=row_idx, column=amount_c).value
            if d is None and a is None:
                continue
            keys.add(dividend_key(d, t, a))
    elif sheet_name == "PORTFOLIO_TRANSACTIONS":
        date_c = headers.get("Trade date")
        ticker_c = headers.get("Ticker")
        ref_c = headers.get("Reference")
        amount_c = headers.get("Value (£)")
        if not (date_c and amount_c):
            return keys
        for row_idx in range(2, ws.max_row + 1):
            d = ws.cell(row=row_idx, column=date_c).value
            t = ws.cell(row=row_idx, column=ticker_c).value if ticker_c else None
            r = ws.cell(row=row_idx, column=ref_c).value if ref_c else None
            a = ws.cell(row=row_idx, column=amount_c).value
            if d is None and a is None:
                continue
            keys.add(transaction_key(d, t, r, a))
    return keys


def key_for_incoming_row(sheet_name, row):
    """row: a CSV DictReader row (raw strings)."""
    if sheet_name == "DIVIDENDS":
        return dividend_key(row.get("Date"), row.get("Ticker"), row.get("Amount"))
    if sheet_name == "PORTFOLIO_TRANSACTIONS":
        return transaction_key(
            row.get("Trade date"), row.get("Ticker"), row.get("Reference"), row.get("Value (£)")
        )
    return None


def process_hl_append_file(wb, filename, sheet_name):
    rows = read_csv_rows(INPUTS_HL / filename, HL_ENCODING)
    ws = wb[sheet_name]
    headers = header_map(ws)

    existing_keys = existing_append_keys(ws, headers, sheet_name)

    next_row = ws.max_row + 1
    appended, skipped = 0, 0
    for row in rows:
        key = key_for_incoming_row(sheet_name, row)
        if key is not None and key in existing_keys:
            skipped += 1
            continue

        for col_name, value in row.items():
            if col_name in headers:
                write_cell(ws, next_row, headers[col_name], value)

        if key is not None:
            existing_keys.add(key)
        next_row += 1
        appended += 1

    return appended, skipped


# ---------------------------------------------------------------------------
# Ticker enrichment (Sprint 2 Run 5 fixes #2/#3) + DIVIDENDS Month (fix #4)
# ---------------------------------------------------------------------------

LOOKUPS_SHEET = "Lookups"
LOOKUPS_TICKER_COL = 12   # column L
LOOKUPS_DESCRIPTION_COL = 13  # column M


def build_lookup_pairs(wb):
    """
    [(ticker, description), ...] from the Lookups tab's L/M columns, sorted
    by description length descending so the longest (most specific) match
    wins when checking "does this row's Description start with this
    Lookups Description" against multiple candidates.
    """
    ws = wb[LOOKUPS_SHEET]
    pairs = []
    for row_idx in range(2, ws.max_row + 1):
        ticker = ws.cell(row=row_idx, column=LOOKUPS_TICKER_COL).value
        description = ws.cell(row=row_idx, column=LOOKUPS_DESCRIPTION_COL).value
        if ticker and description:
            pairs.append((str(ticker), str(description)))
    pairs.sort(key=lambda p: -len(p[1]))
    return pairs


def match_ticker_by_description(description, lookup_pairs):
    """
    Lookups' Description column holds the "clean" security name; DIVIDENDS/
    PORTFOLIO_TRANSACTIONS descriptions carry extra suffix text ("... 20 @
    5377.8393", "... Overseas Dividend Payment"). A row matches a Lookups
    entry if its Description starts with that entry's Description. Checking
    longest-first means a more specific Lookups entry wins over a shorter
    one that happens to also be a prefix.
    """
    if not description:
        return None
    for ticker, lookup_desc in lookup_pairs:
        if description.startswith(lookup_desc):
            return ticker
    return None


def _month_str(date_val):
    """DIVIDENDS' Month column, "YYYY-MM" -- mirrors the TEXT(date,"YYYY-MM")
    formula pre-existing rows use, but written as a literal value so it
    doesn't depend on Excel recalculating formulas for Python-appended rows."""
    if date_val is None:
        return None
    if isinstance(date_val, datetime):
        return date_val.strftime("%Y-%m")
    s = str(date_val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    return None


def enrich_dividends(wb, lookup_pairs, run_date):
    """
    Sweeps every DIVIDENDS row: blank Ticker cells are matched against
    Lookups via Description (gap-logged if no match); blank Month cells are
    derived from Date. Returns (matched, gapped, month_filled) counts plus
    the list of gap dicts (caller logs them to Scheduler alongside Barchart
    gaps).
    """
    ws = wb["DIVIDENDS"]
    headers = header_map(ws)
    date_c, desc_c, ticker_c, month_c = (
        headers.get("Date"), headers.get("Description"),
        headers.get("Ticker"), headers.get("Month"),
    )
    matched, gapped, month_filled = 0, 0, 0
    gaps = []
    for row_idx in range(2, ws.max_row + 1):
        description = ws.cell(row=row_idx, column=desc_c).value if desc_c else None

        if ticker_c:
            existing_ticker = ws.cell(row=row_idx, column=ticker_c).value
            if existing_ticker is None or str(existing_ticker).strip() == "":
                found = match_ticker_by_description(description, lookup_pairs)
                if found:
                    write_cell(ws, row_idx, ticker_c, found)
                    matched += 1
                elif description:
                    gapped += 1
                    gaps.append(description)

        if month_c and date_c:
            existing_month = ws.cell(row=row_idx, column=month_c).value
            if existing_month is None or str(existing_month).strip() == "":
                date_val = ws.cell(row=row_idx, column=date_c).value
                month_val = _month_str(date_val)
                if month_val:
                    write_cell(ws, row_idx, month_c, month_val)
                    month_filled += 1

    gap_records = [
        {
            "gap_type": "DIVIDENDS Ticker Gap",
            "source_file": "DIVIDENDS",
            "ticker": "",
            "note": f"DIVIDENDS ticker gap — {desc} — manual Lookups update required.",
            "detected_date": run_date.date(),
        }
        for desc in gaps
    ]
    return matched, gapped, month_filled, gap_records


def enrich_portfolio_transactions(wb, lookup_pairs, run_date):
    """
    Sweeps every PORTFOLIO_TRANSACTIONS row with a blank Ticker. Only rows
    the Reference field resolves to BUY/SELL are candidates for matching --
    non-trade cash movements (fees, transfers, interest, regular-savings
    references) legitimately have no ticker and are left blank without
    being logged as a gap. Returns (matched, gapped) counts plus gap dicts.
    """
    ws = wb["PORTFOLIO_TRANSACTIONS"]
    headers = header_map(ws)
    desc_c, ticker_c, ref_c = (
        headers.get("Description"), headers.get("Ticker"), headers.get("Reference"),
    )
    matched, gapped = 0, 0
    gaps = []
    for row_idx in range(2, ws.max_row + 1):
        if ticker_c is None:
            break
        existing_ticker = ws.cell(row=row_idx, column=ticker_c).value
        if existing_ticker is not None and str(existing_ticker).strip() != "":
            continue

        reference = ws.cell(row=row_idx, column=ref_c).value if ref_c else None
        if transaction_type_from_reference(reference) not in ("BUY", "SELL"):
            continue  # cash movement -- no ticker to resolve, not a gap

        description = ws.cell(row=row_idx, column=desc_c).value if desc_c else None
        found = match_ticker_by_description(description, lookup_pairs)
        if found:
            write_cell(ws, row_idx, ticker_c, found)
            matched += 1
        elif description:
            gapped += 1
            gaps.append(description)

    gap_records = [
        {
            "gap_type": "PORTFOLIO_TRANSACTIONS Ticker Gap",
            "source_file": "PORTFOLIO_TRANSACTIONS",
            "ticker": "",
            "note": f"PORTFOLIO_TRANSACTIONS ticker gap — {desc} — manual Lookups update required.",
            "detected_date": run_date.date(),
        }
        for desc in gaps
    ]
    return matched, gapped, gap_records


# ---------------------------------------------------------------------------
# Barchart: match-and-update files
# ---------------------------------------------------------------------------

def effective_bc_ticker(headers, ws, row_idx):
    """M_BC_Ticker if populated, else fall back to M_Ticker (per data dict)."""
    bc_col = headers.get("M_BC_Ticker")
    val = ws.cell(row=row_idx, column=bc_col).value if bc_col else None
    if val:
        return str(val)
    m_col = headers.get("M_Ticker")
    val = ws.cell(row=row_idx, column=m_col).value if m_col else None
    return str(val) if val else None


def universe_ticker_row_map(ws):
    """{effective_ticker: row_index} across the whole Universe tab."""
    headers = header_map(ws)
    mapping = {}
    for row_idx in range(2, ws.max_row + 1):
        ticker = effective_bc_ticker(headers, ws, row_idx)
        if ticker:
            mapping[ticker] = row_idx
    return mapping


def process_bc_match_file(wb, filename, cfg, ticker_row_map, sector_etf):
    rows = read_csv_rows(INPUTS_BC / filename, BC_ENCODING)
    ws = wb[cfg["sheet"]]
    headers = header_map(ws)
    fill_if_blank = cfg.get("fill_if_blank", set())

    for row in rows:
        ticker = row.get(cfg["source_key"])
        if not ticker or _is_bc_footer_row(ticker) or ticker not in ticker_row_map:
            continue
        row_idx = ticker_row_map[ticker]
        for src_col, dest_col in cfg["column_map"].items():
            if dest_col in headers and src_col in row:
                if dest_col in fill_if_blank:
                    current = ws.cell(row=row_idx, column=headers[dest_col]).value
                    if current is not None and str(current).strip() != "":
                        continue  # Sprint 2 Run 5 fix #1: don't overwrite Custom.csv's value
                field_type = FIELD_TYPES.get(dest_col, "text")
                value = clean_field_value(row[src_col], field_type)
                write_cell(
                    ws, row_idx, headers[dest_col], value,
                    number_format=_FIELD_TYPE_FORMATS.get(field_type),
                )
                # Sprint 2 Run 4 fix #2: alongside the untouched S_BC_Sector
                # text, also derive the S_Sector ETF abbreviation.
                if dest_col == "S_BC_Sector" and "S_Sector" in headers:
                    write_cell(
                        ws, row_idx, headers["S_Sector"], map_sector(value, sector_etf),
                    )


# ---------------------------------------------------------------------------
# Archiving
# ---------------------------------------------------------------------------

def archive_file(src_path, archive_dir, timestamp):
    archive_dir.mkdir(parents=True, exist_ok=True)
    stamped_name = f"{src_path.stem}_{timestamp}{src_path.suffix}"
    shutil.move(str(src_path), str(archive_dir / stamped_name))


# ---------------------------------------------------------------------------
# Gap detection & Scheduler action items (Barchart only)
# ---------------------------------------------------------------------------

def expected_watchlist_tickers(ws):
    headers = header_map(ws)
    watchlist_col = headers.get("M_BC_Watchlist")
    if not watchlist_col:
        return set()
    expected = set()
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row=row_idx, column=watchlist_col).value == "Y":
            ticker = effective_bc_ticker(headers, ws, row_idx)
            if ticker:
                expected.add(ticker)
    return expected


def detect_gaps(present_files, expected_tickers, run_date, gap_due_days):
    """present_files: {filename: set_of_tickers_found_in_file or None if missing}"""
    gaps = []
    # Real date objects (not strftime strings) so the Date number_format
    # applied in log_gaps_to_scheduler actually renders as a date in Excel.
    detected_date = run_date.date() if hasattr(run_date, "date") else run_date
    due_date = detected_date + timedelta(days=gap_due_days)
    for filename, found_tickers in present_files.items():
        if found_tickers is None:
            gaps.append({
                "gap_type": "Type A",
                "source_file": filename,
                "ticker": "",
                "note": "expected file missing for this run",
                "detected_date": detected_date,
                "due_date": due_date,
            })
        else:
            for ticker in sorted(expected_tickers - found_tickers):
                gaps.append({
                    "gap_type": "Type B",
                    "source_file": filename,
                    "ticker": ticker,
                    "note": "expected ticker missing from file",
                    "detected_date": detected_date,
                    "due_date": due_date,
                })
    return gaps


def find_action_item_header_row(ws):
    """Locate our own action-item header block if one was already written."""
    for row_idx in range(1, ws.max_row + 1):
        if ws.cell(row=row_idx, column=1).value == ACTION_ITEM_FIELDS[0]:
            return row_idx
    return None


def log_gaps_to_scheduler(wb, gaps, gap_due_days):
    """
    Append gap rows to the workbook's Scheduler tab. The Scheduler tab
    already holds an unrelated manual review table -- our action items get
    their own header block (written once, below whatever's already there)
    rather than being interleaved into that table's columns.

    gaps may omit "due_date" (the ticker-enrichment gaps from fixes #2/#3
    don't set one when built) -- it's backfilled here from detected_date +
    gap_due_days so every gap type shares the same due-date convention read
    from the Legend tab.
    """
    if not gaps:
        return
    ws = wb[SCHEDULER_SHEET]
    header_row = find_action_item_header_row(ws)
    if header_row is None:
        header_row = ws.max_row + 2  # blank separator line from existing content
        for col_idx, name in enumerate(ACTION_ITEM_FIELDS, start=1):
            write_cell(ws, header_row, col_idx, name)

    next_row = ws.max_row + 1
    for gap in gaps:
        detected_date = gap["detected_date"]
        due_date = gap.get("due_date") or (detected_date + timedelta(days=gap_due_days))
        row_values = dict(gap, due_date=due_date)
        for col_idx, field in enumerate(ACTION_ITEM_FIELDS, start=1):
            value = row_values[field]
            fmt = NUMBER_FORMAT_DATE if field in ("detected_date", "due_date") else None
            write_cell(ws, next_row, col_idx, value, number_format=fmt)
        next_row += 1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(workbook_path=None):
    run_date = datetime.now()
    timestamp = run_date.strftime("%Y%m%d_%H%M%S")

    resolved_by_glob = workbook_path is None
    workbook_path = find_workbook() if resolved_by_glob else Path(workbook_path)

    wb = openpyxl.load_workbook(
        workbook_path, keep_vba=workbook_path.suffix == ".xlsm"
    )

    # --- Sprint 2 Run 5 fix #5: Legend tab as source of truth ---
    legend_scalars = read_legend_scalars(wb, ["GAP_DUE_DAYS"])
    gap_due_days = int(legend_scalars["GAP_DUE_DAYS"])
    sector_etf = read_legend_lookup_table(wb, "SECTOR ETF LOOKUP")

    # --- HL inputs: process any present, skip any absent, independently ---
    for filename, cfg in HL_MATCH_FILES.items():
        path = INPUTS_HL / filename
        if path.exists():
            try:
                process_hl_match_file(wb, filename, cfg)
                archive_file(path, ARCHIVE_HL, timestamp)
            except Exception as exc:
                print(f"[SKIP] {filename}: {exc}")

    for filename, sheet_name in HL_APPEND_FILES.items():
        path = INPUTS_HL / filename
        if path.exists():
            try:
                appended, skipped = process_hl_append_file(wb, filename, sheet_name)
                archive_file(path, ARCHIVE_HL, timestamp)
                print(f"  {filename}: {appended} row(s) appended, {skipped} duplicate(s) skipped.")
            except Exception as exc:
                print(f"[SKIP] {filename}: {exc}")

    # --- Sprint 2 Run 5 fixes #2/#3/#4: Ticker enrichment + Month ---
    lookup_pairs = build_lookup_pairs(wb)
    div_matched, div_gapped, month_filled, div_gap_records = enrich_dividends(
        wb, lookup_pairs, run_date
    )
    pt_matched, pt_gapped, pt_gap_records = enrich_portfolio_transactions(
        wb, lookup_pairs, run_date
    )
    print(
        f"  DIVIDENDS ticker enrichment: {div_matched} matched, {div_gapped} gap(s) "
        f"logged, {month_filled} Month value(s) backfilled."
    )
    print(
        f"  PORTFOLIO_TRANSACTIONS ticker enrichment: {pt_matched} matched, "
        f"{pt_gapped} gap(s) logged."
    )

    # --- Barchart inputs: process any present, track presence for gaps ---
    universe_ws = wb["Universe"]
    ticker_row_map = universe_ticker_row_map(universe_ws)
    expected_tickers = expected_watchlist_tickers(universe_ws)

    present_files = {}
    for filename, cfg in BC_MATCH_FILES.items():
        path = INPUTS_BC / filename
        if path.exists():
            try:
                rows = read_csv_rows(path, BC_ENCODING)
                found_tickers = {
                    r[cfg["source_key"]] for r in rows
                    if r.get(cfg["source_key"]) and not _is_bc_footer_row(r[cfg["source_key"]])
                }
                process_bc_match_file(wb, filename, cfg, ticker_row_map, sector_etf)
                archive_file(path, ARCHIVE_BC, timestamp)
                present_files[filename] = found_tickers
            except Exception as exc:
                print(f"[SKIP] {filename}: {exc}")
                present_files[filename] = None
        else:
            present_files[filename] = None

    gaps = detect_gaps(present_files, expected_tickers, run_date, gap_due_days)
    all_gaps = gaps + div_gap_records + pt_gap_records
    log_gaps_to_scheduler(wb, all_gaps, gap_due_days)

    if resolved_by_glob:
        workbook_path = save_workbook_with_increment(wb, workbook_path)
    else:
        wb.save(workbook_path)

    print(
        f"Run complete. {len(gaps)} Barchart gap(s), {div_gapped + pt_gapped} "
        f"ticker gap(s) logged to Scheduler tab. Workbook: {workbook_path.name}"
    )
    return wb


if __name__ == "__main__":
    run()
