"""
Fetch Engine — Weekly cadence (fundamentals)
Spec: 00_Portfolio_Automation_Spec_V3.8.md, Section 7c. Sprint 2 script
re-engineering: replaces fetch_engine_monthly.py entirely.

Runs: Sunday (weekly).

Scope (checked in this order):
    M_Eliminated != "No Touch"          (universal kill switch, checked first)
    M_Div_Coupon_Class in ELIGIBLE_DIV_COUPON_CLASSES
        {Aristocrat_King, Aristocrat, Achiever, Contender, HighIncome, MedIncome}
        (same eligibility list as fetch_engine_monthly.py -- copied unchanged)

No M_Fetch_Cadence column filter -- §7c: cadence is now governed by the
Sunday run schedule itself, not a per-row "Weekly" cadence value (same
posture as price_action_intraday.py / price_action_eod.py dropping their
own cadence checks). This is the only scope change from
fetch_engine_monthly.py -- everything computed, and how, is unchanged.

S_ columns fed -- direct 1:1 off a single yfinance `.info` pull:
    S_Name, S_Sector, S_Country, S_Exchange, S_Industry, S_Sub_Industry,
    S_MCap, S_Beta, S_PE_Ratio, S_PayoutRatio, S_DebtEquity, S_ROE,
    S_Dividend_Yield, S_Average_Volume, S_52W_High, S_52W_Low,
    S_ExDividend_Date, S_Reporting_Date, S_LastTradedTime

S_ columns fed -- computed 5Y CAGR / averages (best-effort, see notes in
the original fetch_engine_monthly.py -- unchanged here):
    S_EPS_Growth_5Y, S_DivGrowth_5Y, S_DivYield_5YAvg, S_PE_5YAvg,
    S_Yrs_DivIncome_Buys_1Share

Out of scope for this script: S_DivStreak_Years is fed annually from the
CCC list per §7c's "Source" line, not from yfinance -- never written here.

Gap logging: one row per ticker/column that failed to resolve, written to
Outputs/BC/fetch_gaps_YYYYMMDD.csv (unchanged from fetch_engine_monthly.py).

Note (flagged, not implemented -- out of the read scope for this build):
§7c of the master spec also calls for updating a
"Process_Owner_mapping_pyScript_Name" field on the Data_Dictionary tab from
"fetch_engine_monthly.py" to "fetch_engine_weekly.py". The task instructions
for this Sprint 2 pass scoped reading to §7a/§7b/§10 only and did not ask
for that Data_Dictionary update, so it isn't implemented here -- flagging it
as a follow-up in case it was an intentional omission vs. an oversight.

Sprint 2 Run 5 fix #5 (inherited unchanged): the yfinance info['country'] ->
S_Country abbreviation table is read from the Legend tab's "COUNTRY LOOKUP"
block at run() start and threaded through map_country() as a parameter.

openpyxl constraint (§1, inherited convention): surgical cell edits only,
never a pandas rewrite of the sheet.
"""

import csv
from datetime import datetime
from pathlib import Path

import yfinance as yf
import openpyxl

from workbook_io import (
    find_workbook,
    save_workbook_with_increment,
    OUTPUTS_BC,
    write_cell,
    read_legend_lookup_table,
    NUMBER_FORMAT_NUMBER,
    NUMBER_FORMAT_PERCENT_SCALED,
    NUMBER_FORMAT_PERCENT_FRACTION,
    NUMBER_FORMAT_DATE,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SHEET_NAME = "Universe"

NO_TOUCH = "No Touch"

ELIGIBLE_DIV_COUPON_CLASSES = {
    "Aristocrat_King", "Aristocrat", "Achiever", "Contender", "HighIncome", "MedIncome",
}

GAP_OUTPUT_DIR = OUTPUTS_BC


def map_country(raw_country, country_lookup):
    if raw_country is None:
        return None
    return country_lookup.get(raw_country, raw_country)

S_COLUMNS_INFO_DIRECT = [
    "S_Name", "S_Sector", "S_Country", "S_Exchange", "S_Industry", "S_Sub_Industry",
    "S_MCap", "S_Beta", "S_PE_Ratio", "S_PayoutRatio", "S_DebtEquity", "S_ROE",
    "S_Dividend_Yield", "S_Average_Volume", "S_52W_High", "S_52W_Low",
    "S_ExDividend_Date", "S_Reporting_Date", "S_LastTradedTime",
]

S_COLUMNS_COMPUTED = [
    "S_EPS_Growth_5Y", "S_DivGrowth_5Y", "S_DivYield_5YAvg", "S_PE_5YAvg",
    "S_Yrs_DivIncome_Buys_1Share",
]

S_COLUMNS_ALL = S_COLUMNS_INFO_DIRECT + S_COLUMNS_COMPUTED

FIELD_TYPES = {
    "S_Name": "text",
    "S_Sector": "text",
    "S_Country": "text",
    "S_Exchange": "text",
    "S_Industry": "text",
    "S_Sub_Industry": "text",
    "S_MCap": "number",
    "S_Beta": "number",
    "S_PE_Ratio": "number",
    "S_PayoutRatio": "percent_fraction",
    "S_DebtEquity": "number",
    "S_ROE": "percent_fraction",
    "S_Dividend_Yield": "percent_fraction",
    "S_Average_Volume": "number",
    "S_52W_High": "number",
    "S_52W_Low": "number",
    "S_ExDividend_Date": "date",
    "S_Reporting_Date": "date",
    "S_LastTradedTime": "text",
    "S_EPS_Growth_5Y": "percent_scaled",
    "S_DivGrowth_5Y": "percent_scaled",
    "S_DivYield_5YAvg": "percent_scaled",
    "S_PE_5YAvg": "number",
    "S_Yrs_DivIncome_Buys_1Share": "number",
}

_FIELD_TYPE_FORMATS = {
    "number": NUMBER_FORMAT_NUMBER,
    "percent_scaled": NUMBER_FORMAT_PERCENT_SCALED,
    "percent_fraction": NUMBER_FORMAT_PERCENT_FRACTION,
    "date": NUMBER_FORMAT_DATE,
    "text": None,
}


# ---------------------------------------------------------------------------
# Sheet helpers (same pattern as price_action_intraday.py / price_action_eod.py)
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
# Scope filter -- §7c
# ---------------------------------------------------------------------------

def in_scope(ws, headers, row_idx):
    if get_cell(ws, headers, row_idx, "M_Eliminated") == NO_TOUCH:
        return False
    div_class = get_cell(ws, headers, row_idx, "M_Div_Coupon_Class")
    return div_class in ELIGIBLE_DIV_COUPON_CLASSES


def resolve_ticker(ws, headers, row_idx):
    yf_ticker = get_cell(ws, headers, row_idx, "S_YF_Ticker")
    if yf_ticker:
        return str(yf_ticker)
    m_ticker = get_cell(ws, headers, row_idx, "M_Ticker")
    return str(m_ticker) if m_ticker else None


# ---------------------------------------------------------------------------
# yfinance fetch -- direct info[] columns
# ---------------------------------------------------------------------------

def _unix_to_date_str(ts):
    """Returns a real date object (not a string) so the Date number_format
    applied via write_cell/FIELD_TYPES actually renders as a date in Excel.
    Not read back elsewhere in this script, so this is a safe conversion."""
    if ts is None:
        return None
    try:
        return datetime.utcfromtimestamp(ts).date()
    except (TypeError, ValueError, OSError):
        return None


def _unix_to_datetime_str(ts):
    if ts is None:
        return None
    try:
        return datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError, OSError):
        return None


def fetch_info_fields(info, country_lookup):
    """Direct 1:1 mapping off a single yfinance .info dict pull."""
    return {
        "S_Name": info.get("longName") or info.get("shortName"),
        "S_Sector": info.get("sector"),
        "S_Country": map_country(info.get("country"), country_lookup),
        "S_Exchange": info.get("exchange"),
        "S_Industry": info.get("industry"),
        "S_Sub_Industry": info.get("industryDisp") or info.get("industry"),
        "S_MCap": info.get("marketCap"),
        "S_Beta": info.get("beta"),
        "S_PE_Ratio": info.get("trailingPE"),
        "S_PayoutRatio": info.get("payoutRatio"),
        "S_DebtEquity": info.get("debtToEquity"),
        "S_ROE": info.get("returnOnEquity"),
        "S_Dividend_Yield": info.get("dividendYield"),
        "S_Average_Volume": info.get("averageVolume"),
        "S_52W_High": info.get("fiftyTwoWeekHigh"),
        "S_52W_Low": info.get("fiftyTwoWeekLow"),
        "S_ExDividend_Date": _unix_to_date_str(info.get("exDividendDate")),
        "S_Reporting_Date": _unix_to_date_str(
            info.get("earningsTimestamp") or info.get("mostRecentQuarter")
        ),
        "S_LastTradedTime": _unix_to_datetime_str(info.get("regularMarketTime")),
    }


# ---------------------------------------------------------------------------
# yfinance fetch -- computed 5Y CAGR / average columns (best-effort)
# ---------------------------------------------------------------------------

def cagr(start_value, end_value, years):
    if not start_value or not end_value or start_value <= 0 or end_value <= 0 or not years:
        return None
    try:
        return ((end_value / start_value) ** (1 / years) - 1) * 100
    except (ZeroDivisionError, ValueError):
        return None


def fetch_computed_fields(ticker_obj, info):
    """
    Best-effort 5Y CAGR / average fields. yfinance doesn't expose these as
    single info[] values -- each requires pulling a history series off the
    same Ticker object and deriving a CAGR/average. Returns None (not an
    exception) for any field where insufficient history exists, so a
    thin/new-listing row degrades gracefully instead of failing the whole
    row.
    """
    result = {col: None for col in S_COLUMNS_COMPUTED}
    current_year = datetime.now().year

    # --- EPS growth 5Y: diluted EPS from annual income statement ---
    try:
        income = ticker_obj.income_stmt
        if income is not None and not income.empty:
            eps_row = None
            for label in ("Diluted EPS", "Basic EPS"):
                if label in income.index:
                    eps_row = income.loc[label].dropna()
                    break
            if eps_row is not None and len(eps_row) >= 2:
                eps_row = eps_row.sort_index()  # oldest -> newest
                span_years = len(eps_row) - 1  # yfinance annual stmts: usually ~4yr, not a true 5Y span
                result["S_EPS_Growth_5Y"] = cagr(eps_row.iloc[0], eps_row.iloc[-1], span_years)
    except Exception:
        pass

    # --- Dividend growth 5Y + Dividend yield 5Y avg: dividend + price history ---
    annual_divs = None
    try:
        dividends = ticker_obj.dividends
        if dividends is not None and not dividends.empty:
            annual = dividends.groupby(dividends.index.year).sum()
            annual = annual[annual.index < current_year]  # drop in-progress year
            annual_divs = annual.tail(6)  # up to 6 year-ends -> up to 5Y span
            if len(annual_divs) >= 2:
                span_years = annual_divs.index[-1] - annual_divs.index[0]
                result["S_DivGrowth_5Y"] = cagr(
                    annual_divs.iloc[0], annual_divs.iloc[-1], span_years
                )
    except Exception:
        pass

    try:
        if annual_divs is not None and len(annual_divs) >= 1:
            hist = ticker_obj.history(period="6y")
            if hist is not None and not hist.empty:
                hist_annual_price = hist["Close"].groupby(hist.index.year).mean()
                yields = []
                for yr, div_sum in annual_divs.items():
                    price = hist_annual_price.get(yr)
                    if price:
                        yields.append(div_sum / price * 100)
                if yields:
                    result["S_DivYield_5YAvg"] = sum(yields) / len(yields)
    except Exception:
        pass

    # --- PE 5Y avg: simplification -- 5Y avg close / current trailing EPS ---
    try:
        trailing_eps = info.get("trailingEps")
        hist = ticker_obj.history(period="5y")
        if trailing_eps and hist is not None and not hist.empty:
            avg_close = hist["Close"].mean()
            result["S_PE_5YAvg"] = avg_close / trailing_eps
    except Exception:
        pass

    # --- Yrs of dividend income to "buy back" 1 share at current price ---
    try:
        price = info.get("currentPrice") or info.get("regularMarketPrice")
        annual_dividend = info.get("dividendRate")
        if price and annual_dividend:
            result["S_Yrs_DivIncome_Buys_1Share"] = price / annual_dividend
    except Exception:
        pass

    return result


# ---------------------------------------------------------------------------
# Combined per-ticker fetch
# ---------------------------------------------------------------------------

def fetch_fundamentals(symbol, country_lookup):
    """
    One yfinance Ticker object per row; .info + .income_stmt + .dividends +
    .history() all pulled from it. Returns (values_dict, gap_rows) where
    values_dict is None on a full-row failure (bad symbol / no data), and
    gap_rows lists any individual S_ columns that resolved to None.
    """
    try:
        t = yf.Ticker(symbol)
        info = t.info
    except Exception as exc:
        print(f"[SKIP] {symbol}: yfinance error: {exc}")
        return None, [{"column": "ALL", "reason": f"yfinance error: {exc}"}]

    if not info or (info.get("regularMarketPrice") is None and info.get("currentPrice") is None):
        print(f"[SKIP] {symbol}: empty/invalid info payload")
        return None, [{"column": "ALL", "reason": "empty or invalid info payload"}]

    values = fetch_info_fields(info, country_lookup)
    values.update(fetch_computed_fields(t, info))

    gaps = [
        {"column": col, "reason": "not available from yfinance for this ticker"}
        for col, val in values.items()
        if val is None
    ]
    return values, gaps


# ---------------------------------------------------------------------------
# Gap logging -- §7c: Outputs/BC/fetch_gaps_YYYYMMDD.csv
# ---------------------------------------------------------------------------

def log_gaps(gap_rows, run_date, output_dir=GAP_OUTPUT_DIR):
    if not gap_rows:
        return None
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"fetch_gaps_{run_date.strftime('%Y%m%d')}.csv"
    file_exists = out_path.exists()
    fieldnames = ["ticker", "column", "reason", "detected_date"]
    with open(out_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerows(gap_rows)
    return out_path


# ---------------------------------------------------------------------------
# Main run
# ---------------------------------------------------------------------------

def run(workbook_path=None, run_date=None, gap_output_dir=GAP_OUTPUT_DIR):
    """
    workbook_path: if omitted, the live .xlsm is resolved by globbing the
    base path (aborts on 0 matches) and the save at the end rotates the
    patch-digit filename, moving the old one to Archive/Workbook/. If an
    explicit path is passed (e.g. by the test harness), it's used as-is
    and saved in place.
    """
    run_date = run_date or datetime.now()
    resolved_by_glob = workbook_path is None
    workbook_path = find_workbook() if resolved_by_glob else Path(workbook_path)

    wb = openpyxl.load_workbook(workbook_path, keep_vba=str(workbook_path).endswith(".xlsm"))
    ws = wb[SHEET_NAME]
    headers = header_map(ws)

    country_lookup = read_legend_lookup_table(wb, "COUNTRY LOOKUP")

    processed, skipped = 0, 0
    all_gaps = []

    for row_idx in range(2, ws.max_row + 1):
        if not in_scope(ws, headers, row_idx):
            continue

        symbol = resolve_ticker(ws, headers, row_idx)
        if not symbol:
            skipped += 1
            continue

        values, gaps = fetch_fundamentals(symbol, country_lookup)
        for g in gaps:
            g["ticker"] = symbol
            g["detected_date"] = run_date.strftime("%Y-%m-%d")
        all_gaps.extend(gaps)

        if values is None:
            skipped += 1
            continue

        for col_name in S_COLUMNS_ALL:
            set_cell(ws, headers, row_idx, col_name, values[col_name])

        processed += 1

    gap_path = log_gaps(all_gaps, run_date, gap_output_dir)

    if resolved_by_glob:
        workbook_path = save_workbook_with_increment(wb, workbook_path)
    else:
        wb.save(workbook_path)

    print(
        f"Run complete. {processed} row(s) updated, {skipped} skipped, "
        f"{len(all_gaps)} gap(s)"
        + (f" logged to {gap_path}." if gap_path else " (none logged).")
        + f" Workbook: {workbook_path.name}"
    )
    return wb


if __name__ == "__main__":
    run()
