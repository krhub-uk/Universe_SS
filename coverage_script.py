"""
coverage_script.py -- COVERAGE tab population
Spec: 00_Portfolio_Automation_Spec_V5.5.md, Wave 7. Sprint 4.

Runs: weekly cadence (called from run_weekly.sh after derive_engine.py).

Populates the COVERAGE tab with two blocks:
  1. SECTOR COVERAGE  -- Universe rows grouped by S_Sector, S_Invested="Y" only
  2. PORTFOLIO COVERAGE -- Universe rows grouped by M_Sleeve, S_Invested="Y" only

Clear-and-rewrite on each run. No Excel formula dependency.

SECTOR COVERAGE columns (16 sectors + TOTAL):
  Sector, Holdings, Market Value GBP, Portfolio %, Cost GBP, G/L GBP, G/L %

PORTFOLIO COVERAGE columns (sleeves from Lookups Allocations + TOTAL):
  Sleeve, Holdings, Market Value GBP, Portfolio %, vs Target, Cost GBP, G/L GBP

S_Invested = "Y" filter applied throughout.
Sleeve target % from Lookups Allocations table: Allocations / Alloc_Pct columns.
Target stored as raw integer in Lookups (35 = 35%) -- divide by 100 for vs Target.

Wave 8: Structured PHASE/ACTION/TICKER/RUN_ID/UID logging.
"""

import logging
import sys
import uuid
from datetime import datetime

import openpyxl
from openpyxl.styles import Font

from workbook_io import (
    BASE_PATH,
    find_workbook,
    save_workbook_with_increment,
    write_cell,
)

# ---------------------------------------------------------------------------
# Logging (Wave 8)
# ---------------------------------------------------------------------------

LOG_DIR = BASE_PATH / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "coverage_script.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("coverage_script")

_ctx: dict = {}
_VOCAB_KEYS = ("PHASE=", "ACTION=", "TICKER=", "RUN_ID=", "UID=")


def _log(level: str, phase: str, action: str, ticker: str, message: str) -> None:
    run_id = _ctx.get("run_id", "")
    uid    = _ctx.get("uid", "")
    line   = (
        f"PHASE={phase} ACTION={action} TICKER={ticker} "
        f"RUN_ID={run_id} UID={uid} | {message}"
    )
    for key in _VOCAB_KEYS:
        if key not in line:
            log.warning(f"[VOCAB_FAIL] Missing {key}")
    getattr(log, level.lower())(line)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

WRITTEN_FONT_SIZE = 12


def header_map(ws):
    """{header_name: 0-based_index} from row 1 values."""
    return {
        cell.value: cell.column - 1
        for cell in ws[1]
        if cell.value is not None
    }


def _safe_float(v, default=0.0):
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _write(ws, row, col_1based, value):
    """Thin wrapper: writes value using openpyxl directly, font size 12."""
    cell = ws.cell(row=row, column=col_1based)
    cell.value = value
    f = cell.font
    cell.font = Font(
        name=f.name, size=WRITTEN_FONT_SIZE, bold=f.bold, italic=f.italic,
        vertAlign=f.vertAlign, underline=f.underline, strike=f.strike,
        color=f.color,
    )


def clear_coverage_tab(ws):
    """Delete all data rows (keep row 1 header if present, else clear all)."""
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)


# ---------------------------------------------------------------------------
# Load Universe data
# ---------------------------------------------------------------------------

def load_universe_rows(wb):
    """
    Returns a list of dicts for every Universe row where S_Invested = "Y".
    Keys: S_Sector, M_Sleeve, S_MarketValue_GBP, S_CostBasis, S_PnL_GBP, S_PnL_Pct
    """
    ws = wb["Universe"]
    cm = header_map(ws)

    needed = [
        "S_Invested", "S_Sector", "M_Sleeve",
        "S_MarketValue_GBP", "S_CostBasis", "S_PnL_GBP", "S_PnL_Pct",
    ]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # S_Invested filter
        invested = row[cm["S_Invested"]] if "S_Invested" in cm else None
        if invested != "Y":
            continue

        rows.append({
            "S_Sector":         row[cm["S_Sector"]]         if "S_Sector"         in cm else None,
            "M_Sleeve":         row[cm["M_Sleeve"]]         if "M_Sleeve"         in cm else None,
            "S_MarketValue_GBP": row[cm["S_MarketValue_GBP"]] if "S_MarketValue_GBP" in cm else None,
            "S_CostBasis":      row[cm["S_CostBasis"]]      if "S_CostBasis"      in cm else None,
            "S_PnL_GBP":        row[cm["S_PnL_GBP"]]        if "S_PnL_GBP"        in cm else None,
            "S_PnL_Pct":        row[cm["S_PnL_Pct"]]        if "S_PnL_Pct"        in cm else None,
        })

    return rows


# ---------------------------------------------------------------------------
# Load Lookups Allocations table
# ---------------------------------------------------------------------------

def load_allocations(wb):
    """
    Reads Lookups tab Allocations table by column name.
    Returns dict: {sleeve_name: alloc_pct_raw} where alloc_pct_raw is 35 (not 0.35).
    """
    ws   = wb["Lookups"]
    rows = list(ws.iter_rows(values_only=True))

    # Find header row that contains both 'Allocations' and 'Alloc_Pct'
    header_row_idx = None
    for i, row in enumerate(rows):
        if row and "Allocations" in row and "Alloc_Pct" in row:
            header_row_idx = i
            break

    if header_row_idx is None:
        _log("warning", "COVERAGE", "ALLOC_MISSING", "",
             "Lookups tab: could not find Allocations/Alloc_Pct header row")
        return {}

    header = rows[header_row_idx]
    alloc_col = list(header).index("Allocations")
    pct_col   = list(header).index("Alloc_Pct")

    allocations = {}
    for row in rows[header_row_idx + 1:]:
        if not row or row[alloc_col] is None:
            break
        sleeve = row[alloc_col]
        pct    = row[pct_col]
        if sleeve and pct is not None:
            allocations[str(sleeve).strip()] = _safe_float(pct)

    return allocations


# ---------------------------------------------------------------------------
# Aggregate helpers
# ---------------------------------------------------------------------------

def _aggregate(filtered_rows):
    """Sum holdings, mkt_val, cost, pnl_gbp for a list of rows."""
    holdings = len(filtered_rows)
    mkt_val  = sum(_safe_float(r["S_MarketValue_GBP"]) for r in filtered_rows)
    cost     = sum(_safe_float(r["S_CostBasis"])        for r in filtered_rows)
    pnl_gbp  = sum(_safe_float(r["S_PnL_GBP"])          for r in filtered_rows)
    return holdings, mkt_val, cost, pnl_gbp


def _pnl_pct(cost, pnl_gbp):
    if cost and cost != 0:
        return round(pnl_gbp / cost * 100, 2)
    return None


# ---------------------------------------------------------------------------
# SECTOR COVERAGE block
# ---------------------------------------------------------------------------

SECTOR_ORDER = [
    "Basic Materials", "Communication Services", "Consumer Cyclical",
    "Consumer Defensive", "Energy", "Financial Services",
    "Healthcare", "Industrials", "Real Estate",
    "Technology", "Utilities",
    # Catch-alls for any others
    "Unknown", "Other",
]


def build_sector_coverage(all_rows):
    """
    Returns list of dicts, one per sector, plus TOTAL.
    Sectors driven by what exists in the data; TOTAL appended last.
    """
    sector_map = {}
    for row in all_rows:
        sector = str(row["S_Sector"] or "Unknown").strip() or "Unknown"
        sector_map.setdefault(sector, []).append(row)

    # Sort sectors: known order first, then any extras alphabetically
    known = [s for s in SECTOR_ORDER if s in sector_map]
    extra = sorted(s for s in sector_map if s not in SECTOR_ORDER)
    ordered_sectors = known + extra

    total_mkt = sum(_safe_float(r["S_MarketValue_GBP"]) for r in all_rows)

    result = []
    for sector in ordered_sectors:
        rows = sector_map[sector]
        holdings, mkt_val, cost, pnl_gbp = _aggregate(rows)
        portfolio_pct = round(mkt_val / total_mkt * 100, 2) if total_mkt else None
        pnl_pct       = _pnl_pct(cost, pnl_gbp)
        result.append({
            "label":         sector,
            "holdings":      holdings,
            "mkt_val":       round(mkt_val, 2),
            "portfolio_pct": portfolio_pct,
            "cost":          round(cost, 2),
            "pnl_gbp":       round(pnl_gbp, 2),
            "pnl_pct":       pnl_pct,
        })

    # TOTAL row
    t_holdings, t_mkt, t_cost, t_pnl = _aggregate(all_rows)
    result.append({
        "label":         "TOTAL",
        "holdings":      t_holdings,
        "mkt_val":       round(t_mkt, 2),
        "portfolio_pct": 100.0 if all_rows else None,
        "cost":          round(t_cost, 2),
        "pnl_gbp":       round(t_pnl, 2),
        "pnl_pct":       _pnl_pct(t_cost, t_pnl),
    })

    return result


# ---------------------------------------------------------------------------
# PORTFOLIO COVERAGE block
# ---------------------------------------------------------------------------

def build_portfolio_coverage(all_rows, allocations):
    """
    Returns list of dicts, one per sleeve (from Lookups Allocations), plus TOTAL.
    vs_target = Portfolio% - (Alloc_Pct / 100)
    """
    total_mkt = sum(_safe_float(r["S_MarketValue_GBP"]) for r in all_rows)

    sleeve_map = {}
    for row in all_rows:
        sleeve = str(row["M_Sleeve"] or "Unknown").strip() or "Unknown"
        sleeve_map.setdefault(sleeve, []).append(row)

    # Ordered by Lookups Allocations table; extras appended alphabetically
    known_sleeves = list(allocations.keys())
    extra_sleeves = sorted(s for s in sleeve_map if s not in allocations)
    ordered_sleeves = known_sleeves + extra_sleeves

    result = []
    for sleeve in ordered_sleeves:
        rows = sleeve_map.get(sleeve, [])
        holdings, mkt_val, cost, pnl_gbp = _aggregate(rows)
        portfolio_pct = round(mkt_val / total_mkt * 100, 2) if total_mkt else None
        alloc_raw     = allocations.get(sleeve)  # e.g. 35 (= 35%)
        alloc_pct     = alloc_raw / 100 if alloc_raw is not None else None
        vs_target     = (
            round(portfolio_pct - alloc_pct * 100, 2)  # portfolio% - target%
            if (portfolio_pct is not None and alloc_pct is not None)
            else None
        )
        result.append({
            "label":         sleeve,
            "holdings":      holdings,
            "mkt_val":       round(mkt_val, 2),
            "portfolio_pct": portfolio_pct,
            "vs_target":     vs_target,
            "cost":          round(cost, 2),
            "pnl_gbp":       round(pnl_gbp, 2),
        })

    # TOTAL row
    t_holdings, t_mkt, t_cost, t_pnl = _aggregate(all_rows)
    result.append({
        "label":         "TOTAL",
        "holdings":      t_holdings,
        "mkt_val":       round(t_mkt, 2),
        "portfolio_pct": 100.0 if all_rows else None,
        "vs_target":     None,
        "cost":          round(t_cost, 2),
        "pnl_gbp":       round(t_pnl, 2),
    })

    return result


# ---------------------------------------------------------------------------
# Write to COVERAGE tab
# ---------------------------------------------------------------------------

def write_coverage_tab(wb, sector_rows, portfolio_rows):
    """
    Clear and rewrite COVERAGE tab.
    Layout:
      Row 1: blank (or existing header preserved -- we clear all and rewrite)
      Row 2: SECTOR COVERAGE header
      Row 3..N: sector data
      Row N+1: blank separator
      Row N+2: PORTFOLIO COVERAGE header
      Row N+3..M: portfolio data
    """
    ws = wb["COVERAGE"]

    # Clear all content
    ws.delete_rows(1, ws.max_row)

    current_row = 1

    # --- SECTOR COVERAGE ---
    sector_headers = [
        "Sector", "Holdings", "Market Value GBP",
        "Portfolio %", "Cost GBP", "G/L GBP", "G/L %",
    ]
    for col_i, hdr in enumerate(sector_headers, start=1):
        _write(ws, current_row, col_i, hdr)
    current_row += 1

    for row_data in sector_rows:
        _write(ws, current_row, 1, row_data["label"])
        _write(ws, current_row, 2, row_data["holdings"])
        _write(ws, current_row, 3, row_data["mkt_val"])
        _write(ws, current_row, 4, row_data["portfolio_pct"])
        _write(ws, current_row, 5, row_data["cost"])
        _write(ws, current_row, 6, row_data["pnl_gbp"])
        _write(ws, current_row, 7, row_data["pnl_pct"])
        current_row += 1

    current_row += 1  # blank separator row

    # --- PORTFOLIO COVERAGE ---
    portfolio_headers = [
        "Sleeve", "Holdings", "Market Value GBP",
        "Portfolio %", "vs Target", "Cost GBP", "G/L GBP",
    ]
    for col_i, hdr in enumerate(portfolio_headers, start=1):
        _write(ws, current_row, col_i, hdr)
    current_row += 1

    for row_data in portfolio_rows:
        _write(ws, current_row, 1, row_data["label"])
        _write(ws, current_row, 2, row_data["holdings"])
        _write(ws, current_row, 3, row_data["mkt_val"])
        _write(ws, current_row, 4, row_data["portfolio_pct"])
        _write(ws, current_row, 5, row_data["vs_target"])
        _write(ws, current_row, 6, row_data["cost"])
        _write(ws, current_row, 7, row_data["pnl_gbp"])
        current_row += 1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run():
    _ctx["run_id"] = str(uuid.uuid4())[:8]
    _ctx["uid"]    = ""

    start = datetime.now()
    _log("info", "STARTUP", "RUN_START", "", "coverage_script.py starting")

    wb_path = find_workbook()
    wb      = openpyxl.load_workbook(wb_path, keep_vba=True)

    _log("info", "COVERAGE", "UNIVERSE_LOAD", "", "Loading Universe invested rows")
    all_rows = load_universe_rows(wb)
    _log("info", "COVERAGE", "UNIVERSE_LOAD", "", f"{len(all_rows)} S_Invested=Y rows")

    _log("info", "COVERAGE", "ALLOC_LOAD", "", "Loading Lookups Allocations table")
    allocations = load_allocations(wb)
    _log("info", "COVERAGE", "ALLOC_LOAD", "", f"{len(allocations)} sleeve allocations loaded")

    _log("info", "COVERAGE", "SECTOR_BUILD", "", "Building SECTOR COVERAGE block")
    sector_rows = build_sector_coverage(all_rows)

    _log("info", "COVERAGE", "PORTFOLIO_BUILD", "", "Building PORTFOLIO COVERAGE block")
    portfolio_rows = build_portfolio_coverage(all_rows, allocations)

    _log("info", "COVERAGE", "WRITE_START", "", "Writing COVERAGE tab (clear-and-rewrite)")
    write_coverage_tab(wb, sector_rows, portfolio_rows)

    save_workbook_with_increment(wb, wb_path)

    elapsed = str(datetime.now() - start).split(".")[0]
    _log("info", "COMPLETE", "RUN_END", "",
         f"{len(sector_rows)} sector rows, {len(portfolio_rows)} portfolio rows. "
         f"Duration {elapsed}")


if __name__ == "__main__":
    run()
