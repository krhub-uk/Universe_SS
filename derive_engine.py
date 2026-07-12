"""
derive_engine.py — D_ computation pass
Fifth script in the Portfolio Automation Engine.

Reads S_ and M_ values from Universe, computes all D_ columns,
writes results back via openpyxl surgical cell writes.

NO Excel formulas — all D_ values are Python-computed.
Immune to #REF! errors from column structure changes.

Cadence: runs after fetch_engine_weekly.py. Can also be triggered standalone.
Bolt-on pattern: new derived metrics added as functions in BOLT_ONS registry.

Execution order:
  1. DIVIDENDS aggregates (D_AnnualIncome, D_Income_LTM_GBP, D_Income_YTD_GBP)
  2. QScore sub-components + D_QScore + D_QTier
  3. Price derived (D_PE_Discount, D_Yield_Premium, D_Price_vs_52W_High)
  4. Sell Board signals (Curator → Grower → Chartist → Treasurer → Comptroller → Board)
  5. Income / YOC (D_Sleeve_Weight, D_YOC, D_YoC_PerHolding, D_GBP_Per_GBP1_Div)
  6. Bolt-ons (D_ValueClass etc — append here)

D_ columns that stay in price_action_eod.py (EOD cadence, not touched here):
  D_Cap_Candle_Mid, D_Price_vs_Cap_Mid, D_Close_Direction, D_Volume_Delta_%,
  D_Volume_Flag, D_Context_Flag, D_Cap_Mid_Anchor, D_Cap_Anchor_Date, D_Cap_Anchor_Active

Spec: GDrive/Claude/TradingUniverse/00_Portfolio_Automation_Spec_V4.5.md §7e
"""

import os
import sys
import logging
import openpyxl
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

from workbook_io import find_workbook, write_cell, save_workbook_with_increment

# ── Logging ──────────────────────────────────────────────────────────────────
LOG_PATH = "/var/log/portfolio/derive_engine.log"
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG_PATH), logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


# ── QScore v2 helpers ─────────────────────────────────────────────────────────
# Sprint 3: QScore v2 replaces the prior 4-component weighted model.
# 4 inputs, equal weight, averaged. Any blank input → Insufficient Data.
# Eligibility: M_Eliminated ≠ No Touch, M_Asset_Type ≠ ETF,
# M_Div_Coupon_Class ∈ QSCORE_V2_ELIGIBLE_CLASSES.

QSCORE_V2_ELIGIBLE_CLASSES = {
    "Aristocrat_King", "Aristocrat", "Achiever", "Contender",
    "HighIncome", "MedIncome", "LowNoIncome",
}

SLEEVE_STREAK_TIERS = {"Aristocrat_King", "Aristocrat", "Achiever", "Contender"}

# Streak proxy scores (input 1)
STREAK_PROXY_V2 = {
    "Aristocrat_King": 100,
    "Aristocrat":       80,
    "Achiever":         60,
    "Contender":        50,
    "HighIncome":       40,
    "MedIncome":        30,
    "LowNoIncome":      20,
}

# Kept for non-QScore callers (Sell Board)
ELIGIBLE_CLASSES = QSCORE_V2_ELIGIBLE_CLASSES


def _score_divgrowth_v2(v):
    """DivGrowth_5Y banding: ≥10%=100, 5–10%=80, 2–5%=60, 0–2%=40, <0%=0."""
    if v >= 10.0:  return 100
    if v >= 5.0:   return 80
    if v >= 2.0:   return 60
    if v >= 0.0:   return 40
    return 0


def _score_payout_v2(v):
    """PayoutRatio banding: ≤0.40=100, 0.40–0.60=80, 0.60–0.75=60, 0.75–0.90=40, >0.90=0."""
    if v <= 0.40:  return 100
    if v <= 0.60:  return 80
    if v <= 0.75:  return 60
    if v <= 0.90:  return 40
    return 0


def _score_beta_v2(v):
    """Beta banding: 0.50–0.75=100, 0.75–1.00=80, 1.00–1.25=60, 0.25–0.50=40, <0.25 or >1.25=0."""
    if 0.50 <= v <= 0.75:  return 100
    if 0.75 <  v <= 1.00:  return 80
    if 1.00 <  v <= 1.25:  return 60
    if 0.25 <= v <  0.50:  return 40
    return 0


def compute_qscore(row, cm):
    """
    QScore v2: 4 inputs equal weight, averaged.
    Any blank input → D_QTier = 'Insufficient Data', D_QScore = blank.
    Retired sub-components written as blank (headers kept).
    Sprint 3 replacement for v1 weighted model.
    """
    cls   = row[cm["M_Div_Coupon_Class"]] if "M_Div_Coupon_Class" in cm else None
    elim  = row[cm["M_Eliminated"]]       if "M_Eliminated"       in cm else None
    asset = row[cm["M_Asset_Class"]]      if "M_Asset_Class"      in cm else None

    retired_blank = {
        "D_QScore_DivSafety":     "",
        "D_QScore_Debt":          "",
        "D_QScore_Profitability": "",
        "D_QScore_Stability":     "",
    }

    insufficient = dict(retired_blank, D_QScore="", D_QTier="Insufficient Data")
    ineligible   = dict(retired_blank, D_QScore="", D_QTier="")

    # Eligibility gates
    if elim == "No Touch":
        return ineligible
    if asset != "EQUITIES":
        return ineligible
    if cls not in QSCORE_V2_ELIGIBLE_CLASSES:
        return ineligible

    # Input 1: Streak Proxy (from M_Div_Coupon_Class — always populated if eligible)
    sc_streak = STREAK_PROXY_V2.get(cls)  # None if class somehow missing from map

    # Input 2: DivGrowth_5Y
    dg_raw = row[cm["S_DivGrowth_5Y"]] if "S_DivGrowth_5Y" in cm else None
    sc_divgrowth = _score_divgrowth_v2(float(dg_raw)) if dg_raw is not None else None

    # Input 3: PayoutRatio
    pr_raw = row[cm["S_PayoutRatio"]] if "S_PayoutRatio" in cm else None
    sc_payout = _score_payout_v2(float(pr_raw)) if pr_raw is not None else None

    # Input 4: Beta
    beta_raw = row[cm["S_Beta"]] if "S_Beta" in cm else None
    sc_beta = _score_beta_v2(float(beta_raw)) if beta_raw is not None else None

    # Any blank input → Insufficient Data
    inputs = [sc_streak, sc_divgrowth, sc_payout, sc_beta]
    if any(s is None for s in inputs):
        return insufficient

    score = round(sum(inputs) / 4, 1)
    if score >= 80:
        tier = "A"
    elif score >= 75:
        tier = "B"
    elif score >= 60:
        tier = "C"
    else:
        tier = "Excluded"

    return dict(retired_blank, D_QScore=score, D_QTier=tier)


# ── DIVIDENDS aggregates ──────────────────────────────────────────────────────

def build_dividend_aggregates(wb):
    """
    Read DIVIDENDS tab, return per-ticker dicts:
      all_income[ticker]     = sum of ALL Amount rows, no date filter (inception)
      annual_income[ticker]  = rolling 12 months from today (D_AnnualIncome)
      ltm_income[ticker]     = same as annual_income (alias kept for clarity)
      ytd_income[ticker]     = sum of current calendar year
    Sprint 3: added all_income (D_All_Income); D_AnnualIncome is now rolling
    12m (was previously all-time sum — renamed to all_income).
    """
    ws = wb["DIVIDENDS"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    date_idx   = headers.index("Date")
    amount_idx = headers.index("Amount")
    ticker_idx = headers.index("Ticker")

    today    = date.today()
    ltm_from = today - relativedelta(months=12)
    ytd_from = date(today.year, 1, 1)

    all_time = {}  # D_All_Income: inception to date, no filter
    ltm      = {}  # D_AnnualIncome: rolling 12m
    ytd      = {}  # D_Income_YTD_GBP

    for row in ws.iter_rows(min_row=2, values_only=True):
        ticker = row[ticker_idx]
        amount = row[amount_idx]
        raw_dt = row[date_idx]

        if not ticker or amount is None:
            continue

        try:
            if isinstance(raw_dt, (datetime, date)):
                dt = raw_dt.date() if isinstance(raw_dt, datetime) else raw_dt
            else:
                dt = datetime.strptime(str(raw_dt), "%Y-%m-%d").date()
        except Exception:
            continue

        amount = float(amount)
        all_time[ticker] = all_time.get(ticker, 0) + amount

        if dt >= ltm_from:
            ltm[ticker] = ltm.get(ticker, 0) + amount

        if dt >= ytd_from:
            ytd[ticker] = ytd.get(ticker, 0) + amount

    return all_time, ltm, ytd


# ── Sell Board helpers ────────────────────────────────────────────────────────

def compute_sell_board(row, cm, total_portfolio_value):
    """Compute all five director signals + Board verdict for one row."""
    cls        = row[cm["M_Div_Coupon_Class"]]
    sleeve     = row[cm["M_Sleeve"]]
    div_growth = row[cm["S_DivGrowth_5Y"]]
    chart      = row[cm["M_Chart_Structure_WTF"]]
    better     = row[cm["M_Better_Candidate"]]
    mkt_val    = row[cm["S_MarketValue_GBP"]]

    sleeve_weight = (mkt_val / total_portfolio_value * 100) if (mkt_val and total_portfolio_value) else None

    # D_Curator_Veto
    curator = "HOLDS" if cls in SLEEVE_STREAK_TIERS else ("REVIEW" if cls in ELIGIBLE_CLASSES else "N/A")

    # D_Grower_Signal
    grower = ("FLAG" if (div_growth is not None and div_growth < 0)
              else ("PASS" if div_growth is not None else "N/A"))

    # D_Chartist_Signal
    chartist = ("CONFIRM" if chart == "DOWNTREND"
                else ("NEUTRAL" if chart == "RANGING"
                else ("BLOCK" if chart == "UPTREND" else "N/A")))

    # D_Treasurer_Signal
    treasurer = ("FLAG" if (sleeve_weight is not None and sleeve_weight >= 4.75 and better)
                 else "PASS")

    return {
        "D_Curator_Veto":      curator,
        "D_Grower_Signal":     grower,
        "D_Chartist_Signal":   chartist,
        "D_Treasurer_Signal":  treasurer,
        "D_Sleeve_Weight":     round(sleeve_weight, 4) if sleeve_weight is not None else "",
    }


def compute_comptroller_ranks(universe_rows, cm):
    """
    Rank all DIV_CORE rows by D_GBP_Per_GBP1_Div (ascending — lower is more efficient).
    Returns dict: row_idx → (rank, signal)
    """
    div_core_rows = []
    for i, row in enumerate(universe_rows):
        if row[cm["M_Sleeve"]] == "DIV_CORE" and row[cm["M_Eliminated"]] != "No Touch":
            cost  = row[cm["S_CostBasis"]]
            ann   = row[cm.get("__annual_income_computed__", -1)] if "__annual_income_computed__" in cm else None
            div_core_rows.append((i, cost, ann))

    # We'll compute this after D_AnnualIncome is written
    # Return placeholder — caller must handle this
    return {}


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    start = datetime.now()
    log.info("derive_engine.py started")

    wb_path = find_workbook()
    wb = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)
    ws_uni = wb['Universe']
    headers = [c.value for c in next(ws_uni.iter_rows(min_row=1, max_row=1))]
    cm = {h: i for i, h in enumerate(headers) if h}

    # ── Step 1: DIVIDENDS aggregates ─────────────────────────────────────────
    log.info("Step 1: DIVIDENDS aggregates")
    all_time_map, ltm_map, ytd_map = build_dividend_aggregates(wb)

    # ── Load all Universe rows into memory ────────────────────────────────────
    all_rows = list(ws_uni.iter_rows(min_row=2, values_only=True))
    log.info(f"Universe rows: {len(all_rows)}")

    # Track per-row computed values needed for downstream steps
    computed = [{} for _ in all_rows]

    # ── Step 1 write + capture for downstream ────────────────────────────────
    # Sprint 3: D_AnnualIncome is now rolling 12m; D_All_Income is all-time.
    # D_Income_Payback_Pct = D_All_Income / S_CostBasis (cumulative payback).
    # D_YoC = D_AnnualIncome / S_CostBasis (blended YoC on current holding).
    annual_income_by_row = {}  # row_idx → rolling-12m income (for downstream YoC/Comptroller)
    total_portfolio_value = 0.0
    for i, row in enumerate(all_rows):
        ticker = row[cm["M_Ticker"]]
        if not ticker:
            continue
        all_income = all_time_map.get(ticker, 0)
        ann_income = ltm_map.get(ticker, 0)   # rolling 12m = D_AnnualIncome
        ytd        = ytd_map.get(ticker, 0)

        computed[i]["D_All_Income"]      = all_income if all_income else ""
        computed[i]["D_AnnualIncome"]    = ann_income if ann_income else ""
        computed[i]["D_Income_LTM_GBP"]  = ann_income if ann_income else ""  # alias
        computed[i]["D_Income_YTD_GBP"]  = ytd if ytd else ""
        annual_income_by_row[i]          = ann_income

        # D_Income_Payback_Pct = D_All_Income / S_CostBasis
        cost_basis = row[cm["S_CostBasis"]] if "S_CostBasis" in cm else None
        if all_income and cost_basis:
            try:
                computed[i]["D_Income_Payback_Pct"] = round(all_income / float(cost_basis), 4)
            except (TypeError, ValueError, ZeroDivisionError):
                computed[i]["D_Income_Payback_Pct"] = ""
        else:
            computed[i]["D_Income_Payback_Pct"] = ""

        # D_YOC is left unchanged (yfinance yield — written in Step 5, not here).

        mkt_val = row[cm["S_MarketValue_GBP"]] if "S_MarketValue_GBP" in cm else None
        if mkt_val:
            total_portfolio_value += float(mkt_val)

    log.info(f"Total portfolio value for sleeve weighting: £{total_portfolio_value:,.2f}")

    # ── QScore trajectory: shift D_QScore → S_QScore_Prior before computing new QScore ──
    # Sprint 3: S_QScore_Prior replaces M_QScore_Prior in the column header.
    # We read the current D_QScore (from last run, stored in the workbook)
    # and write it to S_QScore_Prior before overwriting D_QScore below.
    log.info("Step 1b: QScore trajectory shift (D_QScore → S_QScore_Prior)")
    prior_col = "S_QScore_Prior"
    current_d_col = "D_QScore"
    if prior_col in cm and current_d_col in cm:
        for i, row in enumerate(all_rows):
            current_qscore = row[cm[current_d_col]]
            if current_qscore is not None and current_qscore != "":
                computed[i][prior_col] = current_qscore
            # If blank/None, leave S_QScore_Prior as-is (don't overwrite with blank)
    elif "M_QScore_Prior" in cm and current_d_col in cm:
        # Fallback: workbook header not yet renamed — write to M_QScore_Prior column
        log.warning("S_QScore_Prior not found in Universe headers; falling back to M_QScore_Prior")
        for i, row in enumerate(all_rows):
            current_qscore = row[cm[current_d_col]]
            if current_qscore is not None and current_qscore != "":
                computed[i]["M_QScore_Prior"] = current_qscore
    else:
        log.warning("Neither S_QScore_Prior nor M_QScore_Prior found — trajectory shift skipped.")

    # ── Step 2: QScore ────────────────────────────────────────────────────────
    log.info("Step 2: QScore")
    scored = ineligible = 0
    for i, row in enumerate(all_rows):
        result = compute_qscore(row, cm)
        computed[i].update(result)
        if result["D_QScore"] != "":
            scored += 1
        else:
            ineligible += 1
    log.info(f"QScore: {scored} scored, {ineligible} ineligible/blank")

    # ── Step 3: Price derived ─────────────────────────────────────────────────
    log.info("Step 3: Price derived")
    for i, row in enumerate(all_rows):
        if row[cm["M_Eliminated"]] == "No Touch":
            computed[i].update({"D_PE_Discount": "", "D_Yield_Premium": "", "D_Price_vs_52W_High": ""})
            continue

        pe     = row[cm["S_PE_Ratio"]]
        pe_avg = row[cm["S_PE_5YAvg"]]
        yld    = row[cm["S_Dividend_Yield"]]
        yavg   = row[cm["S_DivYield_5YAvg"]]
        price  = row[cm["S_Last_Price"]]
        high   = row[cm["S_52W_High"]]

        pe_disc = round((pe - pe_avg) / pe_avg, 4) if (pe and pe_avg and pe_avg != 0) else ""
        yld_prem = round(yld - yavg, 4) if (yld is not None and yavg is not None) else ""
        p52 = round((price - high) / high, 4) if (price and high and high != 0) else ""

        computed[i].update({
            "D_PE_Discount":       pe_disc,
            "D_Yield_Premium":     yld_prem,
            "D_Price_vs_52W_High": p52,
        })

    # ── Step 4: Sell Board signals ────────────────────────────────────────────
    log.info("Step 4: Sell Board signals")
    for i, row in enumerate(all_rows):
        if row[cm["M_Eliminated"]] == "No Touch":
            computed[i].update({k: "" for k in [
                "D_Curator_Veto","D_Grower_Signal","D_Chartist_Signal",
                "D_Treasurer_Signal","D_Sleeve_Weight"]})
            continue
        sb = compute_sell_board(row, cm, total_portfolio_value)
        computed[i].update(sb)

    # Comptroller rank across DIV_CORE (needs D_GBP_Per_GBP1_Div — computed in step 5)
    # We'll compute D_GBP_Per_GBP1_Div first, then rank

    # ── Step 5: Income / YOC + D_GBP_Per_GBP1_Div ───────────────────────────
    log.info("Step 5: Income / YOC")
    gbp_per_gbp1 = {}  # row_idx → value, for Comptroller ranking
    for i, row in enumerate(all_rows):
        if row[cm["M_Eliminated"]] == "No Touch":
            computed[i].update({k: "" for k in ["D_YOC","D_YoC_PerHolding","D_GBP_Per_GBP1_Div"]})
            continue

        avg_cost  = row[cm["S_AvgCost"]]   if "S_AvgCost"   in cm else None
        cost_bas  = row[cm["S_CostBasis"]] if "S_CostBasis" in cm else None
        yld       = row[cm["S_Dividend_Yield"]] if "S_Dividend_Yield" in cm else None
        ann_inc   = annual_income_by_row.get(i, 0)

        yoc         = round(yld, 4) if yld is not None else ""  # yfinance yield — unchanged
        yoc_holding = round(ann_inc / float(cost_bas), 4) if (ann_inc and cost_bas) else ""
        gbp_eff     = round(float(cost_bas) / ann_inc, 2) if (ann_inc and cost_bas and ann_inc > 0) else ""

        computed[i].update({
            "D_YOC":              yoc,
            "D_YoC_PerHolding":   yoc_holding,
            "D_GBP_Per_GBP1_Div": gbp_eff,
        })
        if isinstance(gbp_eff, float):
            gbp_per_gbp1[i] = gbp_eff

    # ── Comptroller rank (across DIV_CORE only) ───────────────────────────────
    log.info("Step 4b: Comptroller rank")
    div_core_vals = {
        i: v for i, v in gbp_per_gbp1.items()
        if all_rows[i][cm["M_Sleeve"]] == "DIV_CORE"
    }
    if div_core_vals:
        sorted_vals = sorted(div_core_vals.values())
        n = len(sorted_vals)
        p75_threshold = sorted_vals[int(n * 0.75)]

        for i, row in enumerate(all_rows):
            if row[cm["M_Sleeve"]] != "DIV_CORE" or i not in gbp_per_gbp1:
                computed[i].update({"D_Comptroller_Rank": "", "D_Comptroller_Signal": ""})
                continue
            val = gbp_per_gbp1[i]
            rank = sorted(div_core_vals.values()).index(val) + 1
            signal = "FLAG" if val >= p75_threshold else "PASS"
            computed[i].update({
                "D_Comptroller_Rank":   rank,
                "D_Comptroller_Signal": signal,
            })
    else:
        for i in range(len(all_rows)):
            computed[i].update({"D_Comptroller_Rank": "", "D_Comptroller_Signal": ""})

    # ── Board verdict ─────────────────────────────────────────────────────────
    log.info("Step 4c: Board verdict")
    flag_cols = ["D_Comptroller_Signal","D_Grower_Signal","D_Chartist_Signal","D_Treasurer_Signal"]
    for i in range(len(all_rows)):
        curator = computed[i].get("D_Curator_Veto", "")
        if not curator:
            computed[i].update({"D_Board_Flags": "", "D_Board_Verdict": "N/A"})
            continue
        flags = sum(1 for k in flag_cols if computed[i].get(k) == "FLAG")
        computed[i]["D_Board_Flags"] = flags
        if curator == "HOLDS":
            computed[i]["D_Board_Verdict"] = "VETOED"
        elif flags >= 3:
            computed[i]["D_Board_Verdict"] = "SELL REVIEW"
        else:
            computed[i]["D_Board_Verdict"] = "HOLD"

    # ── Step 6: Bolt-ons ──────────────────────────────────────────────────────
    # D_ValueClass — TBD, write blank placeholder
    for i in range(len(all_rows)):
        computed[i]["D_ValueClass"] = ""

    # ── Write all computed values back to sheet ───────────────────────────────
    log.info("Writing D_ values to Universe sheet")

    written = skipped = 0

    for i, vals in enumerate(computed):
        excel_row = i + 2  # 1-indexed, row 1 = header
        for col_name, value in vals.items():
            if col_name not in cm:
                continue
            col_idx = cm[col_name] + 1  # 1-indexed for openpyxl
            write_cell(ws_uni, excel_row, col_idx, value)
            written += 1

    log.info(f"Cells written: {written} | Rows skipped (no ticker): {skipped}")

    # ── Save ──────────────────────────────────────────────────────────────────
    save_workbook_with_increment(wb, wb_path)
    log.info("Workbook saved")

    end = datetime.now()
    elapsed = str(end - start).split(".")[0]
    log.info(f"--- Started: {start.strftime('%Y-%m-%d %H:%M:%S')} | "
             f"Ended: {end.strftime('%H:%M:%S')} | Duration: {elapsed} ---")


if __name__ == "__main__":
    run()
