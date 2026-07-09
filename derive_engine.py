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


# ── QScore helpers ────────────────────────────────────────────────────────────

ELIGIBLE_CLASSES = {"Aristocrat_King", "Aristocrat", "Achiever", "Contender",
                    "HighIncome", "MedIncome"}

STREAK_PROXY = {
    "Aristocrat_King": 100,
    "Aristocrat":      100,
    "Achiever":         50,
    "Contender":        50,
    "HighIncome":        0,
    "MedIncome":         0,
}

SLEEVE_STREAK_TIERS = {"Aristocrat_King", "Aristocrat", "Achiever", "Contender"}


def _score(value, good_fn, ok_fn):
    """Return 100/50/0 based on threshold functions, or None if value is None."""
    if value is None:
        return None
    if good_fn(value):
        return 100
    if ok_fn(value):
        return 50
    return 0


def _sub_score(scores):
    """
    Average a list of scores, ignoring None.
    If ≥2 blanks in a 2-input sub-component, or all blank, return None.
    Spec: ≤1 blank → score remaining; 2+ blanks → sub-component blank.
    """
    valid = [s for s in scores if s is not None]
    if not valid:
        return None
    blanks = len(scores) - len(valid)
    if blanks >= 2:
        return None
    return round(sum(valid) / len(valid), 2)


def compute_qscore(row, cm):
    """Compute D_QScore_* sub-components + D_QScore + D_QTier for one row."""
    cls   = row[cm["M_Div_Coupon_Class"]]
    elim  = row[cm["M_Eliminated"]]
    asset = row[cm["M_Asset_Class"]]
    dyld  = row[cm["S_Dividend_Yield"]]

    blank = {k: "" for k in ["D_QScore_DivSafety", "D_QScore_Debt",
                               "D_QScore_Profitability", "D_QScore_Stability",
                               "D_QScore", "D_QTier"]}

    # Eligibility gates
    if elim == "No Touch":
        return blank
    if asset != "EQUITIES":
        return blank
    if cls not in ELIGIBLE_CLASSES:
        return blank
    if not dyld or dyld < 1.0:
        return blank

    # DivSafety (3 inputs)
    payout    = row[cm["S_PayoutRatio"]]
    streak_sc = STREAK_PROXY.get(cls)          # always populated if eligible
    divgrowth = row[cm["S_DivGrowth_5Y"]]

    sc_payout = _score(payout,
                       lambda v: v <= 0.75,
                       lambda v: v <= 0.90) if payout is not None else None
    sc_divg   = _score(divgrowth,
                       lambda v: v >= 5.0,
                       lambda v: v >= 0.0) if divgrowth is not None else None

    div_safety = _sub_score([sc_payout, streak_sc, sc_divg])

    # Debt (1 input)
    de = row[cm["S_DebtEquity"]]
    sc_de = _score(de, lambda v: v <= 1.0, lambda v: v <= 2.0) if de is not None else None
    debt = sc_de  # single input — blank if None

    # Profitability (2 inputs)
    roe      = row[cm["S_ROE"]]
    eps_grow = row[cm["S_EPS_Growth_5Y"]]
    sc_roe  = _score(roe,      lambda v: v >= 15.0, lambda v: v >= 5.0) if roe      is not None else None
    sc_eps  = _score(eps_grow, lambda v: v >= 5.0,  lambda v: v >= 0.0) if eps_grow is not None else None
    profit  = _sub_score([sc_roe, sc_eps])

    # Stability (2 inputs)
    beta      = row[cm["S_Beta"]]
    price     = row[cm["S_Last_Price"]]
    high_52w  = row[cm["S_52W_High"]]

    sc_beta = None
    if beta is not None:
        if 0.5 <= beta <= 1.0:
            sc_beta = 100
        elif 0.0 <= beta < 0.5 or 1.0 < beta <= 1.3:
            sc_beta = 50
        else:
            sc_beta = 0

    sc_p52 = None
    if price is not None and high_52w and high_52w > 0:
        ratio = price / high_52w
        sc_p52 = _score(ratio, lambda v: v >= 0.80, lambda v: v >= 0.60)

    stability = _sub_score([sc_beta, sc_p52])

    # Weighted total
    components = {"DivSafety": (div_safety, 0.35), "Debt": (debt, 0.25),
                  "Profitability": (profit, 0.20), "Stability": (stability, 0.20)}

    valid_comps = {k: (v, w) for k, (v, w) in components.items() if v is not None}
    if not valid_comps:
        return blank

    # Reweight if any component is blank
    total_weight = sum(w for _, w in valid_comps.values())
    score = sum(v * (w / total_weight) for _, (v, w) in valid_comps.items())
    score = round(score, 1)

    tier = "A" if score >= 80 else "B" if score >= 75 else "C" if score >= 60 else "Excluded"

    return {
        "D_QScore_DivSafety":    round(div_safety,  1) if div_safety  is not None else "",
        "D_QScore_Debt":         round(debt,         1) if debt        is not None else "",
        "D_QScore_Profitability":round(profit,       1) if profit      is not None else "",
        "D_QScore_Stability":    round(stability,    1) if stability   is not None else "",
        "D_QScore":              score,
        "D_QTier":               tier,
    }


# ── DIVIDENDS aggregates ──────────────────────────────────────────────────────

def build_dividend_aggregates(wb):
    """
    Read DIVIDENDS tab, return per-ticker dicts:
      annual_income[ticker]  = sum of all Amount rows (annualised proxy = LTM sum)
      ltm_income[ticker]     = sum of last 12 calendar months
      ytd_income[ticker]     = sum of current calendar year
    """
    ws = wb["DIVIDENDS"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    date_idx   = headers.index("Date")
    amount_idx = headers.index("Amount")
    ticker_idx = headers.index("Ticker")

    today    = date.today()
    ltm_from = today - relativedelta(months=12)
    ytd_from = date(today.year, 1, 1)

    annual = {}
    ltm    = {}
    ytd    = {}

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
        annual[ticker] = annual.get(ticker, 0) + amount

        if dt >= ltm_from:
            ltm[ticker] = ltm.get(ticker, 0) + amount

        if dt >= ytd_from:
            ytd[ticker] = ytd.get(ticker, 0) + amount

    return annual, ltm, ytd


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
    annual_map, ltm_map, ytd_map = build_dividend_aggregates(wb)

    # ── Load all Universe rows into memory ────────────────────────────────────
    all_rows = list(ws_uni.iter_rows(min_row=2, values_only=True))
    log.info(f"Universe rows: {len(all_rows)}")

    # Track per-row computed values needed for downstream steps
    computed = [{} for _ in all_rows]

    # ── Step 1 write + capture for downstream ────────────────────────────────
    annual_income_by_row = {}
    total_portfolio_value = 0.0
    for i, row in enumerate(all_rows):
        ticker = row[cm["M_Ticker"]]
        if not ticker:
            continue
        ann = annual_map.get(ticker, 0)
        ltm = ltm_map.get(ticker, 0)
        ytd = ytd_map.get(ticker, 0)

        computed[i]["D_AnnualIncome"]    = ann if ann else ""
        computed[i]["D_Income_LTM_GBP"] = ltm if ltm else ""
        computed[i]["D_Income_YTD_GBP"] = ytd if ytd else ""
        annual_income_by_row[i] = ann

        mkt_val = row[cm["S_MarketValue_GBP"]]
        if mkt_val:
            total_portfolio_value += float(mkt_val)

    log.info(f"Total portfolio value for sleeve weighting: £{total_portfolio_value:,.2f}")

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

        avg_cost  = row[cm["S_AvgCost"]]
        cost_bas  = row[cm["S_CostBasis"]]
        yld       = row[cm["S_Dividend_Yield"]]
        ann_inc   = annual_income_by_row.get(i, 0)

        yoc          = round(yld, 4) if yld is not None else ""  # yfinance yield = annual %
        yoc_holding  = round(ann_inc / float(cost_bas), 4) if (ann_inc and cost_bas) else ""
        gbp_eff      = round(float(cost_bas) / ann_inc, 2) if (ann_inc and cost_bas and ann_inc > 0) else ""

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
