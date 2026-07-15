# Universe_SS — Portfolio Automation Engine

Systematic portfolio management engine for a SIPP-based investment universe. Fetches price action, fundamentals, and derived analytics across a 570+ row ticker universe, writes everything back into a single Excel workbook, and generates TradingView watchlists on demand.

---

## Architecture

**Excel is the source of truth.** Python fetches, computes, and writes — it never restructures the workbook. All column references are by header name, not position, so the sheet can evolve without breaking scripts.

**Column prefix convention:**
- `M_` Manual / static — set by hand, never overwritten by scripts
- `S_` Sourced / synced — written by fetch scripts
- `D_` Derived — written by `derive_engine.py` only, never Excel formulas

---

## Scripts

### Cron scripts (automated)

| Script | Cadence | Purpose |
|---|---|---|
| `price_action_intraday.py` | 07:00 11:00 15:00 19:00 weekdays | 9 price columns. Lightest yfinance call — no `info[]`, no CSV. |
| `price_action_eod.py` | 23:00 weekdays | 15 price columns + 10 D_ EOD write-backs. Triggers CSV ingestion if input files are present. |
| `fetch_engine_weekly.py` | 23:30 Friday | Full fundamentals across all eligible tickers. yfinance `info[]` + computed 5Y CAGRs. |
| `derive_engine.py` | After weekly fetch | All D_ column computation — QScore, Sell Board signals, Income/YOC, price derived. Bolt-on pattern: new metrics register as functions here. |
| `portfolio_csv_ingestion.py` | Triggered by EOD | Processes up to 6 input CSVs (3 HL, 3 Barchart). Archives on success. Logs gaps to Scheduler tab. |

### On-demand scripts

| Script | Purpose |
|---|---|
| `sync_engine_tv.py` | Generates TradingView watchlist `.txt` files. Pushes to GDrive. Runs as a persistent systemd service (`universe_sync_tv.service`) on port 8001 — triggered via ICU dashboard Run Now or directly with `--run` flag. |
| `coverage_script.py` | Populates SECTOR_COVERAGE and ALLOC_COVERAGE tabs. Runs automatically at end of weekly sequence via `run_weekly.sh`. |

---

## Process control

Process state is tracked and controlled via the **ICU dashboard at `icu.krhub.uk`**.

| State | Meaning |
|---|---|
| `IDLE` | Script not running — last run complete |
| `RUNNING` | Script currently executing |
| `PAUSED` | ICU gate closed — next scheduled run will exit cleanly with `SKIPPED` |
| `ERROR` | Last run failed — check logs |

### ICU dashboard controls

- **Pause / Resume** — flips the `allowed_to_run` flag in ICU. The next scheduled run checks this on wake and exits cleanly if paused. Any run already in progress finishes normally.
- **Run Now** — triggers immediate execution via the component's trigger endpoint. Currently available for `universe_ss_sync_tv` only (port 8001). EOD, intraday, and weekly Run Now wiring is backlogged.

### How it works

Each script calls `check_gate()` from `icu_client.py` on wake. If ICU returns `allowed_to_run: false`, the script pushes a `SKIPPED` status and exits. If ICU is unreachable, the script runs anyway (fail open — ICU downtime never blocks execution).

Status transitions (RUNNING → IDLE/ERROR) are pushed to ICU on every state change via `push_status()`. These appear live on the dashboard.

### Emergency direct stop (if needed)

```bash
# Kill a running script immediately
kill $(pgrep -f price_action_eod.py)

# Pause via .env fallback (pre-ICU method — still works as a belt-and-braces stop)
sed -i 's/UNIVERSE_PROCESS=Y/UNIVERSE_PROCESS=N/' /opt/dev/universe_SS/.env
```

---

## Crontab

```
TZ=Europe/London
0 7,11,15,19 * * 1-5 /opt/dev/universe_SS/run_intraday.sh
0 23 * * 1-5 /opt/dev/universe_SS/run_eod.sh
30 23 * * 5 /opt/dev/universe_SS/run_weekly.sh
```

All cron jobs run via wrapper `.sh` scripts — never inline commands. Each wrapper handles: env loading → pause check → pidfile → GDrive pull → python execution → GDrive push → log push → Pushover alert on failure.

---

## Folder structure

```
/opt/dev/universe_SS/
  .env                          ← UNIVERSE_PROCESS, PUSHOVER_*, ICU_API_KEY, ICU_INGEST_URL, ICU_CONTROL_URL (never committed)
  .gitignore
  requirements.txt              ← pip freeze of current venv
  venv/                         ← virtual environment (not committed)
  workbook_io.py                ← shared module: workbook glob, write_cell(), alignment preservation
  icu_client.py                 ← ICU integration: push_status, check_gate, resolve_version
  setup_legend_config.py        ← idempotent Legend tab population
  logging_config.py             ← structured logging: RUN_ID, UID placeholder, vocabulary validator
  price_action_intraday.py
  price_action_eod.py
  fetch_engine_weekly.py
  portfolio_csv_ingestion.py
  derive_engine.py
  coverage_script.py            ← SECTOR_COVERAGE + ALLOC_COVERAGE tab writer (runs in weekly)
  sync_engine_tv.py             ← TV watchlist generator + FastAPI trigger server (port 8001)
  universe_sync_tv.service      ← systemd service file (also deployed to /etc/systemd/system/)
  run_intraday.sh
  run_eod.sh
  run_weekly.sh
  Output/
    TV/                         ← generated TradingView .txt files (gitignored)
  Cowork/
    Inputs/
      HL/                       ← account-summary.csv, portfolio-summary.csv, income-transactions.csv
      Barchart/                 ← Custom.csv, Performance.csv, Fundamental.csv
    Archive/
      HL/                       ← processed HL CSVs (_YYYYMMDD_HHMMSS suffix)
      Barchart/                 ← processed Barchart CSVs
      Workbook/                 ← superseded .xlsm versions
      Scripts/                  ← retired scripts
```

---

## TradingView Watchlists

### Running sync_engine_tv

`sync_engine_tv.py` runs as a persistent systemd service (`universe_sync_tv.service`) waiting for ICU Run Now signals. It can also be triggered directly:

```bash
# Trigger via ICU dashboard Run Now button (preferred)
# — open icu.krhub.uk, click Run Now on Universe SS Sync (TV)

# Or run directly on the server
python3 /opt/dev/universe_SS/sync_engine_tv.py --run

# Check service status
sudo systemctl status universe_sync_tv
```

> **Note:** sync_engine_tv reads the **local workbook** on the server — it does not pull from GDrive before running. Always ensure the server workbook is current before generating watchlists. The weekly run (`run_weekly.sh`) refreshes it automatically.

---

### Columns used

| Column | Type | Values | Role |
|---|---|---|---|
| `M_Eliminated` | Fixed | `No Touch` / blank | Kill switch — `No Touch` excludes ticker from all watchlists, checked first |
| `M_Export_TV` | Variable | Any string / `N` | Free-text watchlist name. `N` = exclude. Value becomes the `.txt` filename. |
| `M_Sleeve_Watchlist` | Fixed | `Y` / blank | `Y` = include this ticker in an auto-generated watchlist named after its `M_Sleeve` value |
| `M_Universe_Watchlist` | Fixed | `Y` / blank | `Y` = include in auto-generated watchlist named after its `M_Universe` value |
| `M_Div_Coupon_Class_Watchlist` | Fixed | `Y` / blank | `Y` = include in auto-generated watchlist named after its `M_Div_Coupon_Class` value |
| `M_Related_To` | Variable | Any string | Section label within a watchlist file. Distinct values become `####SECTION` headers. No fixed set — whatever is in the column is used. Never controls inclusion/exclusion. |
| `S_TV_Ticker` | Variable | TradingView ticker string | The value written into the `.txt` file. Falls back to `M_Ticker` if blank. |

### Two independent mechanisms

**1. Free-text routing** (`M_Export_TV`):
- Populate with a literal watchlist name (e.g. `Core`, `Portfolio`, `Growth`)
- Same ticker can appear in multiple watchlists with different names
- Set to `N` to exclude from this mechanism entirely

**2. Boolean-flag auto-derivation** (three columns):
- `M_Sleeve_Watchlist = Y` → generates one `.txt` per distinct `M_Sleeve` value (e.g. `DIV_CORE.txt`, `GROWTH.txt`)
- `M_Universe_Watchlist = Y` → generates one `.txt` per distinct `M_Universe` value
- `M_Div_Coupon_Class_Watchlist = Y` → generates one `.txt` per distinct `M_Div_Coupon_Class` value (e.g. `Aristocrat_King.txt`)

Both mechanisms run on every execution. A ticker can appear in many watchlists simultaneously — this is by design.

### Section headers within each file

Tickers within a watchlist are grouped by their `M_Related_To` value. Example output:

```
####XLK
NASDAQ:AAPL
NASDAQ:MSFT
####XLF
NYSE:JPM
NYSE:BAC
####Other
LSE:HLMA
```

Tickers with no `M_Related_To` value land in `####Other` at the end.

### Output

```
/opt/dev/universe_SS/Output/TV/
  Core.txt
  Portfolio.txt
  DIV_CORE.txt
  Aristocrat_King.txt
  ...
```

Files are overwritten on every run. Upload whichever files you need to TradingView via the "Upload list…" button — daily-use watchlists and weekend-analysis watchlists are separate human choices at upload time.

```bash
# Run any script directly
python3 /opt/dev/universe_SS/derive_engine.py

# Generate TV watchlists (direct)
python3 /opt/dev/universe_SS/sync_engine_tv.py --run

# Run the full weekly sequence
bash /opt/dev/universe_SS/run_weekly.sh

# Check ICU dashboard
open https://icu.krhub.uk

# Tail logs
tail -f /var/log/portfolio/run_weekly.log
tail -f /var/log/portfolio/eod.log
```

---

## Environment variables (`.env`)

```
UNIVERSE_PROCESS=Y
PUSHOVER_TOKEN=<token>
PUSHOVER_USER=<user>
ICU_INGEST_URL=https://icu.krhub.uk/ingest/status
ICU_CONTROL_URL=https://icu.krhub.uk/control
ICU_API_KEY=<key>
```

`icu_client.py` self-loads `.env` at import time via python-dotenv — no reliance on the invoking shell having sourced it.

---

## Logging

All script output goes to `/var/log/portfolio/`. Each run appends to its own `.log` file with a structured final line:

```
--- Started: 2026-07-09 03:35:56 | Ended: 03:36:00 | Duration: 0:00:03 ---
```

Logs are rotated daily, compressed, and kept for 30 days (`/etc/logrotate.d/portfolio`). Post-run logs are also rclone'd to `gdrive:Claude/TradingUniverse/Logs/`.

---

## Workbook versioning

The active workbook is the single highest-versioned `.xlsm` in the working directory (resolved via `sort -V | tail -1`). Old versions are automatically archived to `Cowork/Archive/Workbook/` on each save. GDrive is the canonical master — the server copy is a working clone refreshed at the start of every wrapper run.
