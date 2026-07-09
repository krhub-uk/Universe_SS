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

### On-demand scripts (manual trigger)

| Script | Purpose |
|---|---|
| `sync_engine_tv.py` | Generates TradingView watchlist `.txt` files. Pushes to GDrive. Run when needed — not in cron. |

---

## Process control

The engine uses a simple three-state system readable from the filesystem at any time.

| State | Indicator |
|---|---|
| `IDLE` | No `.pid` file present |
| `RUNNING` | `.pid` file present |
| `PAUSED` | `UNIVERSE_PROCESS=N` in `.env` |

### Aliases (add to `~/.bashrc`)

```bash
alias universe-pause="sed -i 's/UNIVERSE_PROCESS=Y/UNIVERSE_PROCESS=N/' /opt/dev/universe_SS/.env && echo 'Universe paused'"
alias universe-resume="sed -i 's/UNIVERSE_PROCESS=N/UNIVERSE_PROCESS=Y/' /opt/dev/universe_SS/.env && echo 'Universe resumed'"
alias universe-status="[ -f /opt/dev/universe_SS/.pid ] && echo RUNNING || (grep -q 'UNIVERSE_PROCESS=N' /opt/dev/universe_SS/.env && echo PAUSED || echo IDLE)"
```

**Pausing** sets `UNIVERSE_PROCESS=N` in `.env`. The next scheduled run reads this on startup and exits cleanly — no cron changes, no mid-run kills. Any run already in progress finishes normally.

**Manual HALT** (kills a running script immediately):
```bash
kill $(cat /opt/dev/universe_SS/.pid)
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
  .env                          ← UNIVERSE_PROCESS, PUSHOVER_*, API keys (never committed)
  .pid                          ← present only when a script is RUNNING
  .gitignore
  workbook_io.py                ← shared module: workbook glob, write_cell(), versioning
  setup_legend_config.py        ← idempotent Legend tab population
  price_action_intraday.py
  price_action_eod.py
  fetch_engine_weekly.py
  portfolio_csv_ingestion.py
  derive_engine.py
  sync_engine_tv.py
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

## Running manually

```bash
# Run any script directly
python3 /opt/dev/universe_SS/derive_engine.py

# Generate TV watchlists
python3 /opt/dev/universe_SS/sync_engine_tv.py

# Run the full weekly sequence
bash /opt/dev/universe_SS/run_weekly.sh

# Check what's happening
universe-status
tail -f /var/log/portfolio/run_weekly.log
```

---

## Environment variables (`.env`)

```
UNIVERSE_PROCESS=Y
PUSHOVER_TOKEN=<token>
PUSHOVER_USER=<user>
```

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
