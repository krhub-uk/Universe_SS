#!/usr/bin/env python3
"""
preflight_inputs.py
-------------------
Checks /opt/dev/universe_ss/Inputs/ for BC and HL CSVs.
If either is missing, pulls from GDrive via rclone.
Intended to run before fetch_engine_weekly.py / price_action_eod.py.
"""

import os
import subprocess
import logging
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
INPUTS_DIR = "/opt/dev/universe_ss/Inputs"

GDRIVE_BC = "gdrive:Claude/TradingUniverse/input/BC"
GDRIVE_HL = "gdrive:Claude/TradingUniverse/input/HL"

LOG_PATH = "/var/log/universe_ss/preflight_inputs.log"

# Expected filename fragments — partial match so date-stamped filenames still match
BC_PATTERN = "bc"
HL_PATTERN = "hl"

# ── Logging ───────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [PREFLIGHT] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ── Helpers ───────────────────────────────────────────────────────────────────
def find_csv(directory: str, pattern: str) -> str | None:
    """Return first CSV in directory whose lowercase name contains pattern."""
    try:
        for f in os.listdir(directory):
            if f.lower().endswith(".csv") and pattern in f.lower():
                return f
    except FileNotFoundError:
        pass
    return None


def pull_from_gdrive(gdrive_path: str, local_dir: str, label: str) -> bool:
    """
    Pull CSV(s) from GDrive folder into local_dir using rclone copy.
    Returns True on success.
    """
    log.info(f"{label}: pulling from {gdrive_path} → {local_dir}")
    result = subprocess.run(
        ["rclone", "copy", gdrive_path, local_dir, "--include", "*.csv"],
        capture_output=True,
        text=True,
    )
    if result.returncode == 0:
        log.info(f"{label}: pull succeeded")
        return True
    else:
        log.error(f"{label}: rclone failed — {result.stderr.strip()}")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> int:
    log.info("=== Preflight check started ===")
    os.makedirs(INPUTS_DIR, exist_ok=True)

    errors = []

    for label, pattern, gdrive_path in [
        ("BC", BC_PATTERN, GDRIVE_BC),
        ("HL", HL_PATTERN, GDRIVE_HL),
    ]:
        found = find_csv(INPUTS_DIR, pattern)
        if found:
            log.info(f"{label}: found locally → {found}")
        else:
            log.warning(f"{label}: not found locally — attempting GDrive pull")
            ok = pull_from_gdrive(gdrive_path, INPUTS_DIR, label)
            if ok:
                # Confirm file arrived
                found = find_csv(INPUTS_DIR, pattern)
                if found:
                    log.info(f"{label}: confirmed after pull → {found}")
                else:
                    log.error(f"{label}: pull reported success but file still missing")
                    errors.append(label)
            else:
                errors.append(label)

    if errors:
        log.error(f"=== Preflight FAILED — missing: {', '.join(errors)} ===")
        return 1

    log.info("=== Preflight passed — all inputs present ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
