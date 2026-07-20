#!/usr/bin/env python3
"""
preflight_inputs.py
-------------------
Checks Inputs/BC/ and Inputs/HL/ for CSVs.
If either folder is empty, pulls from GDrive via rclone.
Intended to run before price_action_eod.py.

GDrive sources:
  gdrive:Claude/TradingUniverse/input/BC  ->  Inputs/BC/
  gdrive:Claude/TradingUniverse/input/HL  ->  Inputs/HL/

portfolio_csv_ingestion.py handles individual file presence
gracefully, so preflight only needs to confirm files exist —
not which specific ones.
"""

import os
import subprocess
import logging

# ── Config ────────────────────────────────────────────────────────────────────
BASE_INPUTS  = "/opt/dev/universe_SS/Inputs"
INPUTS_BC    = f"{BASE_INPUTS}/BC"
INPUTS_HL    = f"{BASE_INPUTS}/HL"

GDRIVE_BC    = "gdrive:Claude/TradingUniverse/input/BC"
GDRIVE_HL    = "gdrive:Claude/TradingUniverse/input/HL"

LOG_PATH     = "/var/log/universe_SS/preflight_inputs.log"

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
def has_csv(directory: str) -> bool:
    """Return True if directory contains at least one CSV file."""
    try:
        return any(f.lower().endswith(".csv") for f in os.listdir(directory))
    except FileNotFoundError:
        return False


def pull_from_gdrive(gdrive_path: str, local_dir: str, label: str) -> bool:
    os.makedirs(local_dir, exist_ok=True)
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

    errors = []

    for label, local_dir, gdrive_path in [
        ("BC", INPUTS_BC, GDRIVE_BC),
        ("HL", INPUTS_HL, GDRIVE_HL),
    ]:
        os.makedirs(local_dir, exist_ok=True)

        if has_csv(local_dir):
            files = [f for f in os.listdir(local_dir) if f.lower().endswith(".csv")]
            log.info(f"{label}: found locally — {', '.join(files)}")
        else:
            log.warning(f"{label}: no CSVs found locally — attempting GDrive pull")
            ok = pull_from_gdrive(gdrive_path, local_dir, label)
            if ok:
                if has_csv(local_dir):
                    files = [f for f in os.listdir(local_dir) if f.lower().endswith(".csv")]
                    log.info(f"{label}: confirmed after pull — {', '.join(files)}")
                else:
                    log.error(f"{label}: pull reported success but no CSVs arrived")
                    errors.append(label)
            else:
                errors.append(label)

    if errors:
        log.error(f"=== Preflight FAILED — no CSVs in: {', '.join(errors)} ===")
        return 1

    log.info("=== Preflight passed — all input folders populated ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
