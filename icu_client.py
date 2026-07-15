"""
icu_client.py -- shared ICU integration module for Universe_SS scripts.
Ref: ICU_to_Universe_Handover_v1.0 - ICU_Contract_Spec_v1.2_BasketA

Standalone module. Every wired script imports from it directly.
Nothing here is folded into logging_config.py.

Locked principles:
  - push_status() is always best-effort -- never raises, never blocks.
  - check_gate() fails open -- if ICU is unreachable, returns True.
  - VERSION is never hardcoded -- resolve_version() extracts it from the
    workbook filename resolved via the existing glob pattern.
"""

import httpx
import os
from datetime import datetime, timezone
from pathlib import Path

# Load .env directly -- don't rely on the invoking shell script having
# sourced it. Scripts wired to icu_client (price_action_eod.py etc.) are
# normally launched via run_*.sh wrappers that `source .env` first, but
# sync_engine_tv.py has no such wrapper and is called as a bare
# `python3 sync_engine_tv.py`, so this module must be self-sufficient.
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass  # falls back to whatever is already in the process environment

ICU_INGEST_URL = os.getenv("ICU_INGEST_URL", "https://icu.krhub.uk/ingest/status")
ICU_CONTROL_URL = os.getenv("ICU_CONTROL_URL", "https://icu.krhub.uk/control")

# POST /ingest/status and GET /control/{component_id} require this header.
# PATCH /control/{component_id} stays OAuth-only (dashboard human action) --
# not something this client calls.
_ICU_HEADERS = {"X-ICU-Key": os.getenv("ICU_API_KEY", "")}


def push_status(
    component_id: str,
    status: str,
    version: str,
    last_run_utc: str | None = None,
    last_run_result: str | None = None,
    trigger: str | None = None,
    message: str | None = None,
    metrics: dict | None = None,
    health: dict | None = None,
) -> None:
    """Push a status payload to ICU. Best-effort -- never raises."""
    payload = {
        "schema_version": "1.0",
        "component_id": component_id,
        "status": status,
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "version": version,
        "last_run_utc": last_run_utc,
        "last_run_result": last_run_result,
    }
    if trigger:  payload["trigger"] = trigger
    if message:  payload["message"] = message
    if metrics:  payload["metrics"] = metrics
    if health:   payload["health"] = health

    try:
        httpx.post(ICU_INGEST_URL, json=payload, headers=_ICU_HEADERS, timeout=5)
    except Exception:
        pass


def check_gate(component_id: str) -> bool:
    """
    Returns True if ICU allows this component to run.
    Defaults to True on any error -- fail open, ICU downtime never blocks scripts.
    """
    try:
        r = httpx.get(f"{ICU_CONTROL_URL}/{component_id}", headers=_ICU_HEADERS, timeout=3)
        return r.json().get("allowed_to_run", True)
    except Exception:
        return True


def resolve_version(workbook_path: str) -> str:
    """
    Extract version string from workbook filename.
    e.g. '00_Portfolio_Engine_Universe_v1_4_142.xlsm' -> '1.4.142'
    Returns 'unknown' if pattern does not match.
    """
    import re
    match = re.search(r'v(\d+)_(\d+)_(\d+)', str(workbook_path))
    if match:
        return f"{match.group(1)}.{match.group(2)}.{match.group(3)}"
    return "unknown"
