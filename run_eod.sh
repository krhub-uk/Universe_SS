#!/bin/bash
# run_eod.sh — Portfolio Automation Engine wrapper (V4.1)
# Steps per S3-OPS-01 spec, S3-OPS-06 pre-run pull inserted at step 6

set -uo pipefail

BASE_DIR="/opt/dev/universe_SS"
LOG_FILE="/var/log/portfolio/eod.log"
PID_FILE="$BASE_DIR/.pid"

# 1. Source .env
if [ -f "$BASE_DIR/.env" ]; then
    set -a
    source "$BASE_DIR/.env"
    set +a
else
    echo "$(date '+%Y-%m-%d %H:%M:%S') FATAL: .env not found at $BASE_DIR/.env" >> "$LOG_FILE"
    exit 1
fi

# 2. UNIVERSE_PROCESS=N check — clean exit, not a failure
if [ "${UNIVERSE_PROCESS:-Y}" = "N" ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') SKIPPED — UNIVERSE_PROCESS=N (paused)" >> "$LOG_FILE"
    exit 0
fi

# 3. Start timestamp
START_TS=$(date '+%Y-%m-%d %H:%M:%S')
START_EPOCH=$(date +%s)

# 4. pidfile write
echo $$ > "$PID_FILE"

# 5. SIGTERM trap — log HALT, remove pidfile, exit 1
trap 'echo "$(date "+%Y-%m-%d %H:%M:%S") HALT — SIGTERM received" >> "$LOG_FILE"; rm -f "$PID_FILE"; exit 1' SIGTERM

cd "$BASE_DIR" || { echo "$(date '+%Y-%m-%d %H:%M:%S') FATAL: cannot cd to $BASE_DIR" >> "$LOG_FILE"; rm -f "$PID_FILE"; exit 1; }

# 6. Pre-run rclone pull — SINGLE highest-versioned .xlsm from GDrive root only (S3-OPS-06)
# GDrive root accumulates historical versions (deliberate, no-deletion policy) —
# must select the canonical highest version by filename sort, not glob everything.
# Only pull if remote is genuinely newer than local, to avoid overwriting unsynced local edits
# and to preserve single-file discipline in both directions.
LATEST_REMOTE=$(rclone lsf gdrive:Claude/TradingUniverse/ --max-depth 1 --files-only --include "00_Portfolio_Engine_Universe_v*.xlsm" 2>>"$LOG_FILE" | sort -V | tail -1)
LOCAL_FILE=$(ls "$BASE_DIR"/00_Portfolio_Engine_Universe_v*.xlsm 2>/dev/null | sort -V | tail -1)
LOCAL_NAME=$(basename "${LOCAL_FILE:-}" 2>/dev/null)

if [ -z "$LATEST_REMOTE" ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') WARNING: no remote .xlsm found matching pattern — skipping pull" >> "$LOG_FILE"
elif [ "$LATEST_REMOTE" = "$LOCAL_NAME" ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') Pre-run pull: local already current ($LOCAL_NAME) — skipped" >> "$LOG_FILE"
elif [ "$(printf '%s\n%s\n' "$LATEST_REMOTE" "$LOCAL_NAME" | sort -V | tail -1)" = "$LATEST_REMOTE" ]; then
    # Remote is newer — clear stale local file(s) to Archive, then pull
    if [ -n "$LOCAL_FILE" ]; then
        mkdir -p "$BASE_DIR/Archive/Workbook"
        mv "$LOCAL_FILE" "$BASE_DIR/Archive/Workbook/" >> "$LOG_FILE" 2>&1
    fi
    rclone copy "gdrive:Claude/TradingUniverse/$LATEST_REMOTE" "$BASE_DIR/" >> "$LOG_FILE" 2>&1
    echo "$(date '+%Y-%m-%d %H:%M:%S') Pre-run pull: $LATEST_REMOTE (remote newer than local $LOCAL_NAME)" >> "$LOG_FILE"
else
    echo "$(date '+%Y-%m-%d %H:%M:%S') Pre-run pull: local ($LOCAL_NAME) newer than remote ($LATEST_REMOTE) — skipped, keeping local" >> "$LOG_FILE"
fi

# 7. venv activation + python execution
source "$BASE_DIR/venv/bin/activate"
python3 "$BASE_DIR/price_action_eod.py" >> "$LOG_FILE" 2>&1
SCRIPT_EXIT=$?

# 8. End timestamp + elapsed
END_TS=$(date '+%Y-%m-%d %H:%M:%S')
END_EPOCH=$(date +%s)
ELAPSED=$((END_EPOCH - START_EPOCH))
ELAPSED_FMT=$(printf '%02d:%02d:%02d' $((ELAPSED/3600)) $((ELAPSED%3600/60)) $((ELAPSED%60)))

# 9. Final structured log line
echo "--- Started: $START_TS | Ended: $END_TS | Duration: $ELAPSED_FMT ---" >> "$LOG_FILE"

# 10. rclone: xlsm + Archive/ -> GDrive
# Self-healing: if more than one local .xlsm exists, archive all but the highest-versioned
# before pushing, so a stray file never breaks the single-arg rclone copy call.
mapfile -t ALL_LOCAL < <(ls "$BASE_DIR"/00_Portfolio_Engine_Universe_v*.xlsm 2>/dev/null | sort -V)
if [ "${#ALL_LOCAL[@]}" -gt 1 ]; then
    mkdir -p "$BASE_DIR/Archive/Workbook"
    for f in "${ALL_LOCAL[@]:0:$((${#ALL_LOCAL[@]}-1))}"; do
        echo "$(date '+%Y-%m-%d %H:%M:%S') WARNING: stray local .xlsm archived pre-push: $(basename "$f")" >> "$LOG_FILE"
        mv "$f" "$BASE_DIR/Archive/Workbook/" >> "$LOG_FILE" 2>&1
    done
fi
LATEST_LOCAL=$(ls "$BASE_DIR"/00_Portfolio_Engine_Universe_v*.xlsm 2>/dev/null | sort -V | tail -1)
if [ -n "$LATEST_LOCAL" ]; then
    rclone copy "$LATEST_LOCAL" gdrive:Claude/TradingUniverse/ >> "$LOG_FILE" 2>&1
else
    echo "$(date '+%Y-%m-%d %H:%M:%S') WARNING: no local .xlsm found — push skipped" >> "$LOG_FILE"
fi
rclone copy "$BASE_DIR/Archive/" gdrive:Claude/TradingUniverse/Archive/ >> "$LOG_FILE" 2>&1

# 11. rclone: log file -> GDrive
rclone copy "$LOG_FILE" gdrive:Claude/TradingUniverse/Logs/ >> /dev/null 2>&1

# 12. Pushover failure alert on non-zero exit
if [ "$SCRIPT_EXIT" -ne 0 ]; then
    curl -s --form-string "token=${PUSHOVER_TOKEN}" \
         --form-string "user=${PUSHOVER_USER}" \
         --form-string "message=EOD script failed $(date '+%Y-%m-%d %H:%M:%S')" \
         https://api.pushover.net/1/messages.json >> /dev/null 2>&1
fi

# 13. Remove pidfile
rm -f "$PID_FILE"

exit "$SCRIPT_EXIT"
