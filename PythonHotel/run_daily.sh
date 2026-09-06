#!/bin/bash
# ==========================================================================
# Daily scrape wrapper for launchd/cron.
# --------------------------------------------------------------------------
# WHAT THIS FILE DOES
#   1. scrape   hotel_rates.py   -> output/hotel_rates_<stamp>.xlsx
#   2. present  build_report.py  -> output/report_<date>.xlsx
#   3. prune    delete logs older than LOG_KEEP_DAYS
#
# WHAT THIS FILE DELIBERATELY DOES NOT DO
#   It passes NO --days and NO --nights. Those are settings, and a setting
#   belongs in exactly one place:
#       how many days ahead to scrape   -> DAYS_AHEAD  in hotel_rates.py
#       how many days wide the report   -> REPORT_DAYS in build_report.py
#       stay length per hotel           -> "nights"    in hotels.json
#   Passing them here as well would silently beat those values, so changing
#   hotel_rates.py would appear to do nothing on the scheduled run. If you
#   ever need a one-off different window, run the script by hand with the
#   flag - don't add it back here.
#
# EDITABLE SETTINGS IN THIS FILE
#   LOG_KEEP_DAYS   how many days of logs to keep (below)
#   the schedule itself lives in com.tomas.hotelrates.plist, not here;
#   install it with ./install_schedule.sh
#
# OTHER COMMANDS (all use .venv/bin/python)
#   price_report.py <report>   ask a single question on the terminal
#       reports: summary | evolution | by-checkin | availability |
#                changes | sales
#   full command reference: see GUIDE.md, or the header of any .py file
# ==========================================================================
#
# Scheduled jobs get a minimal environment and no terminal, so this resolves
# everything relative to this file's own location and sends all output to a
# dated log. That means the project folder can be moved or copied anywhere
# without editing this script.

set -uo pipefail

LOG_KEEP_DAYS=90

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="$PROJECT_DIR/.venv/bin/python"
LOG_DIR="$PROJECT_DIR/output/logs"
LOG_FILE="$LOG_DIR/$(date +%Y-%m-%d).log"

mkdir -p "$LOG_DIR"

{
    echo "=============================================================="
    echo "run started $(date '+%Y-%m-%d %H:%M:%S %Z')"
    echo "project dir: $PROJECT_DIR"

    if [[ ! -x "$PYTHON" ]]; then
        echo "FATAL: no interpreter at $PYTHON"
        echo "       run ./setup.sh in $PROJECT_DIR to create the venv"
        exit 127
    fi

    # No flags on purpose - see the header. hotel_rates.py uses its own
    # DAYS_AHEAD, and each hotel's stay length comes from hotels.json.
    "$PYTHON" "$PROJECT_DIR/hotel_rates.py"
    status=$?

    if [[ $status -eq 0 ]]; then
        echo "run finished OK $(date '+%H:%M:%S')"
        # Refresh the presentation report from the new data. A failure here
        # must not mask a successful scrape, so its status is reported
        # separately and does not overwrite $status.
        if "$PYTHON" "$PROJECT_DIR/build_report.py"; then
            echo "report rebuilt $(date '+%H:%M:%S')"
        else
            echo "WARNING: scrape succeeded but report build failed"
        fi
    else
        # Most likely a credential problem: see the README.
        echo "run FAILED (exit $status) $(date '+%H:%M:%S')"
    fi

    # Keep LOG_KEEP_DAYS days of logs so this never grows without bound.
    find "$LOG_DIR" -name '*.log' -type f -mtime +"$LOG_KEEP_DAYS" -delete 2>/dev/null

    exit $status
} >> "$LOG_FILE" 2>&1
