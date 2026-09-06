#!/bin/bash
# ==========================================================================
# Install (or remove) the daily scheduled run. macOS only.
# --------------------------------------------------------------------------
# Takes com.tomas.hotelrates.plist, fills in wherever this folder actually
# lives, and hands the result to launchd. Re-run it after moving the folder.
#
#   ./install_schedule.sh            install / reinstall, then load
#   ./install_schedule.sh --status   is it loaded? when did it last run?
#   ./install_schedule.sh --run-now  trigger one run immediately
#   ./install_schedule.sh --remove   unload and delete it
#
# To change the time of day, edit StartCalendarInterval in the plist and run
# this again. To change what the run does, edit run_daily.sh - no reinstall
# needed, launchd re-reads the script every time.
# ==========================================================================

set -uo pipefail

LABEL="com.tomas.hotelrates"
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEMPLATE="$PROJECT_DIR/$LABEL.plist"
AGENT_DIR="$HOME/Library/LaunchAgents"
INSTALLED="$AGENT_DIR/$LABEL.plist"

if [[ "$(uname)" != "Darwin" ]]; then
    echo "This installer is macOS-only (launchd). On Linux use cron; on"
    echo "Windows use Task Scheduler. See GUIDE.md."
    exit 1
fi

# `launchctl bootout` is the modern verb; `unload` still works everywhere and
# is quiet about an agent that is not loaded, so both are tried.
unload_agent() {
    launchctl unload "$INSTALLED" 2>/dev/null
    launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null
    return 0
}

case "${1:-}" in
    --remove)
        unload_agent
        rm -f "$INSTALLED"
        echo "Removed. The daily run is no longer scheduled."
        echo "Your data in output/ is untouched."
        exit 0
        ;;

    --status)
        if launchctl list | grep -q "$LABEL"; then
            echo "LOADED:"
            launchctl list | grep "$LABEL" | sed 's/^/  /'
            echo "  (columns: PID  last-exit-status  label)"
            echo "  last-exit-status 0 means the last run succeeded."
        else
            echo "NOT loaded. Run ./install_schedule.sh to schedule it."
        fi
        echo
        if [[ -f "$INSTALLED" ]]; then
            echo "Installed plist: $INSTALLED"
        else
            echo "Installed plist: none ($INSTALLED does not exist)"
        fi
        LATEST_LOG="$(ls -t "$PROJECT_DIR/output/logs/"*.log 2>/dev/null | head -1)"
        if [[ -n "$LATEST_LOG" ]]; then
            echo "Most recent log:  $LATEST_LOG"
            tail -5 "$LATEST_LOG" | sed 's/^/  /'
        else
            echo "No run logs yet in output/logs/."
        fi
        exit 0
        ;;

    --run-now)
        if ! launchctl list | grep -q "$LABEL"; then
            echo "Not loaded - install it first with ./install_schedule.sh"
            exit 1
        fi
        launchctl start "$LABEL"
        echo "Started. It runs in the background; watch it with:"
        echo "  tail -f $PROJECT_DIR/output/logs/$(date +%Y-%m-%d).log"
        exit 0
        ;;

    "") ;;  # no argument: install

    *)
        echo "Unknown option: $1"
        echo "Use: (no args) | --status | --run-now | --remove"
        exit 2
        ;;
esac

# ---- install ------------------------------------------------------------

if [[ ! -f "$TEMPLATE" ]]; then
    echo "FATAL: missing $TEMPLATE"
    exit 1
fi

if [[ ! -x "$PROJECT_DIR/.venv/bin/python" ]]; then
    echo "FATAL: no virtual environment found."
    echo "       Run ./setup.sh first, then this script."
    exit 1
fi

chmod +x "$PROJECT_DIR/run_daily.sh"
mkdir -p "$AGENT_DIR" "$PROJECT_DIR/output/logs"

# Any previous copy must go before the new one is written, or launchd keeps
# running the old paths until the next login.
unload_agent

# `|` as the sed delimiter, since the path contains `/`.
sed "s|__PROJECT_DIR__|$PROJECT_DIR|g" "$TEMPLATE" > "$INSTALLED"

if grep -q "__PROJECT_DIR__" "$INSTALLED"; then
    echo "FATAL: placeholder substitution failed - is the template intact?"
    rm -f "$INSTALLED"
    exit 1
fi

if ! plutil -lint "$INSTALLED" >/dev/null; then
    echo "FATAL: the generated plist is not valid XML."
    rm -f "$INSTALLED"
    exit 1
fi

launchctl load "$INSTALLED"

if launchctl list | grep -q "$LABEL"; then
    PB=/usr/libexec/PlistBuddy
    HH="$($PB -c "Print :StartCalendarInterval:Hour" "$INSTALLED" 2>/dev/null)"
    MM="$($PB -c "Print :StartCalendarInterval:Minute" "$INSTALLED" 2>/dev/null)"
    if [[ -n "$HH" && -n "$MM" ]]; then
        WHEN="$(printf '%02d:%02d' "$HH" "$MM")"
    else
        WHEN="the time set in the plist"
    fi
    echo "Scheduled. It will run daily at $WHEN (24h clock)."
    echo "  project:  $PROJECT_DIR"
    echo "  logs:     $PROJECT_DIR/output/logs/"
    echo "  check:    ./install_schedule.sh --status"
    echo "  test now: ./install_schedule.sh --run-now"
    echo
    echo "Note: the Mac must be awake at that time. launchd will catch up on"
    echo "the next wake if it was asleep, but not if it was shut down."
else
    echo "WARNING: launchd did not report the agent as loaded."
    echo "Check: launchctl list | grep $LABEL"
    exit 1
fi
