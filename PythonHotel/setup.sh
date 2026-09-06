#!/bin/bash
# ==========================================================================
# One-time setup. Run this once after downloading or moving the folder.
# --------------------------------------------------------------------------
#     ./setup.sh
#
# Creates .venv/ next to this file and installs the packages in
# requirements.txt into it. Safe to run again at any time - it will rebuild
# the environment from scratch, which is also the fix if the venv ever breaks
# (typically after moving the folder, or after a macOS/Homebrew Python
# upgrade).
#
# It does NOT install Python itself. If Python 3.10+ is missing it says so
# and stops, rather than guessing at your system - see GUIDE.md.
# ==========================================================================

set -uo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$PROJECT_DIR/.venv"
MIN_MINOR=10   # we need Python 3.10 or newer

echo "Project folder: $PROJECT_DIR"
echo

# ---- 1. find a usable Python --------------------------------------------
# The venv records an absolute path to whichever interpreter built it, so a
# Homebrew Python that may be upgraded out from under us is still preferable
# to no Python at all. Newest first.
PYBIN=""
for candidate in python3.13 python3.12 python3.11 python3.10 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
        minor="$("$candidate" -c 'import sys; print(sys.version_info[1])' 2>/dev/null)"
        if [[ -n "$minor" && "$minor" -ge "$MIN_MINOR" ]]; then
            PYBIN="$(command -v "$candidate")"
            break
        fi
    fi
done

if [[ -z "$PYBIN" ]]; then
    echo "FATAL: no Python 3.$MIN_MINOR or newer found."
    echo
    echo "macOS ships an old Python that cannot run this. Install a current one:"
    echo "  1. install Homebrew (once):"
    echo '     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
    echo "  2. install Python:"
    echo "     brew install python@3.13"
    echo "  3. run ./setup.sh again"
    echo
    echo "Or download the installer from https://www.python.org/downloads/"
    exit 1
fi

echo "Using Python: $PYBIN ($("$PYBIN" --version 2>&1))"

# ---- 2. (re)create the virtual environment -------------------------------
# Always from scratch: a venv that was copied from another folder or another
# Mac looks present but is broken in ways that produce confusing errors much
# later, so replacing it is cheaper than testing it.
if [[ -d "$VENV" ]]; then
    echo "Removing the existing .venv (it is rebuilt, not reused)..."
    rm -rf "$VENV"
fi

echo "Creating .venv ..."
if ! "$PYBIN" -m venv "$VENV"; then
    echo "FATAL: could not create the virtual environment."
    echo "       If this says 'No module named venv', install it:"
    echo "       brew install python@3.13   (or reinstall Python)"
    exit 1
fi

# ---- 3. install the packages --------------------------------------------
echo "Installing packages from requirements.txt ..."
"$VENV/bin/python" -m pip install --upgrade pip --quiet
if ! "$VENV/bin/python" -m pip install -r "$PROJECT_DIR/requirements.txt"; then
    echo
    echo "FATAL: package installation failed."
    echo "       Most often this is no internet connection, or a corporate"
    echo "       proxy blocking pypi.org. Check both and run ./setup.sh again."
    exit 1
fi

# ---- 4. make the shell scripts runnable ----------------------------------
chmod +x "$PROJECT_DIR"/*.sh 2>/dev/null

# ---- 5. prove it works ---------------------------------------------------
echo
echo "Verifying ..."
if ! "$VENV/bin/python" -c "import pandas, openpyxl, requests; print('  packages OK')"; then
    echo "FATAL: packages installed but cannot be imported."
    exit 1
fi

mkdir -p "$PROJECT_DIR/output/logs"

if [[ ! -f "$PROJECT_DIR/hotels.json" ]]; then
    echo "  WARNING: hotels.json is missing - the scraper has no properties"
    echo "           to visit. See GUIDE.md."
fi

cat <<EOF

Setup complete.

Next steps
  1. collect today's rates (this takes a while - see DAYS_AHEAD):
       ./.venv/bin/python hotel_rates.py
  2. build the Excel report from everything collected so far:
       ./.venv/bin/python build_report.py
     -> $PROJECT_DIR/output/report_<date>.xlsx
  3. optional, to have both run every morning by themselves:
       ./install_schedule.sh

To try a quick 3-day scrape first, so you are not waiting on a full run:
  ./.venv/bin/python hotel_rates.py --days 3

Everything else is explained in GUIDE.md.
EOF
