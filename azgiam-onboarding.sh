#!/usr/bin/env bash
# Simple wrapper so operators don't have to remember bulkprovision.cli's
# full flag set. Auto-installs the one Python dependency, and forwards
# lrp's own install/login prompts straight to your terminal.
#
# There is no separate catalogue file - `lrp search-products` itself
# validates each requested group; a group it can't find becomes an error
# row in the report without blocking any of that recipient's other groups.
#
# Usage:
#   ./azgiam-onboarding.sh                Interactive menu (batch run, check-auth,
#                                          login, find-user, search-products,
#                                          list-orders)
#   ./azgiam-onboarding.sh <input.xlsx>   Run one batch directly, no menu
set -uo pipefail

if [ $# -gt 1 ]; then
    echo "Usage: $0                (interactive menu)" >&2
    echo "   or: $0 <input.xlsx>   (run one batch directly)" >&2
    exit 64
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# A venv, not system pip: many distros (Debian 12+/Ubuntu 23.10+ and others
# following PEP 668) refuse `pip install` outright ("externally-managed-
# environment") to protect the OS's own Python. That failure isn't fatal
# under `set -uo pipefail` (no `-e`), so the script used to plow ahead into
# a confusing `ModuleNotFoundError` instead of stopping at the real cause.
VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON="$VENV_DIR/bin/python3"

if [ ! -x "$PYTHON" ]; then
    echo "Setting up virtual environment ($VENV_DIR)..." >&2
    if ! python3 -m venv "$VENV_DIR"; then
        echo "Failed to create a virtual environment. On Debian/Ubuntu, try:" >&2
        echo "  sudo apt install python3-venv" >&2
        exit 1
    fi
fi

if ! "$PYTHON" -c "import openpyxl" 2>/dev/null; then
    echo "Installing dependencies (openpyxl)..." >&2
    if ! "$VENV_DIR/bin/pip" install -r requirements.txt; then
        echo "Dependency install failed - see the error above." >&2
        exit 1
    fi
fi

if [ $# -eq 0 ]; then
    "$PYTHON" -m bulkprovision.cli
    exit $?
fi

INPUT="$1"
RUN_ID="$(date +%Y%m%d-%H%M%S)"
REPORT="report-${RUN_ID}.csv"
RETRY="retry-${RUN_ID}.json"

"$PYTHON" -m bulkprovision.cli \
    --input "$INPUT" \
    --report "$REPORT" \
    --retry-out "$RETRY" \
    --run-id "$RUN_ID"
EXIT_CODE=$?

echo ""
echo "Report:      $REPORT"
echo "Retry store: $RETRY"
exit "$EXIT_CODE"
