#!/usr/bin/env bash
# Isolated development-only browser tools; the application has no runtime dependencies.
set -euo pipefail
cd "$(dirname "$0")/.."
TOOLS="${RUNNER_TEMP:-/tmp}/gridline-browser"
python3 -m venv "$TOOLS"
"$TOOLS/bin/python" -m pip install playwright==1.57.0
"$TOOLS/bin/python" -m playwright install --with-deps chromium
CHROMIUM="$("$TOOLS/bin/python" -c 'from playwright.sync_api import sync_playwright; p=sync_playwright().start(); print(p.chromium.executable_path); p.stop()')"
"$TOOLS/bin/python" tests/browser-smoke.py --chromium "$CHROMIUM"
