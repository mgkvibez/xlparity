#!/usr/bin/env bash
# One gate, identical locally and in CI: lint, tests, and a README count that cannot drift.
set -euo pipefail
cd "$(dirname "$0")/.."

python -m ruff check .
python -m pytest -q
python tools/check_test_count.py
python tools/scoreboard.py
