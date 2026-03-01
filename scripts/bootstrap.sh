#!/usr/bin/env bash
set -euo pipefail

# Choose python3.11 if available, otherwise fallback to python3
PYTHON_CMD=python3.11
if ! command -v "$PYTHON_CMD" >/dev/null 2>&1; then
  PYTHON_CMD=python3
fi

echo "Using $PYTHON_CMD to create virtualenv"
$PYTHON_CMD -m venv .venv
.venv/bin/python -m pip install --upgrade pip setuptools wheel
.venv/bin/pip install -r requirements.txt

echo "Done. Activate with: source .venv/bin/activate"