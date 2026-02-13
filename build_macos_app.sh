#!/usr/bin/env bash
set -euo pipefail

PYTHON_BIN="${PYTHON_BIN:-python3}"
APP_NAME="SpendTagger"

"$PYTHON_BIN" -m pip install -r requirements_macos_app.txt pyinstaller

"$PYTHON_BIN" -m PyInstaller \
  --noconfirm \
  --windowed \
  --name "$APP_NAME" \
  spending_tagger_app.py

echo "Build done: dist/${APP_NAME}.app"
