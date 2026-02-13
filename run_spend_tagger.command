#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [ -x "$HOME/anaconda3/envs/tmp/bin/python" ]; then
  PYTHON_BIN="$HOME/anaconda3/envs/tmp/bin/python"
elif [ -x "$HOME/miniconda3/envs/tmp/bin/python" ]; then
  PYTHON_BIN="$HOME/miniconda3/envs/tmp/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
else
  osascript -e 'display alert "Python not found" message "python3 또는 python 실행 파일을 찾지 못했습니다." as critical'
  exit 1
fi

"$PYTHON_BIN" spending_tagger_app.py
