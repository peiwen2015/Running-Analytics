#!/bin/zsh
set -e

cd "$(dirname "$0")"

PYTHON_CMD="python3"
APP_PYTHON="$PYTHON_CMD"
APP_PORT="8765"

stop_existing_server() {
  local pids
  pids=$(lsof -tiTCP:"$APP_PORT" -sTCP:LISTEN 2>/dev/null || true)
  if [ -z "$pids" ]; then
    return 0
  fi

  echo "偵測到本機網站 http://127.0.0.1:$APP_PORT 已經在執行，正在關閉舊伺服器..."
  for pid in ${(f)pids}; do
    kill "$pid" 2>/dev/null || true
  done

  for _ in {1..30}; do
    if ! lsof -tiTCP:"$APP_PORT" -sTCP:LISTEN >/dev/null 2>&1; then
      echo "舊伺服器已關閉，準備重新啟動。"
      return 0
    fi
    sleep 0.2
  done

  echo "無法自動關閉舊伺服器。請先關掉舊的轉檔視窗，或在終端機執行："
  echo "lsof -nP -iTCP:$APP_PORT -sTCP:LISTEN"
  exit 1
}

if [ -x ".venv/bin/python" ]; then
  APP_PYTHON=".venv/bin/python"
elif [ -d ".venv/lib/python3.14/site-packages" ]; then
  export PYTHONPATH=".venv/lib/python3.14/site-packages${PYTHONPATH:+:$PYTHONPATH}"
else
  "$PYTHON_CMD" -m venv .venv
  APP_PYTHON=".venv/bin/python"
fi

if ! "$APP_PYTHON" - <<'PY' >/dev/null 2>&1
import garmin_fit_sdk
import openpyxl
PY
then
  "$PYTHON_CMD" -m venv --clear .venv
  APP_PYTHON=".venv/bin/python"
  "$APP_PYTHON" -m pip install -r requirements.txt
fi

stop_existing_server
"$APP_PYTHON" app.py
