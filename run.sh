#!/usr/bin/env bash
# SLT Project Management — start script (macOS / Linux)
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
  ./.venv/bin/pip install --upgrade pip >/dev/null
  ./.venv/bin/pip install -r requirements.txt
fi
( sleep 2; (open "http://127.0.0.1:${SLT_PM_PORT:-5000}" 2>/dev/null || xdg-open "http://127.0.0.1:${SLT_PM_PORT:-5000}" 2>/dev/null) ) &
exec ./.venv/bin/python app.py
