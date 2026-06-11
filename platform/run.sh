#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
export DARB_SECRET="${DARB_SECRET:-darb-secret-change-me}"
exec uvicorn app.main:app --host 0.0.0.0 --port "${PORT:-8000}"
