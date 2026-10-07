#!/bin/bash
# Chạy tool tại http://localhost:8787
cd "$(dirname "$0")"
export BS_REDIRECT_URI="http://localhost:8787/oauth/callback"
exec ./.venv/bin/python -m uvicorn app.main:app --host 127.0.0.1 --port 8787
