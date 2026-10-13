#!/bin/bash
# Starts both conversion services and stops the container when either exits
# (run it with a restart policy). Ports: PDF_APP_PORT (default 5000) and
# OFFICE_APP_PORT (default 8081), bound to BIND_HOST (default 0.0.0.0).
set -euo pipefail

host="${BIND_HOST:-0.0.0.0}"

cd /app/pdf_to_doc_converter
gunicorn -c gunicorn_config.py --bind "${host}:${PDF_APP_PORT:-5000}" api:app &
pdf_pid=$!

cd /app/ppt_to_pdf_converter
# --workers 1: the concurrency limits (LO_MAX_CONCURRENCY etc.) are per process.
uvicorn main:app --host "$host" --port "${OFFICE_APP_PORT:-8081}" --workers 1 &
office_pid=$!

trap 'kill -TERM "$pdf_pid" "$office_pid" 2>/dev/null || true' TERM INT

set +e
wait -n "$pdf_pid" "$office_pid"
status=$?
kill -TERM "$pdf_pid" "$office_pid" 2>/dev/null
wait
exit "$status"
