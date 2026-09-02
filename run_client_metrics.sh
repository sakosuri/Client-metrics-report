#!/usr/bin/env bash

# Scheduled execution wrapper for Linux/macOS.
# This script is intended to be run by Linux or macOS
# to execute the Catalyst Center client metrics report automatically.
# change BASE_DIR and PYTHON_BIN according to your script

set -euo pipefail

BASE_DIR=<path_to_client_metrics_report_directory>
PYTHON_BIN=<path_to_python_executable>
LOG_DIR="$BASE_DIR/logs"

mkdir -p "$LOG_DIR"

cd "$BASE_DIR"

LOG_FILE="$LOG_DIR/client_metrics_report_$(date +%Y%m%d).log"

echo "===== Run started: $(date '+%Y-%m-%d %H:%M:%S %Z') =====" >> "$LOG_FILE"

set +e
"$PYTHON_BIN" client_metrics_report.py \
  --time-range 24h \
  >> "$LOG_FILE" 2>&1
EXIT_CODE=$?
set -e

echo "===== Run finished: $(date '+%Y-%m-%d %H:%M:%S %Z') (exit code $EXIT_CODE) =====" >> "$LOG_FILE"

exit $EXIT_CODE
