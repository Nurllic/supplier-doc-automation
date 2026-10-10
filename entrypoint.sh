#!/bin/bash
# Entrypoint for supplier-doc-automation Docker services
set -e

case "$1" in
    monitor)
        echo "[QA-Monitor] Starting email monitor..."
        # Run email monitor in loop (every 15 minutes)
        while true; do
            python3 /app/scripts/qa_email_monitor_main.py
            echo "[QA-Monitor] Sleep 15min..."
            sleep 900
        done
        ;;
    confirmer)
        echo "[QA-Confirmer] Starting confirm monitor..."
        # Run confirm monitor in loop (every 1 hour)
        while true; do
            python3 /app/scripts/qa_confirm_monitor.py
            echo "[QA-Confirmer] Sleep 1h..."
            sleep 3600
        done
        ;;
    sender)
        echo "[QA-Sender] Starting email sender..."
        # Run email sender in loop (every 5 minutes)
        while true; do
            python3 /app/scripts/send_email_tasks.py
            echo "[QA-Sender] Sleep 5min..."
            sleep 300
        done
        ;;
    once)
        echo "[QA] Running single shot..."
        python3 /app/scripts/qa_email_monitor_main.py
        ;;
    *)
        echo "Usage: $0 {monitor|confirmer|once}"
        exit 1
        ;;
esac
