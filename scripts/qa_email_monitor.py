#!/usr/bin/env python3
import subprocess, sys
script = r".\scripts\qa_email_monitor_main.py"
r = subprocess.run([sys.executable, script], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
print(r.stdout, end="")
if r.stderr:
    print(r.stderr, file=sys.stderr, end="")
sys.exit(r.returncode)
