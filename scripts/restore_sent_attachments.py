#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 agently 已发送邮件下载附件（恢复误删协议用）。
用法: python restore_sent_attachments.py --msg <id> [--msg <id> ...] --output <目录>
"""

import argparse
import json
import os
import subprocess


def run(args, cwd=None):
    r = subprocess.run(["agently-cli"] + args, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=60, cwd=cwd)
    out = r.stdout
    try:
        return json.loads(out[out.find("{"):out.rfind("}") + 1])
    except Exception:
        return {"ok": False, "raw": out[:300], "err": r.stderr[:300]}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--msg", action="append", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    os.makedirs(args.output, exist_ok=True)
    for mid in args.msg:
        data = run(["message", "+read", "--id", mid])
        if not data.get("ok"):
            print("读取失败:", mid, data.get("error", data.get("raw", "")))
            continue
        for att in data.get("data", {}).get("attachments", []):
            aid = att.get("attachment_id")
            fn = att.get("filename")
            if not aid:
                continue
            rr = run(["attachment", "+download", "--msg", mid, "--att", aid, "--output", "."], cwd=args.output)
            print("下载:", mid, "|", fn, "|", "OK" if rr.get("ok") else rr)


if __name__ == "__main__":
    main()
