#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""第⑧阶段：供应商回签/回传处理。
识别回传邮件 → 匹配供应商 → 归档原件 → 提取乙方信息 → 条款防篡改比对 → 回写钉钉状态 → 群通知。"""

import difflib
import glob
import json
import os
import re
import shutil
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "agreement-automation"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document  # noqa: E402

from generate_qa_agreement_v3 import agreement_lines, snapshot_agreement_text  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "data")

RETURN_KEYWORDS = ["回传", "回签", "盖章", "已盖章", "乙方信息"]


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def is_return_email(subject, filenames):
    text = (subject or "") + " " + " ".join(filenames or [])
    return any(k in text for k in RETURN_KEYWORDS)


def normalize_serial(s):
    return re.sub(r"^(QA-DOC|QA)[-_]?", "", str(s or "")).strip()


def match_supplier(fname, subject):
    """按 供应商名 / 流水号 匹配 data/ 下的供应商 history"""
    text = (fname or "") + " " + (subject or "")
    for company_dir in sorted(glob.glob(os.path.join(DATA_DIR, "*"))):
        if not os.path.isdir(company_dir):
            continue
        company = os.path.basename(company_dir)
        for supplier_dir in sorted(glob.glob(os.path.join(company_dir, "*"))):
            hp = os.path.join(supplier_dir, "history.json")
            if not os.path.exists(hp):
                continue
            supplier = os.path.basename(supplier_dir)
            with open(hp, encoding="utf-8") as f:
                history = json.load(f)
            versions = history.get("versions", [])
            for v in versions:
                v.setdefault("company_folder", company)
                v.setdefault("company_tag", "平台C" if company == "示例丙公司" else "生物")
            if supplier in text:
                return supplier, history, (versions[-1] if versions else None)
            for v in reversed(versions):
                serial = normalize_serial(v.get("oa_serial") or v.get("oa_serial_temp") or "")
                if serial and serial in text:
                    return supplier, history, v
    return None, None, None


def extract_party_b_info(docx_path):
    """从协议第二页【乙方】区域提取 住所/联系人/电话/邮箱"""
    doc = Document(docx_path)
    for table in doc.tables:
        if len(table.rows) < 6 or len(table.columns) < 2:
            continue
        if "【乙方】" not in table.rows[0].cells[1].text:
            continue
        info = {"name": "", "address": "", "contact": "", "phone": "", "email": "", "date": ""}
        name_txt = table.rows[0].cells[1].text.strip()
        if name_txt.startswith("【乙方】"):
            info["name"] = name_txt.replace("【乙方】", "", 1).strip()
        for ri, key in ((1, "address"), (2, "contact"), (3, "phone"), (4, "email")):
            txt = table.rows[ri].cells[1].text.strip()
            if "：" in txt:
                info[key] = txt.split("：", 1)[1].strip()
            elif txt:
                info[key] = txt
        date_txt = table.rows[5].cells[1].text.strip()
        if "：" in date_txt:
            info["date"] = date_txt.split("：", 1)[1].strip()
        return info
    return {}


def update_supplier_info(supplier_dir, info, returned_file, oa_serial, verified, diff_summary):
    p = os.path.join(supplier_dir, "supplier_info.json")
    data = {}
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
    data.setdefault("versions", []).append({
        "time": datetime.now().isoformat(),
        "returned_file": os.path.relpath(returned_file, supplier_dir),
        "oa_serial": oa_serial,
        "info": info,
        "verified": verified,
        "diff_summary": diff_summary or "",
    })
    data["latest"] = info
    data["updated_at"] = datetime.now().isoformat()
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def verify_agreement(sent_lines, returned_path):
    """内容级比对：返回 (是否一致, 摘要)"""
    returned_lines = agreement_lines(returned_path)
    if returned_lines == sent_lines:
        return True, ""
    diff = list(difflib.unified_diff(sent_lines, returned_lines, lineterm="", n=1))
    return False, "\n".join(diff[:40]) or "内容存在差异"


def update_records_status(cfg, record_ids, status):
    import requests
    if not record_ids:
        return True
    token_resp = requests.post(
        "https://api.dingtalk.com/v1.0/oauth2/accessToken",
        json={"appKey": cfg["app_key"], "appSecret": cfg["app_secret"]},
        timeout=10,
    )
    token = token_resp.json().get("accessToken", "")
    if not token:
        return False
    headers = {
        "x-acs-dingtalk-access-token": token,
        "Content-Type": "application/json",
        "operatorId": cfg["operator_id"],
    }
    url = f"https://api.dingtalk.com/v1.0/notable/bases/{cfg['base_id']}/sheets/{cfg['sheet2_id']}/records"
    payload = {"records": [{"id": rid, "fields": {"节点状态": status}} for rid in record_ids]}
    r = requests.put(url, headers=headers, json=payload, timeout=15)
    return r.status_code == 200


def notify(cfg, title, text):
    import base64
    import hashlib
    import hmac
    import time
    import requests
    url = cfg.get("webhook", "")
    if not url:
        return False
    secret = cfg.get("robot_secret", "")
    if secret and secret.strip():
        timestamp = str(round(time.time() * 1000))
        sign_str = timestamp + "\n" + secret
        signature = base64.b64encode(
            hmac.new(secret.encode(), sign_str.encode(), digestmod=hashlib.sha256).digest()
        ).decode()
        sep = "&" if "?" in url else "?"
        url = url + sep + "timestamp=" + timestamp + "&sign=" + signature
    payload = {"msgtype": "markdown", "markdown": {"title": title, "text": text}}
    try:
        r = requests.post(url, json=payload, timeout=10)
        return r.json().get("errcode") == 0
    except Exception:
        return False


def process_returned_files(files, subject="", cfg=None, dry_run=False):
    docx_files = [f for f in (files or []) if str(f).lower().endswith((".docx", ".doc"))]
    if not docx_files:
        log("回传邮件中未找到 docx 附件")
        return "no_docx"
    target = docx_files[0]
    supplier, history, version = match_supplier(os.path.basename(target), subject)
    if not supplier:
        unmatched_dir = os.path.join(DATA_DIR, "_returned_unmatched")
        os.makedirs(unmatched_dir, exist_ok=True)
        dst = os.path.join(unmatched_dir, os.path.basename(target))
        shutil.copy2(target, dst)
        log(f"无法自动匹配供应商，已放入 {dst}")
        if cfg and not dry_run:
            notify(cfg, "供应商回传待人工匹配", f"回传文件无法自动匹配供应商：{os.path.basename(target)}")
        return "unmatched"

    company_folder = (version or {}).get("company_folder") or (
        "示例丙公司"
        if (version or {}).get("company_tag") == "平台C" or "【平台C】" in ((version or {}).get("output_file") or "")
        else "示例甲公司"
    )
    supplier_dir = os.path.join(DATA_DIR, company_folder, supplier)
    returned_dir = os.path.join(supplier_dir, "returned")
    os.makedirs(returned_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    archived = os.path.join(returned_dir, f"v{ts}_{os.path.basename(target)}")
    shutil.copy2(target, archived)

    info = extract_party_b_info(archived)
    sent_lines = (version or {}).get("content_lines") or []
    verified, diff_summary = verify_agreement(sent_lines, archived)
    serial = (version or {}).get("oa_serial") or ""
    update_supplier_info(supplier_dir, info, archived, serial, verified, diff_summary if not verified else "")

    status = "已归档闭环" if verified else "待人工复核"
    record_ids = (version or {}).get("record_ids") or []
    if cfg and not dry_run:
        if record_ids:
            ok = update_records_status(cfg, record_ids, status)
            log(("钉钉状态已更新: " if ok else "钉钉状态更新失败: ") + status + f"（{len(record_ids)}条）")
        title = "供应商已回传（校验通过）" if verified else "供应商已回传（需人工复核）"
        text = (
            f"### {title}\n\n"
            f"- **供应商：** {supplier}\n"
            f"- **流水号：** {serial or '—'}\n"
            f"- **乙方联系人：** {info.get('contact') or '—'}\n"
            f"- **电话：** {info.get('phone') or '—'}\n"
            f"- **邮箱：** {info.get('email') or '—'}\n"
        )
        if not verified:
            text += f"- **差异摘要：**\n```\n{diff_summary[:600]}\n```\n"
        notify(cfg, title, text)

    log(f"回传处理完成: 供应商={supplier} 校验={'通过' if verified else '不一致'} 归档={archived}")
    return "verified" if verified else "mismatch"


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--subject", default="供应商回传")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    from qa_config_loader import load_qa_config
    cfg = None if args.dry_run else load_qa_config()
    result = process_returned_files([args.file], args.subject, cfg=cfg, dry_run=args.dry_run)
    print("RESULT:", result)
