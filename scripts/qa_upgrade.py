#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""物料升版处理：主数据更新、完整性门禁、变更摘要、钉钉记录回写。"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "agreement-automation"))

from material_master import (  # noqa: E402
    _norm, apply_supplement, check_missing_fields, diff_snapshots,
    set_standard, to_indicators_dict,
)


def detect_change_type(fields, subject=""):
    """从钉钉记录字段或邮件主题识别变更类型（标准升版/信息升版/混合升版/升版）"""
    raw = fields.get("变更类型", "")
    if isinstance(raw, dict):
        raw = raw.get("text") or raw.get("name") or ""
    if _norm(raw):
        return str(raw).strip()
    text = (subject or "") + "|" + str(fields.get("升版原因", ""))
    if "升版" in text:
        has_std = "标准" in text
        has_info = "信息" in text or "明细" in text
        if has_std and has_info:
            return "混合升版"
        if has_std:
            return "标准升版"
        if has_info:
            return "信息升版"
        return "升版"
    return ""


def apply_record_to_master(master, code, fields):
    """把钉钉记录里的物料名称/指定品牌/业务平台/物料信息补充写入主数据"""
    entry = master.setdefault(code, {"code": code, "suppliers": []})
    entry.setdefault("suppliers", [])
    for key, col in [("name", "物料名称"), ("brand", "指定品牌"), ("platform", "业务平台")]:
        raw = fields.get(col, "")
        if isinstance(raw, dict):
            raw = raw.get("text") or raw.get("name") or ""
        if _norm(raw):
            entry[key] = str(raw).strip()
    supplement = fields.get("物料信息补充", "")
    if isinstance(supplement, dict):
        supplement = supplement.get("text", "")
    apply_supplement(entry, supplement)
    return entry


def upgrade_standard(master, code, new_std_file, parsed=None):
    """登记新标准并返回 (entry, 变更摘要)。parsed 为新解析出的指标快照。"""
    entry = master.setdefault(code, {"code": code, "suppliers": []})
    entry.setdefault("suppliers", [])
    old_snapshot = dict(entry.get("indicators_snapshot") or {})
    changed = set_standard(entry, new_std_file)
    if parsed is not None:
        entry["indicators_snapshot"] = to_indicators_dict(parsed)
    new_snapshot = dict(entry.get("indicators_snapshot") or {})
    added, removed, modified = diff_snapshots(old_snapshot, new_snapshot)
    summary = {
        "standard_changed": changed,
        "std_version": entry.get("std_version", 1),
        "added": added,
        "removed": removed,
        "changed": modified,
    }
    return entry, summary


def find_latest_std_file(qa_dir, code, name=""):
    """在 QA_DIR 中按编码/名称匹配质量标准文件，取修改时间最新的（升版场景）"""
    if not os.path.exists(qa_dir):
        return ""
    candidates = []
    norm_name = (name or "").replace("(", "（").replace(")", "）").replace("β", "beta")
    for fn in os.listdir(qa_dir):
        if not fn.endswith((".docx", ".doc")) or "质量" not in fn:
            continue
        fp = os.path.join(qa_dir, fn)
        if os.path.getsize(fp) <= 0:
            continue
        if fn.endswith(".doc") and os.path.exists(os.path.splitext(fp)[0] + ".docx"):
            continue
        if code and code in fn:
            candidates.append(fp)
            continue
        if norm_name and (norm_name in fn or name in fn):
            candidates.append(fp)
    if not candidates:
        return ""
    return max(candidates, key=lambda p: os.path.getmtime(p))


def update_record_fields(cfg, record_ids, fields):
    """更新钉钉AI表格记录字段，返回是否全部成功"""
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
    records = [{"id": rid, "fields": fields} for rid in record_ids]
    url = f"https://api.dingtalk.com/v1.0/notable/bases/{cfg['base_id']}/sheets/{cfg['sheet2_id']}/records"
    r = requests.put(url, headers=headers, json={"records": records}, timeout=15)
    return r.status_code == 200


def summary_text(summary):
    parts = []
    if summary.get("standard_changed"):
        parts.append(f"标准 v{summary.get('std_version', '?')}")
    if summary.get("added"):
        parts.append("新增指标: " + "、".join(summary["added"][:8]))
    if summary.get("removed"):
        parts.append("删除指标: " + "、".join(summary["removed"][:8]))
    if summary.get("changed"):
        parts.append("修改指标: " + "、".join(summary["changed"][:8]))
    return "；".join(parts) if parts else "无标准变化"
