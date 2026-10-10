#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""物料主数据：以物料编码为唯一键，维护最新物料信息与质量标准，供新增/升版共用。"""

import hashlib
import json
import os
import re
from datetime import datetime

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "data")
MASTER_PATH = os.path.join(DATA_DIR, "_material_master.json")
_NUMERIC_PLACEHOLDER = re.compile(r"^\d{3,}$")


def load_master(path=MASTER_PATH):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_master(master, path=MASTER_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _norm(value):
    if value is None:
        return ""
    s = str(value).strip().replace("\r", " ").replace("\n", " ")
    return "" if s in ("", "/", "-", "nan", "None") else s


def _merge_field(entry, key, value):
    v = _norm(value)
    if v and not _norm(entry.get(key)):
        entry[key] = v


def check_missing_fields(entry):
    """升版完整性门禁：返回缺失字段中文名列表"""
    missing = []
    required = [
        ("name", "物料名称"),
        ("spec", "规格/浓度"),
        ("part_no", "货号"),
        ("brand", "品牌/产地"),
        ("storage", "储存条件"),
        ("std_file", "质量标准文件"),
        ("indicators_snapshot", "质量标准指标"),
    ]
    for key, label in required:
        value = entry.get(key)
        if key == "indicators_snapshot":
            if not isinstance(value, dict) or not value:
                missing.append(label)
        elif key == "name":
            v = _norm(value)
            if not v or _NUMERIC_PLACEHOLDER.match(v):
                missing.append(label)
        elif not _norm(value):
            missing.append(label)
    return missing


def apply_supplement(entry, supplement_text):
    """解析钉钉表"物料信息补充"列（格式：字段：值;字段：值）"""
    if not supplement_text:
        return
    mapping = {
        "规格": "spec", "浓度": "spec",
        "货号": "part_no",
        "品牌": "brand", "产地": "brand",
        "储存条件": "storage", "储存": "storage",
        "保质期": "shelf_life",
        "主计量单位": "main_unit",
        "物料属性": "property",
        "开瓶次数": "open_times",
        "冻融次数": "freeze_thaw_times",
    }
    text = str(supplement_text).replace("；", ";").replace("\n", ";")
    for part in text.split(";"):
        for sep in ("：", ":"):
            if sep in part:
                key, value = part.split(sep, 1)
                target = mapping.get(key.strip())
                if target:
                    _merge_field(entry, target, value)
                break


def set_standard(entry, std_file):
    """登记标准文件：hash 变化则 std_version+1；返回是否发生变化"""
    if not std_file or not os.path.exists(std_file) or os.path.getsize(std_file) <= 0:
        return False
    current_hash = file_hash(std_file)
    if _norm(entry.get("std_hash")) == current_hash:
        if not _norm(entry.get("std_file")):
            entry["std_file"] = std_file
        return False
    if _norm(entry.get("std_file")):
        entry["std_version"] = int(entry.get("std_version", 1)) + 1
    else:
        entry["std_version"] = 1
    entry["std_file"] = std_file
    entry["std_hash"] = current_hash
    entry["std_updated_at"] = datetime.now().isoformat()
    return True


def diff_snapshots(old, new):
    """新旧标准指标快照 diff，返回 (新增项, 删除项, 修改项)"""
    old = old or {}
    new = new or {}
    added = [k for k in new if k not in old]
    removed = [k for k in old if k not in new]
    changed = [k for k in old if k in new and old[k] != new[k]]
    return added, removed, changed


def to_indicators_dict(parsed):
    """把 parse_quality_standard 的 dict 或 list 统一成 {检验项目: {标准值, 测试方法}}"""
    if isinstance(parsed, dict):
        return {k: {"标准值": v.get("标准值", "/"), "测试方法": v.get("测试方法", "/")}
                for k, v in parsed.items()}
    return {i.get("检验项目", ""): {"标准值": i.get("标准值", "/"), "测试方法": i.get("测试方法", "/")}
            for i in (parsed or []) if i.get("检验项目")}
