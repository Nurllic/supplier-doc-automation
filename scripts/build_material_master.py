#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""回填物料主数据：从 history.json / _supplier_mapping.json / _party_c_std_map.json / email-database/QA 汇总。
输出 data/_material_master.json 与 data/_master_missing_report.json（缺失清单）。"""

import glob
import json
import os
import re
import sys

PROJ_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "agreement-automation")
sys.path.insert(0, PROJ_DIR)

from material_master import (  # noqa: E402
    DATA_DIR, load_master, save_master, _merge_field, _norm,
    set_standard, check_missing_fields, to_indicators_dict,
)
from generate_qa_agreement_v3 import (  # noqa: E402
    parse_raw_material, parse_quality_standard, extract_part_number, extract_spec,
)

QA_DIR = os.path.join(os.path.dirname(PROJ_DIR), "email-database", "QA")


def is_wps_temp(fn):
    """WPS 临时/备份文件：~$ 开头或 名称_YYYYMMDDHHMMSS.docx/xlsx"""
    import re as _re
    return fn.startswith("~$") or bool(_re.search(r"_\d{14}\.(docx|xlsx)$", fn))


_STD_SUPPLIER_WORDS = [
    "Thermo Scientific", "Thermo", "Beckman Coulter", "Hanna Instruments", "Hanna",
    "BrandA", "BrandB", "BrandC", "BrandD",
    "博爱新开源", "博爱", "新开源", "Dongguan", "TianJin",
]


def std_name_key(fn):
    """标准文件名 → 名称关键词（去掉编号前缀/质量标准后缀/供应商词）"""
    base = os.path.splitext(os.path.basename(fn))[0]
    base = re.sub(r"^\s*\d+\s*[-_ ]?", "", base)
    for kw in ["原料企业质量标准", "企业质量标准", "质量标准"]:
        base = base.replace(kw, " ")
    for kw in _STD_SUPPLIER_WORDS:
        base = base.replace(kw, " ")
    return re.sub(r"\s+", "", base).strip()


def load_json(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return None


def ensure_code(master, code):
    code = str(code).strip()
    if not code or not code.isdigit():
        return None
    entry = master.setdefault(code, {"code": code, "suppliers": []})
    entry.setdefault("suppliers", [])
    return entry


def add_supplier(entry, supplier):
    if supplier and supplier not in entry["suppliers"]:
        entry["suppliers"].append(supplier)


def backfill_raw_file(master, fp):
    try:
        materials = parse_raw_material(fp)
    except Exception as e:
        print(f"  [跳过] 原辅料解析失败 {os.path.basename(fp)}: {e}")
        return
    for m in materials:
        entry = ensure_code(master, m.get("物料编码"))
        if not entry:
            continue
        _merge_field(entry, "name", m.get("物料中文名称") or m.get("物料名称"))
        _merge_field(entry, "spec", extract_spec(m.get("规格浓度", "")))
        _merge_field(entry, "part_no", extract_part_number(m.get("货号克隆号CAS", "")))
        _merge_field(entry, "brand", m.get("品牌产地"))
        _merge_field(entry, "storage", m.get("储存条件"))
        _merge_field(entry, "shelf_life", m.get("保质期"))
        _merge_field(entry, "main_unit", m.get("主计量单位"))
        _merge_field(entry, "property", m.get("物料属性"))
        _merge_field(entry, "dangerous", m.get("是否危化品"))
        _merge_field(entry, "open_times", m.get("开瓶次数"))
        _merge_field(entry, "freeze_thaw_times", m.get("冻融次数"))


def backfill_std_file(master, code, fp):
    entry = ensure_code(master, code)
    if not entry or not os.path.exists(fp) or os.path.getsize(fp) <= 0:
        return
    changed = set_standard(entry, fp)
    if changed or not entry.get("indicators_snapshot"):
        try:
            parsed = parse_quality_standard(fp)
            entry["indicators_snapshot"] = to_indicators_dict(parsed)
        except Exception as e:
            print(f"  [警告] 标准解析失败 {os.path.basename(fp)}: {e}")


def main():
    master = load_master()

    # 1) history.json（供应商归档 + 版本明细 + 归档文件）
    for hf in glob.glob(os.path.join(DATA_DIR, "*", "history.json")):
        history = load_json(hf)
        if not history:
            continue
        supplier_dir = os.path.dirname(hf)
        supplier = history.get("supplier_name", os.path.basename(supplier_dir))
        for version in history.get("versions", []):
            for md in version.get("material_details", []):
                entry = ensure_code(master, md.get("code"))
                if entry:
                    _merge_field(entry, "name", md.get("name"))
                    _merge_field(entry, "spec", md.get("spec"))
                    add_supplier(entry, supplier)
            for code in (version.get("new_material_codes", []) + history.get("all_material_codes", [])):
                entry = ensure_code(master, code)
                if entry:
                    add_supplier(entry, supplier)
            for rf in version.get("raw_files", []):
                fp = os.path.join(supplier_dir, rf)
                if os.path.exists(fp):
                    backfill_raw_file(master, fp)
            std_files = version.get("std_files", {})
            items = std_files.items() if isinstance(std_files, dict) else [(None, p) for p in std_files]
            for code, rel in items:
                fp = rel if os.path.isabs(rel) else os.path.join(supplier_dir, rel)
                if os.path.exists(fp):
                    backfill_std_file(master, code, fp)

    # 2) _party_c_std_map.json
    std_map = load_json(os.path.join(DATA_DIR, "_party_c_std_map.json")) or {}
    for code, fp in std_map.items():
        fp = os.path.normpath(fp)
        if os.path.exists(fp):
            backfill_std_file(master, code, fp)

    # 3) email-database/QA 全量扫描（补历史缺失，明细优先）
    if os.path.exists(QA_DIR):
        for fn in sorted(os.listdir(QA_DIR)):
            if is_wps_temp(fn):
                continue
            fp = os.path.join(QA_DIR, fn)
            if not fn.endswith((".docx", ".xlsx")) or os.path.getsize(fp) <= 0:
                continue
            if "原辅料" in fn:
                backfill_raw_file(master, fp)

        # 标准文件：先按文件名物料编码匹配，再按主数据名称匹配
        std_files = []
        for fn in sorted(os.listdir(QA_DIR)):
            if is_wps_temp(fn):
                continue
            if not fn.endswith((".docx", ".doc")) or "质量" not in fn:
                continue
            if fn.endswith(".doc") and os.path.exists(os.path.splitext(os.path.join(QA_DIR, fn))[0] + ".docx"):
                continue
            if os.path.getsize(os.path.join(QA_DIR, fn)) > 0:
                std_files.append(os.path.join(QA_DIR, fn))
        import re
        for fp in std_files:
            base = os.path.basename(fp)
            m = re.search(r"(\d{8,12})", base)
            if m:
                backfill_std_file(master, m.group(1), fp)
        for fp in std_files:
            base = os.path.basename(fp).replace("（", "(").replace("）", ")").replace(".docx", "")
            for code, entry in master.items():
                if _norm(entry.get("std_file")):
                    continue
                name = _norm(entry.get("name")).replace("（", "(").replace("）", ")")
                key = std_name_key(fp).replace("（", "(").replace("）", ")")
                norm_name = re.sub(r"\s+", "", name)
                if key and len(key) >= 2 and (key in norm_name or norm_name in key):
                    backfill_std_file(master, code, fp)
                    break

    # 4) _supplier_mapping.json 兜底（只补空缺，不覆盖明细解析结果）
    mapping = load_json(os.path.join(DATA_DIR, "_supplier_mapping.json")) or {}
    for supplier, items in mapping.items():
        for item in items:
            entry = ensure_code(master, item.get("code"))
            if not entry:
                continue
            _merge_field(entry, "name", item.get("name"))
            _merge_field(entry, "platform", item.get("platform"))
            add_supplier(entry, supplier)

    # 5) 完整性检查 + 报告
    missing_report = {}
    for code, entry in sorted(master.items()):
        missing = check_missing_fields(entry)
        if missing:
            missing_report[code] = {"name": entry.get("name", ""), "missing": missing}

    save_master(master)
    report_path = os.path.join(DATA_DIR, "_master_missing_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump({
            "total": len(master),
            "incomplete": len(missing_report),
            "items": missing_report,
        }, f, ensure_ascii=False, indent=2)

    print(f"主数据物料总数: {len(master)}")
    print(f"不完整物料数: {len(missing_report)}")
    print(f"报告: {report_path}")
    for code, info in missing_report.items():
        print(f"  {code} {info['name']}: 缺 {', '.join(info['missing'])}")


if __name__ == "__main__":
    main()
