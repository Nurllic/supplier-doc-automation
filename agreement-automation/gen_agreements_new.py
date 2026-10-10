
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate QA agreements using the word-doc-extractor skill parser.
Fixes: superscript preservation, field splitting, quality standard parsing.
"""


import os, sys, re, json, logging
from datetime import date
from copy import deepcopy

# Add skill to path for parsing
SKILL_PATH = r"C:\Users\5532\.workbuddy\skills\word-doc-extractor\scripts"
sys.path.insert(0, SKILL_PATH)
if "extract_doc" in sys.modules:
    del sys.modules["extract_doc"]
from extract_doc import (
    DocxParser, extract_table_by_signature, parse_quality_standard,
    split_storage_shelf, split_appearance_danger, get_text, tag, NS
)

# Add project path for rendering
QA_DIR = r".\agreement-automation"
sys.path.insert(0, QA_DIR)
from generate_qa_agreement_v3 import (
    TEMPLATE_DOCX, get_platform_contact, sanitize_filename,
    _format_cell, _set_col_widths, _set_table_full_width, merge_cells_vertically,
    rebuild_material_table, extract_spec, extract_part_number,
    replace_placeholder,
)
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
import pandas as pd

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def _or_slash(val):
    if val is None:
        return "/"
    v = str(val).strip()
    if v == "" or v == "/":
        return "/"
    return v


def _extract_huohao_only(cas_field):
    if not cas_field or cas_field == "/":
        return ""
    m = re.search(r"货号[：:]\s*([^C;；\n]+?)(?:CAS|克隆|$)", cas_field)
    if m:
        return m.group(1).strip()
    m2 = re.search(r"货号[：:]\s*([^;；\n]+)", cas_field)
    if m2:
        return m2.group(1).strip()
    return ""


def parse_raw_material_new(docx_path, material_filter=None):
    result = extract_table_by_signature(docx_path, ["物料编码", "储存条件", "外观性状"])
    if result["status"] != "success":
        raise ValueError(result["message"])
    materials = []
    for idx, row in enumerate(result["data"]):
        code = row.get("物料编码", "")
        if material_filter is not None and code not in material_filter:
            continue
        storage_cond = row.get("储存条件", "")
        shelf_life = row.get("保质期", "/")
        appearance = row.get("外观性状", "")
        is_danger = row.get("是否危化品", "否")
        spec_flat = row.get("规格/浓度", "").replace("\n", "; ")
        brand_raw = row.get("品牌/产地", "").replace("\n", "; ")
        cas_raw = row.get("货号/克隆号或CAS号（如有）", "").replace("\n", "; ")
        huohao = _extract_huohao_only(cas_raw)
        mat_name = row.get("物料名称", "").strip()
        mat_cn_name = row.get("物料中文名称", "").strip()
        mat_secret = row.get("物料保密编码", "").strip()
        if (not mat_name or mat_name == "/") and (not mat_cn_name or mat_cn_name == "/"):
            if mat_secret and mat_secret != "/":
                mat_name = mat_secret
                mat_cn_name = mat_secret
        brand_secret = row.get("产地保密编码", "").strip()
        if (not brand_raw or brand_raw.strip() in ("", "/")) and brand_secret and brand_secret != "/":
            brand_flat = brand_secret
        else:
            brand_flat = brand_raw
        material = {
            "序号": str(idx + 1).zfill(2),
            "物料编码": code,
            "物料名称": _or_slash(mat_name),
            "物料中文名称": _or_slash(mat_cn_name),
            "规格浓度": _or_slash(spec_flat),
            "主计量单位": _or_slash(row.get("主计量单位", "")),
            "货号克隆号CAS": _or_slash(cas_raw),
            "品牌产地": _or_slash(brand_flat),
            "是否危化品": is_danger,
            "物料属性": _or_slash(row.get("物料属性", "")),
            "储存条件": _or_slash(storage_cond),
            "保质期": _or_slash(shelf_life),
            "开瓶次数": _or_slash(row.get("开瓶次数", "")),
            "冻融次数": _or_slash(row.get("冻融次数", "")),
            "外观性状": _or_slash(appearance),
        }
        materials.append(material)
    logger.info(f"解析到 {len(materials)} 条原辅料 (新解析器)")
    return materials


def parse_quality_standard_new(docx_path):
    """Parse quality standard using skill's parser (with superscript, sub-items, etc.)."""
    result = parse_quality_standard(docx_path)
    if result["status"] != "success":
        logger.warning(f"质量标准解析失败: {result.get('message', '')}")
        return []
    
    indicators = []
    for ind in result["data"]:
        indicators.append({
            "检验项目": ind["检验项目"],
            "标准值": ind["标准值"],
            "测试方法": ind["测试方法"],
        })
    
    logger.info(f"解析到 {len(indicators)} 个质量指标 (新解析器)")
    return indicators


def build_multi_quality_table(doc, material_groups, supplier_name, company_name):
    """Build the quality standard table with multiple materials."""
    if len(doc.tables) < 4:
        logger.warning("不足4个表格")
        return
    old_tbl = doc.tables[3]._tbl
    parent = old_tbl.getparent()
    headers = ["物料编码", "货号", "中/英文品名", "规格", "检验项目", "标准值", "测试方法", "备注"]
    note_text = "以上产品需满足：包装牢固、严密；标签及各种标志内容准确，牢固、端正、完整、清洁。"
    row_count = sum(max(1, len(g["indicators"])) for g in material_groups)
    total_rows = 1 + row_count + 1 + 2
    logger.info("data_rows=%d total_rows=%d", row_count, total_rows)
    new_table = doc.add_table(rows=total_rows, cols=8)
    new_table.style = 'Table Grid'
    parent.replace(old_tbl, new_table._tbl)
    for ci, h in enumerate(headers):
        _format_cell(new_table.rows[0].cells[ci], h, font_size=Pt(10.5), bold=True, align='center')
    row_idx = 1
    for g in material_groups:
        inds = g["indicators"]
        if inds:
            for i_idx, ind in enumerate(inds):
                row = new_table.rows[row_idx]
                if i_idx == 0:
                    _format_cell(row.cells[0], g['material_code'], font_size=Pt(9), align='center')
                    _format_cell(row.cells[1], g['code'], font_size=Pt(9), align='center')
                    _format_cell(row.cells[2], g['name'], font_size=Pt(9), align='center')
                    _format_cell(row.cells[3], g['spec'], font_size=Pt(9), align='center')
                _format_cell(row.cells[4], ind['检验项目'], font_size=Pt(9),
                             align='left' if len(ind['检验项目']) > 8 else 'center')
                _format_cell(row.cells[5], ind['标准值'], font_size=Pt(9), align='left')
                _format_cell(row.cells[6], ind['测试方法'], font_size=Pt(9), align='left')
                _format_cell(row.cells[7], "/", font_size=Pt(9), align='center')
                row_idx += 1
            if len(inds) > 1:
                start = row_idx - len(inds)
                end = row_idx - 1
                for col in range(4):
                    merge_cells_vertically(new_table, col, start, end)
        else:
            row = new_table.rows[row_idx]
            _format_cell(row.cells[0], g['material_code'], font_size=Pt(9), align='center')
            _format_cell(row.cells[1], g['code'], font_size=Pt(9), align='center')
            _format_cell(row.cells[2], g['name'], font_size=Pt(9), align='center')
            _format_cell(row.cells[3], g['spec'], font_size=Pt(9), align='center')
            _format_cell(row.cells[4], '/', font_size=Pt(9), align='center')
            _format_cell(row.cells[5], '(暂无质量标准)', font_size=Pt(9), align='center')
            _format_cell(row.cells[6], '/', font_size=Pt(9), align='center')
            _format_cell(row.cells[7], '/', font_size=Pt(9), align='center')
            row_idx += 1
    
    # 注释行
    note_row = new_table.rows[-3]
    nc = note_row.cells[0]
    for mc in range(1, 8):
        try:
            nc.merge(note_row.cells[mc])
            if mc < len(note_row.cells):
                note_row.cells[mc].text = ""
        except:
            break
    _format_cell(nc, note_text, font_size=Pt(8.5), align='left')
    
    # 签章行
    sig1 = new_table.rows[-2]
    sig2 = new_table.rows[-1]
    sig1.cells[0].merge(sig1.cells[1]); sig1.cells[0].merge(sig1.cells[2]); sig1.cells[0].merge(sig1.cells[3])
    _format_cell(sig1.cells[0], "甲方：" + company_name, font_size=Pt(10), bold=True, align="center")
    sig1.cells[4].merge(sig1.cells[5])
    _format_cell(sig1.cells[4], "乙方：" + supplier_name, font_size=Pt(10), bold=True, align="center")
    sig1.cells[6].merge(sig1.cells[7])
    _format_cell(sig1.cells[6], "日期： 年 月 日", font_size=Pt(9), align="center")
    sig2.cells[0].merge(sig2.cells[1]); sig2.cells[0].merge(sig2.cells[2]); sig2.cells[0].merge(sig2.cells[3])
    _format_cell(sig2.cells[0], "（盖章）", font_size=Pt(10), bold=True, align="center")
    sig2.cells[4].merge(sig2.cells[5])
    _format_cell(sig2.cells[4], "（盖章）", font_size=Pt(10), bold=True, align="center")
    try: sig2.cells[6].merge(sig2.cells[7])
    except: pass
    _set_col_widths(new_table, [1.5, 2.0, 2.5, 1.5, 2.5, 4.5, 4.5, 2.0])
    _set_table_full_width(new_table)
    logger.info("质量标准表: %d 指标", row_count)


def generate_agreement(raw_material_docx, quality_std_dir, supplier_name, platform_name,
                       oa_serial, output_path, material_filter=None):
    """Generate a QA agreement using the new skill-based parser."""
    logger.info("=" * 50)
    logger.info(f"生成协议: {supplier_name}")
    logger.info("=" * 50)
    
    doc = Document(TEMPLATE_DOCX)
    
    # 1. Parse raw material table
    logger.info("解析原辅料明细表...")
    materials = parse_raw_material_new(raw_material_docx, material_filter)
    
    # 2. Parse quality standards
    logger.info("解析企业质量标准...")
    # Build huohao map from raw materials
    huohao_map = {}
    for m in materials:
        code = m["物料编码"]
        cas_field = m.get("货号克隆号CAS", "")
        m_hh = re.search(r"货号[：:]\s*([^\n;]+)", cas_field)
        huohao_map[code] = _extract_huohao_only(cas_field) or None
    
    # Find all quality standard files
    std_files = []
    if os.path.isdir(quality_std_dir):
        for f in os.listdir(quality_std_dir):
            if f.endswith(".docx") and not f.startswith("~$"):
                std_files.append(os.path.join(quality_std_dir, f))
    elif os.path.isfile(quality_std_dir):
        std_files = [quality_std_dir]
    
    logger.info(f"找到 {len(std_files)} 个质量标准文件")
    
    # Match std files to materials
    def _match_std(std_path, mats):
        base_fn = os.path.basename(std_path).lower().replace("（", "(").replace("）", ")")
        for m in mats:
            code = m["物料编码"]
            if code in std_path:
                return m
        for m in mats:
            part_no = extract_part_number(m.get("货号克隆号CAS", ""))
            if part_no and part_no in std_path:
                return m
        for m in mats:
            mat_name = m.get("物料名称", "").lower().replace("（", "(").replace("）", ")")
            if mat_name and mat_name in base_fn:
                return m
        return None
    
    material_groups = []
    for m in materials:
        code = m["物料编码"]
        huohao = huohao_map.get(code)
        if huohao is None:
            huohao = "/"
        name = m.get("物料中文名称") or m.get("物料名称") or ""
        spec = extract_spec(m.get("规格浓度", ""))
        material_groups.append({
            "code": huohao,
            "material_code": code,
            "name": name,
            "spec": spec,
            "indicators": []
        })
    
    for sf in std_files:
        mat = _match_std(sf, materials)
        inds = parse_quality_standard_new(sf)
        if mat and inds:
            for g in material_groups:
                if g["material_code"] == mat["物料编码"]:
                    g["indicators"] = inds
                    break
            logger.info(f"  {os.path.basename(sf)} -> {mat['物料编码']} ({len(inds)} 指标)")
        elif not mat:
            logger.warning(f"  {os.path.basename(sf)} -> 未匹配到物料")
        elif not inds:
            logger.warning(f"  {os.path.basename(sf)} -> 无指标")
    
    n_with = sum(1 for g in material_groups if g["indicators"])
    logger.info(f"有质量标准: {n_with}/{len(material_groups)}")
    
    # 3. Get platform contact
    reviewer, phone, email = get_platform_contact(platform_name)
    company_name = "示例甲公司"
    company_address = "YY省YY市YY区YY路16号"
    if platform_name == "平台C":
        company_name = "示例丙公司"
        company_address = "XX省XX市XX区XX路1006号"
    
    today = date.today().strftime("%Y年%m月%d日")
    
    # 4. Replace placeholders
    replace_placeholder(doc, "{1}", supplier_name)
    replace_placeholder(doc, "{2}", reviewer)
    replace_placeholder(doc, "{3}", phone)
    replace_placeholder(doc, "{4}", email)
    replace_placeholder(doc, "{5}", today)
    replace_placeholder(doc, "{6}", oa_serial)
    
    # 5. Replace company name
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if "示例甲公司" in run.text:
                            run.text = run.text.replace("示例甲公司", company_name)
                        if "YY省YY市YY区YY路16号" in run.text:
                            run.text = run.text.replace("YY省YY市YY区YY路16号", company_address)
    
    # 6. Rebuild tables
    rebuild_material_table(doc, materials)
    build_multi_quality_table(doc, material_groups, supplier_name, company_name)
    
    # 7. Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    logger.info(f"OK: {output_path}")
    return output_path


if __name__ == "__main__":
    DATA_DIR = r".\data"
    
    # === 1. 供应商A公司 (16物料, 平台A) ===
    generate_agreement(
        raw_material_docx=os.path.join(DATA_DIR, "供应商A公司/materials/新增原辅料明细_v1.docx"),
        quality_std_dir=os.path.join(DATA_DIR, "供应商A公司/materials/质量标准_v1"),
        supplier_name="供应商A公司",
        platform_name="平台A",
        oa_serial="QA-20260715-001",
        output_path=os.path.join(DATA_DIR, "供应商A公司/agreements/v2_QA-20260715-001-【生物】-供应商A公司-16.docx"),
    )
    
    # === 2. 示例丙公司 (10物料, 平台C) ===
    generate_agreement(
        raw_material_docx=os.path.join(DATA_DIR, "示例丙公司/materials/新增原辅料明细_v1.docx"),
        quality_std_dir=os.path.join(DATA_DIR, "示例丙公司/materials/质量标准_v1"),
        supplier_name="示例丙公司",
        platform_name="平台A",
        oa_serial="QA-20260715-002",
        output_path=os.path.join(DATA_DIR, "示例丙公司/agreements/v2_QA-20260715-002-【生物】-示例丙公司-10.docx"),
    )
    
    print("\n=== 生成完成 ===")
