#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
质量保证协议自动生成系统 v3 - 动态重绘策略
占位符替换 -> 删除旧表 -> 新建美观表格
"""

import hashlib, os, re, logging, json, subprocess, sys
from datetime import date, datetime
from copy import deepcopy
from lxml import etree

from docx import Document
from docx.shared import Pt, Cm, Inches, Emu
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from docx.oxml import parse_xml
import pandas as pd

BASE_DIR = "/app/agreement-automation" if os.name != "nt" else r".\agreement-automation"
TEST_RAW_MATERIAL = ""
TEST_QUALITY_STD  = ""
CONTACT_XLSX      = os.path.join(BASE_DIR, "平台联系人的对应关系.xlsx")
TEMPLATE_DOCX     = os.path.join(BASE_DIR, "质量保证协议模板.docx")
OUTPUT_DIR        = os.path.join(BASE_DIR, "output") if os.path.exists(os.path.join(BASE_DIR, "output")) else "/app/output"

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# 质量标准中“应符合1.X的要求”的引用模式（fallback 解析用）
pat_ref = re.compile(r"符合\s*(\d+\.\d+)\s*的要求")

_STD_SUPPLIER_WORDS = [
    "Thermo Scientific", "Thermo", "Beckman Coulter", "Hanna Instruments", "Hanna",
    "BrandA", "BrandB", "BrandC", "BrandD",
    "博爱新开源", "博爱", "新开源", "Dongguan", "TianJin",
]


def std_filename_key(fn):
    """标准文件名 → 名称关键词（去掉编号前缀/质量标准后缀/供应商词，用于模糊匹配）"""
    base = os.path.splitext(os.path.basename(str(fn)))[0]
    base = re.sub(r"^\s*\d+\s*[-_ ]?", "", base)
    for kw in ["原料企业质量标准", "企业质量标准", "质量标准"]:
        base = base.replace(kw, " ")
    for kw in _STD_SUPPLIER_WORDS:
        base = base.replace(kw, " ")
    return re.sub(r"\s+", "", base).strip()


# word-doc-extr 技能路径
_EXTRACT_SCRIPT = os.environ.get("WORD_DOC_EXTR_SCRIPT", "")
if not _EXTRACT_SCRIPT:
    for _c in [r"C:\Users\5532\.workbuddy\skills\word-doc-extr\scripts\extract_doc.py",
               "/app/scripts/extract_doc.py",
               os.path.join(os.path.dirname(os.path.abspath(__file__)), "extract_doc.py")]:
        if os.path.exists(_c):
            _EXTRACT_SCRIPT = _c
            break


def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name)


def normalize_oa_serial(oa_serial):
    """统一流水号格式：QA-DOC-{流水号}（页眉/文件名共用）"""
    serial = str(oa_serial or "").strip()
    if not serial:
        serial = datetime.now().strftime("%Y%m%d-%H%M%S")
    if not serial.upper().startswith("QA-DOC"):
        serial = re.sub(r"^QA-", "", serial, flags=re.IGNORECASE)
        serial = "QA-DOC-" + serial
    return serial


def build_output_filename(oa_serial, tag, supplier_name):
    """输出文件命名：QA-DOC-流水号-【生物/平台C】-供应商名称.docx"""
    serial = normalize_oa_serial(oa_serial)
    return f"{serial}-【{tag}】-{sanitize_filename(supplier_name)}.docx"


def snapshot_agreement_text(docx_path):
    """协议内容快照（正文+表格，规范化空白后哈希），用于回传防篡改比对"""
    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt:
                    lines.append(txt)
    norm = re.sub(r"[ \t\u3000]+", "", "\n".join(lines))
    return hashlib.sha256(norm.encode("utf-8")).hexdigest()


def agreement_lines(docx_path):
    """协议文本行（规范化），用于差异清单"""
    doc = Document(docx_path)
    lines = [re.sub(r"[ \t\u3000]+", "", p.text.strip()) for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = re.sub(r"[ \t\u3000]+", "", cell.text.strip())
                if txt:
                    lines.append(txt)
    return lines


def fill_party_b_info(doc, info):
    """把乙方（供应商）信息回填到协议模板第二页【乙方】区域：
    公司名称/住所/联系人/联系电话/电子邮件/签订日期（不是签章行）。
    签订日期与甲方保持一致，由调用方传入 date。"""
    if not info:
        return
    for table in doc.tables:
        if len(table.rows) < 6 or len(table.columns) < 2:
            continue
        if "【乙方】" not in table.rows[0].cells[1].text:
            continue
        targets = [
            ("住所：", "address", 1),
            ("联系人：", "contact", 2),
            ("联系电话：", "phone", 3),
            ("电子邮件：", "email", 4),
        ]
        for label, key, ri in targets:
            val = str(info.get(key) or "").strip()
            if not val or val == "/":
                continue
            cell = table.rows[ri].cells[1]
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[-1].text = label + val
            else:
                cell.text = label + val
        name = str(info.get("name") or "").strip()
        if name:
            cell0 = table.rows[0].cells[1]
            if cell0.text.strip() == "【乙方】":
                if cell0.paragraphs and cell0.paragraphs[0].runs:
                    cell0.paragraphs[0].runs[-1].text = "【乙方】" + name
                else:
                    cell0.text = "【乙方】" + name
        d = str(info.get("date") or "").strip()
        if d:
            cell5 = table.rows[5].cells[1]
            cell5.text = "签订日期：" + d
        return


def extract_spec(concise_text):
    text = concise_text.replace("\n", ";").replace("\r", "")
    m = re.search(r"规格[：:]\s*([^;；]+)", text)
    if m:
        return m.group(1).strip()
    parts = text.split(";")
    first = parts[0].strip()
    if "：" in first:
        return first.split("：")[-1].strip()
    return text.strip()


def extract_part_number(cas_field):
    """从"货号：11\nCAS号：11\n克隆号：/"、"货号：B648003; 克隆号：/"或纯值"979878"中提取货号"""
    if not cas_field or cas_field == "/":
        return ""
    m = re.search(r"货号[：:]\s*([^\n;；]+?)(?:\s*(?:CAS号|克隆号)[：:]|$)", cas_field)
    if m:
        return m.group(1).strip()
    if "CAS号" in cas_field or "克隆号" in cas_field:
        return ""  # 有标签但无货号 → 视为无货号
    return cas_field.strip()



def parse_storage_shelf(cell_text):
    text = cell_text.replace("\n", " ").replace("\r", " ").strip()
    m = re.search(r"储存条件[：:]\s*(.*?)\s*保质期[：:]\s*(.*)", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.search(r"保质期[：:]\s*(.*?)\s*储存条件[：:]\s*(.*)", text)
    if m:
        return m.group(2).strip(), m.group(1).strip()
    if "/" in text:  # 合体格式：储存条件/保质期，如 10℃-30℃/60个月
        storage, _, shelf = text.partition("/")
        return storage.strip(), shelf.strip()
    return cell_text.strip(), "/"


def _set_cell_font(run, cn_font="微软雅黑", en_font="Times New Roman",
                   size=Pt(10), bold=False, color=None):
    run.font.size = size
    run.font.name = en_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def _format_cell(cell, text, font_size=Pt(9), bold=False,
                 align='center', cn_font="微软雅黑", en_font="Times New Roman"):
    cell.text = ""
    p = cell.paragraphs[0]
    for extra_p in cell.paragraphs[1:]:
        extra_p._element.getparent().remove(extra_p._element)
    if align == 'center':
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    elif align == 'left':
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    else:
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    pf.line_spacing = Pt(14)
    run = p.add_run(str(text))
    _set_cell_font(run, cn_font=cn_font, en_font=en_font, size=font_size, bold=bold)


def _set_col_widths(table, widths_cm):
    for i, w in enumerate(widths_cm):
        for row in table.rows:
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


# ============================================================
# 1. 解析原辅料明细
# ============================================================

def _enrich_raw_materials(docx_path, materials):
    """按原辅料明细表头映射补全：储存条件/保质期（同格分列）、开瓶次数、冻融次数、货号。
    兼容 14 列 / 16 列等不同表头布局。"""
    if not materials:
        return materials
    try:
        doc = Document(docx_path)
    except Exception:
        return materials
    for table in doc.tables:
        hdr = [re.sub(r"\s+", "", c.text) for c in table.rows[0].cells]

        def col(*keys):
            for i, h in enumerate(hdr):
                if any(k in h for k in keys):
                    return i
            return None

        code_col = col("物料编码")
        if code_col is None:
            continue
        storage_col = col("储存条件")
        shelf_col = next((i for i, h in enumerate(hdr) if "保质期" in h and "储存" not in h), None)
        open_col = col("开瓶次数")
        freeze_col = col("冻融次数")
        cas_col = col("货号")
        if storage_col is None and shelf_col is None and open_col is None and freeze_col is None and cas_col is None:
            continue
        rows = {}
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) <= code_col:
                continue
            code = cells[code_col]
            if code:
                rows[code] = cells
        for m in materials:
            cells = rows.get(m.get("物料编码"))
            if not cells:
                continue

            def cell(idx):
                return cells[idx] if idx is not None and idx < len(cells) else ""

            if storage_col is not None:
                storage, shelf = parse_storage_shelf(cell(storage_col))
                if storage and storage != "/":
                    m["储存条件"] = storage
                if shelf_col is not None:
                    sv = cell(shelf_col)
                    if sv and sv != "/":
                        shelf = sv
                m["保质期"] = shelf if shelf and shelf != "/" else "/"
            if open_col is not None:
                v = cell(open_col)
                if v:
                    m["开瓶次数"] = v
            if freeze_col is not None:
                v = cell(freeze_col)
                if v:
                    m["冻融次数"] = v
            if cas_col is not None:
                v = cell(cas_col)
                if v and not m.get("货号克隆号CAS"):
                    m["货号克隆号CAS"] = v
    return materials


def _parse_raw_material_xlsx(path, material_filter=None):
    """解析 .xlsx 格式的新增原辅料明细表（按表头映射，兼容 14/16 列等布局）"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    hdr = [re.sub(r"\s+", "", str(c or "")) for c in rows[0]]

    def col(*keys):
        for i, h in enumerate(hdr):
            if any(k in h for k in keys):
                return i
        return None

    code_col = col("物料编码")
    if code_col is None:
        raise ValueError("未找到[物料编码]表头 (xlsx)")
    name_col = col("物料名称")
    cn_col = col("物料中文名称")
    spec_col = col("规格")
    unit_col = col("主计量单位")
    cas_col = col("货号")
    brand_col = col("品牌")
    danger_col = col("危化品")
    prop_col = col("物料属性")
    storage_col = col("储存条件")
    shelf_col = next((i for i, h in enumerate(hdr) if "保质期" in h and "储存" not in h), None)
    open_col = col("开瓶次数")
    freeze_col = col("冻融次数")

    materials = []
    for ri, row in enumerate(rows[1:], 1):
        if row is None or len(row) <= code_col:
            continue
        code = str(row[code_col] or "").strip()
        if not code:
            continue
        if material_filter is not None and code not in material_filter:
            continue

        def cell(idx):
            return str(row[idx] or "").strip() if idx is not None and idx < len(row) else ""

        storage, shelf = parse_storage_shelf(cell(storage_col))
        if shelf_col is not None:
            sv = cell(shelf_col)
            if sv and sv != "/":
                shelf = sv
        materials.append({
            "序号": str(ri).zfill(2),
            "物料编码": code,
            "物料名称": cell(name_col),
            "物料中文名称": cell(cn_col),
            "规格浓度": cell(spec_col),
            "主计量单位": cell(unit_col),
            "货号克隆号CAS": cell(cas_col),
            "品牌产地": cell(brand_col),
            "是否危化品": "是" if "是" in cell(danger_col) else "否",
            "物料属性": cell(prop_col),
            "储存条件": storage,
            "保质期": shelf if shelf and shelf != "/" else "/",
            "开瓶次数": cell(open_col) or "/",
            "冻融次数": cell(freeze_col) or "/",
        })
    logger.info(f"解析到 {len(materials)} 条原辅料 (xlsx)")
    return materials


def parse_raw_material(docx_path, material_filter=None):
    """使用 word-doc-extr 技能解析原辅料明细表（零依赖，特征签名定位）"""
    if str(docx_path).lower().endswith(".xlsx"):
        return _parse_raw_material_xlsx(docx_path, material_filter)
    # 调用 extract_doc.py --by-signature
    sig_keys = "物料名称,物料属性,储存条件"
    r = subprocess.run(
        [sys.executable, _EXTRACT_SCRIPT, docx_path, "--by-signature", sig_keys],
        capture_output=True, text=True, timeout=30
    )
    try:
        result = json.loads(r.stdout)
    except:
        logger.warning("word-doc-extr解析失败，回退到python-docx")
        return _parse_raw_material_fallback(docx_path, material_filter)
    
    if result.get("status") != "success":
        logger.warning(f"word-doc-extr: {result.get('message', '')}, 回退到python-docx")
        return _parse_raw_material_fallback(docx_path, material_filter)
    
    raw_rows = result["data"]
    materials = []
    for i, row in enumerate(raw_rows):
        code = row.get("物料编码", "")
        if material_filter is not None and code not in material_filter:
            continue
        
        cas_raw = row.get("货号/克隆号或CAS号（如有）", "")
        brand_raw = row.get("品牌/产地", "")
        storage_raw = row.get("储存条件", "")
        storage_cond, shelf_life = parse_storage_shelf(storage_raw)
        
        is_danger = "是" if "是" in row.get("是否危化品", "否") else "否"
        
        material = {
            "序号": str(i + 1).zfill(2),
            "物料编码": code,
            "物料名称": row.get("物料名称", ""),
            "物料中文名称": row.get("物料中文名称", ""),
            "规格浓度": row.get("规格/浓度", ""),
            "主计量单位": row.get("主计量单位", ""),
            "货号克隆号CAS": cas_raw,
            "品牌产地": brand_raw,
            "是否危化品": is_danger,
            "物料属性": row.get("物料属性", ""),
            "储存条件": storage_cond,
            "保质期": shelf_life,
            "开瓶次数": row.get("开瓶次数", "/"),
            "冻融次数": row.get("冻融次数", "/"),
        }
        materials.append(material)
    
    logger.info(f"解析到 {len(materials)} 条原辅料")
    return _enrich_raw_materials(docx_path, materials)


def _parse_raw_material_fallback(docx_path, material_filter=None):
    """python-docx 回退方案（当 word-doc-extr 不可用时）"""
    doc = Document(docx_path)
    target = None
    for table in doc.tables:
        first = "".join(c.text.strip() for c in table.rows[0].cells)
        if "物料编码" in first and "物料名称" in first:
            target = table
            break
    if not target:
        raise ValueError("未找到[物料编码][物料名称]表头")

    materials = []
    for ri, row in enumerate(target.rows):
        if ri == 0:
            continue
        cells = [c.text.strip() for c in row.cells]
        if not cells[0]:
            continue
        code = cells[0].strip()
        if material_filter is not None and code not in material_filter:
            continue

        storage_text = cells[11] if len(cells) > 11 else ""
        storage_cond, shelf_life = parse_storage_shelf(storage_text)
        appearance_text = cells[9] if len(cells) > 9 else ""
        is_danger = "是" if "是否危化品：是" in appearance_text else "否"
        spec_flat = cells[4].replace("\n", "; ") if len(cells) > 4 else ""
        brand_flat = cells[7].replace("\n", "; ") if len(cells) > 7 else ""
        cas_flat = cells[6].replace("\n", "; ") if len(cells) > 6 else ""

        material = {
            "序号": str(ri).zfill(2),
            "物料编码": cells[0],
            "物料名称": cells[1],
            "物料中文名称": cells[2],
            "规格浓度": spec_flat,
            "主计量单位": cells[5] if len(cells) > 5 else "",
            "货号克隆号CAS": cas_flat,
            "品牌产地": brand_flat,
            "是否危化品": is_danger,
            "物料属性": cells[10] if len(cells) > 10 else "",
            "储存条件": storage_cond,
            "保质期": shelf_life,
            "开瓶次数": cells[13] if len(cells) > 13 else "/",
            "冻融次数": cells[14] if len(cells) > 14 else "/",
        }
        materials.append(material)

    logger.info(f"解析到 {len(materials)} 条原辅料 (fallback)")
    return _enrich_raw_materials(docx_path, materials)


# ============================================================
# 2. 解析质量标准
# ============================================================

def _scan_standard_sections(docx_path):
    """按章节编号扫描标准文档（与章节标题无关）：
    第1章 = 质量指标/企业标准，第2章 = 检验方法/检验步骤。
    返回 (items, sections)；sections 形如 {"1.3": {"name": "外观", "body": "白色均匀片状固体"}}。"""
    doc = Document(docx_path)
    items = {}
    sections = {}
    sec, sub_num, name, body = None, None, None, []
    pat = re.compile(r"^([12])\.(\d+)[、．.\s]*(.*)")

    def flush():
        nonlocal name, body
        if name and body:
            bt = "\n".join(body).strip()
            if sec == "1":
                items.setdefault(name, {"标准值": "", "测试方法": "/"})
                items[name]["标准值"] = bt
                sections[f"1.{sub_num}"] = {"name": name, "body": bt}
            elif sec == "2":
                if name in items:
                    items[name]["测试方法"] = bt
                else:
                    items[name] = {"标准值": "/", "测试方法": bt}
        name = None
        body = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        m = pat.match(text)
        if m:
            flush()
            sec, sub_num, name = m.group(1), m.group(2), m.group(3).strip()
            body = []
        elif name:
            body.append(text)
    flush()
    if items:
        return items, sections
    return _scan_keyword_sections(doc)


def _strip_heading_number(text):
    return re.sub(r"^\s*[（(]?\d+[）)、.．:：]?\s*", "", text).strip()


def _norm_item_name(name):
    n = re.sub(r"\s+", "", name)
    for suffix in ("检查", "试验", "测定", "项目"):
        if n.endswith(suffix) and len(n) > len(suffix):
            return n[: -len(suffix)]
    return n


def _scan_keyword_sections(doc):
    """标题式结构：第1章='企业标准/质量指标'，第2章='检验步骤/检验方法'。
    条目格式为 '名称：正文'。"""
    items = {}
    sections = {}
    sec = 0
    ignore_prefixes = (
        "目的", "适用范围", "责任部门", "引用标准", "检验规则", "判定",
        "抽样", "附录", "术语", "相关记录", "检验工具", "依据", "参考", "样品", "表",
    )

    def is_header(text, kind):
        t = _strip_heading_number(text)
        if kind == 1:
            return t.startswith("企业标准") or t.startswith("质量指标")
        return t.startswith("检验步骤") or t.startswith("检验方法")

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if is_header(text, 1):
            sec = 1
            continue
        if is_header(text, 2):
            sec = 2
            continue
        if sec == 0:
            continue
        if any(text.startswith(pfx) for pfx in ignore_prefixes):
            continue
        if "：" not in text and ":" not in text:
            continue
        name, _, val = text.partition("：") if "：" in text else text.partition(":")
        name = name.strip()
        val = val.strip()
        if not name or not val:
            continue
        if sec == 1:
            if name not in items:
                items[name] = {"标准值": val, "测试方法": "/"}
                idx = len([k for k in sections if k.startswith("1.")]) + 1
                sections[f"1.{idx}"] = {"name": name, "body": val}
        else:
            match = None
            n2 = _norm_item_name(name)
            for k in list(items):
                n1 = _norm_item_name(k)
                if n1 == n2 or n2 in n1 or n1 in n2:
                    match = k
                    break
            if match:
                items[match]["测试方法"] = val
            else:
                items[name] = {"标准值": "/", "测试方法": val}
    return items, sections


def _resolve_quality_ref(text, sections):
    """把"应符合1.X的要求 / 符合1.X的要求"替换为对应章节正文"""
    if not text or str(text).strip() in ("", "/"):
        return text
    text = str(text)
    for ref in pat_ref.findall(text):
        ref_key = ref if isinstance(ref, str) else ref[0]
        section = sections.get(ref_key)
        if not section:
            continue
        ref_body = str(section["body"]).rstrip("。")
        text = text.replace(f"应符合{ref_key}的要求", ref_body)
        text = text.replace(f"符合{ref_key}的要求", ref_body)
    return text


def parse_quality_standard(docx_path):
    """解析企业质量标准。优先 word-doc-extr；无论哪条路径都统一做章节引用解析。
    章节叫法等价：第1章=质量指标/企业标准，第2章=检验方法/检验步骤（按编号识别，与标题无关）。"""
    _, sections = _scan_standard_sections(docx_path)
    try:
        r = subprocess.run(
            [sys.executable, _EXTRACT_SCRIPT, docx_path, "--quality-standard"],
            capture_output=True, text=True, timeout=30
        )
        result = json.loads(r.stdout)
    except Exception:
        logger.warning("word-doc-extr质量标准解析失败，回退到python-docx")
        return _parse_quality_standard_fallback(docx_path, sections)

    if result.get("status") != "success":
        logger.warning(f"word-doc-extr: {result.get('message', '')}, 回退到python-docx")
        return _parse_quality_standard_fallback(docx_path, sections)

    indicators_list = result["data"]
    indicators = {}
    for ind in indicators_list:
        name = ind.get("检验项目", "")
        if not name:
            continue
        sv = _resolve_quality_ref(ind.get("标准值", "/"), sections)
        tm = _resolve_quality_ref(ind.get("测试方法", "/"), sections)
        indicators[name] = {"标准值": sv, "测试方法": tm}

    if not indicators:
        logger.warning("word-doc-extr 解析结果为空，回退到python-docx")
        return _parse_quality_standard_fallback(docx_path, sections)

    logger.info(f"解析到 {len(indicators)} 个质量指标")
    return indicators


def _parse_quality_standard_fallback(docx_path, sections=None):
    """python-docx 回退方案（当 word-doc-extr 不可用时）"""
    items, scanned = _scan_standard_sections(docx_path)
    sections = sections or scanned
    for vals in items.values():
        for key in ("标准值", "测试方法"):
            vals[key] = _resolve_quality_ref(vals[key], sections)
    logger.info(f"解析到 {len(items)} 个质量指标 (fallback)")
    return items

def get_platform_contact(platform_name):
    df = pd.read_excel(CONTACT_XLSX, engine="openpyxl")
    row = df[df["平台"] == platform_name]
    if row.empty:
        return "请填写", "请填写", "请填写"
    r = row.iloc[0]
    phone = str(int(r["电话"])) if pd.notna(r["电话"]) else "请填写"
    return r["审核人"], phone, r["邮箱地址"]


def _resolve_party_a_company(company_text, platform_name):
    """根据钉钉AI表格人工确认的甲方公司或平台默认，返回 (公司名称, 公司地址)"""
    known = {
        "示例丙公司": ("示例丙公司", "XX省XX市XX区XX路1006号"),
        "示例甲": ("示例甲公司", "YY省YY市YY区YY路16号"),
    }
    if company_text:
        text = str(company_text).strip()
        for key, val in known.items():
            if key in text:
                return val
        logger.warning("未识别甲方公司: %s，按原文本使用", text)
        return text, "请填写"
    if platform_name == "平台C":
        return known["示例丙公司"]
    return known["示例甲"]


def get_party_a_info(platform_name, party_a_contact=None, party_a_company=None):
    """根据钉钉AI表格人工确认的甲方公司/联系人返回
    (公司名称, 公司地址, 审核人, 电话, 邮箱)。
    甲方公司优先于平台默认；甲方联系人只决定审核人/电话/邮箱，不再推断公司。"""
    company_name, company_address = _resolve_party_a_company(party_a_company, platform_name)
    reviewer, phone, email = get_platform_contact(platform_name)

    if party_a_contact:
        try:
            df = pd.read_excel(CONTACT_XLSX, engine="openpyxl")
            row = df[df["审核人"] == party_a_contact]
            if not row.empty:
                r = row.iloc[0]
                reviewer = str(r["审核人"])
                email = str(r["邮箱地址"]) if pd.notna(r["邮箱地址"]) else email
                phone = str(int(r["电话"])) if pd.notna(r["电话"]) else phone
            else:
                logger.warning("甲方联系人不在对照表中: %s，使用平台默认", party_a_contact)
        except Exception as e:
            logger.warning("读取甲方联系人失败: %s，使用平台默认", e)

    return company_name, company_address, reviewer, phone, email


# ============================================================
# 4. 占位符替换
# ============================================================

def replace_placeholder(document, placeholder, new_text):
    new_text = str(new_text)
    for p in document.paragraphs:
        for run in p.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, new_text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, new_text)
    for section in document.sections:
        for p in section.header.paragraphs:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)


# ============================================================
# 5. 重绘表格
# ============================================================




def merge_cells_vertically(table, col_idx, start_row_idx, end_row_idx):
    """
    底层 XML 操作实现真实垂直合并（避免 python-docx save bug）
    合并起始行设 restart，后续行设 continue 并彻底清空内容
    """
    for row_idx in range(start_row_idx, end_row_idx + 1):
        cell = table.cell(row_idx, col_idx)
        tcPr = cell._element.get_or_add_tcPr()
        existing_vMerge = tcPr.find(qn('w:vMerge'))
        if existing_vMerge is not None:
            tcPr.remove(existing_vMerge)
        vMerge = OxmlElement('w:vMerge')
        if row_idx == start_row_idx:
            vMerge.set(qn('w:val'), 'restart')
        else:
            vMerge.set(qn('w:val'), 'continue')
            # 彻底清空此 cell 的所有段落文本
            for p in cell._element.findall(qn('w:p')):
                cell._element.remove(p)
            # 添加一个空段落（Word 需要至少一个段落）
            from docx.oxml import parse_xml
            cell._element.append(
                parse_xml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        tcPr.append(vMerge)


def _set_table_full_width(table):
    """设置表格宽度为页面100%"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    # 移除旧的 tblW
    old_w = tblPr.find(qn('w:tblW'))
    if old_w is not None:
        tblPr.remove(old_w)
    # 设置 100% 宽度
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:type'), 'pct')
    tblW.set(qn('w:w'), '5000')
    tblPr.append(tblW)


def rebuild_material_table(doc, materials):
    """重绘原辅料明细表 - 全宽美观表格"""
    if len(doc.tables) < 3:
        logger.warning("表格不足")
        return

    old_tbl = doc.tables[2]._tbl
    parent = old_tbl.getparent()

    headers = ["序号", "物料编码", "物料名称", "物料中文名称", "规格/浓度",
               "主计量单位", "货号/克隆号或CAS号", "品牌/产地", "是否危化品",
               "物料属性", "储存条件", "保质期", "开瓶次数", "冻融次数"]

    data_rows = []
    for idx, m in enumerate(materials):
        data_rows.append([
            str(idx+1).zfill(2), m["物料编码"], m["物料名称"], m.get("物料中文名称", m["物料名称"]),
            m["规格浓度"], m["主计量单位"], m["货号克隆号CAS"], m["品牌产地"],
            m["是否危化品"], m["物料属性"], m["储存条件"], m["保质期"],
            m["开瓶次数"], m["冻融次数"],
        ])

    new_table = doc.add_table(rows=1+len(data_rows), cols=len(headers))
    new_table.style = 'Table Grid'
    parent.replace(old_tbl, new_table._tbl)

    # 表头
    for ci, h in enumerate(headers):
        _format_cell(new_table.rows[0].cells[ci], h,
                     font_size=Pt(10), bold=True, align='center')

    # 数据行
    for ri, row_data in enumerate(data_rows):
        for ci, val in enumerate(row_data):
            txt = str(val)
            is_long = len(txt) > 15
            _format_cell(new_table.rows[1+ri].cells[ci], txt,
                         font_size=Pt(9), bold=False,
                         align='left' if is_long else 'center')

    # 列宽 + 全宽
    col_widths = [0.6, 1.5, 1.5, 1.5, 1.8, 0.6, 2.0, 2.0, 0.8, 0.6, 1.0, 0.6, 0.6, 0.6]
    _set_col_widths(new_table, col_widths)
    # 表格全宽
    _set_table_full_width(new_table)
    
    logger.info(f"重绘原辅料明细表: {len(materials)} 行")


def build_multi_quality_table(doc, material_groups, supplier_name, company_name):
    """重绘产品质量标准表（8列：物料编码|货号|品名|规格|检验项目|标准值|测试方法|备注）"""
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

    # 注释行 = 倒数第3行
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

    # 签章行：甲方(4格) | 乙方(2格) | 日期(2格)
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
    try:
        sig2.cells[6].merge(sig2.cells[7])
    except:
        pass
    _set_col_widths(new_table, [1.5, 2.0, 2.5, 1.5, 2.5, 4.5, 4.5, 2.0])
    _set_table_full_width(new_table)
    logger.info("质量标准: %d 指标, %d 物料",
                row_count - sum(1 for g in material_groups if not g['indicators']),
                len(material_groups))


def _group_indicators_by_material(indicators, materials=None):
    """把带 _material 归属的指标列表分组为 build_multi_quality_table 需要的 material_groups"""
    groups, by_code = [], {}
    mat_by_code = {m.get("物料编码"): m for m in (materials or [])}
    for ind in indicators:
        mat = ind.get("_material") or mat_by_code.get(ind.get("物料编码", ""), {})
        code = mat.get("物料编码") or ind.get("物料编码", "")
        if not code:
            continue
        if code not in by_code:
            by_code[code] = {
                "code": extract_part_number(mat.get("货号克隆号CAS", "")) or ind.get("货号", ""),
                "material_code": code,
                "name": mat.get("物料名称") or mat.get("物料中文名称") or ind.get("物料名称", "") or code,
                "spec": extract_spec(mat.get("规格浓度", "") or ind.get("规格", "")),
                "indicators": [],
            }
            groups.append(by_code[code])
        by_code[code]["indicators"].append({
            "检验项目": ind.get("检验项目", ""),
            "标准值": ind.get("标准值", "/"),
            "测试方法": ind.get("测试方法", "/"),
        })
    return groups


def rebuild_quality_table(doc, indicators, material_code="", material_name="", spec="", supplier_name="", company_name="示例甲公司"):
    """重绘产品质量标准表 - 全宽美观表格
    支持多物料：传入 materials_with_inds=[(mat, inds), ...] 可为每个物料独立渲染
    """
    if len(doc.tables) < 4:
        logger.warning("表格不足")
        return

    old_tbl = doc.tables[3]._tbl
    parent = old_tbl.getparent()

    headers = ["货号", "中/英文品名", "规格", "检验项目", "标准值", "测试方法", "备注"]
    note_text = "以上产品需满足：包装牢固、严密；标签及各种标志内容准确，牢固、端正、完整、清洁。"
    n = len(indicators)
    sig_data = [
        ["甲方", "", "", "", "", "", ""],
        ["（签章）", "", "", "", "乙方", "", "日期： 年 月 日"],
    ]

    # 如果一个物料有多条指标，前三列垂直合并；多个物料时每个物料独立
    total_rows = 1 + n + 1 + 2
    new_table = doc.add_table(rows=total_rows, cols=7)
    new_table.style = 'Table Grid'
    parent.replace(old_tbl, new_table._tbl)

    # 表头
    for ci, h in enumerate(headers):
        _format_cell(new_table.rows[0].cells[ci], h,
                     font_size=Pt(10.5), bold=True, align='center')

    # 指标数据行
    for ri, ind in enumerate(indicators):
        row = new_table.rows[1+ri]
        mat = ind.get("_material", {})
        is_first_of_mat = ind.get("_first_of_mat", False)
        if is_first_of_mat:
            part_no = extract_part_number(mat.get("货号克隆号CAS", ""))
            _format_cell(row.cells[0], part_no, font_size=Pt(9), align='center')
            _format_cell(row.cells[1], mat.get("物料名称", ""), font_size=Pt(9), align='center')
            _format_cell(row.cells[2], extract_spec(mat.get("规格浓度", "")), font_size=Pt(9), align='center')
        _format_cell(row.cells[3], ind["检验项目"], font_size=Pt(9),
                     align='left' if len(ind["检验项目"]) > 8 else 'center')
        _format_cell(row.cells[4], ind["标准值"], font_size=Pt(9), align='left')
        _format_cell(row.cells[5], ind["测试方法"], font_size=Pt(9), align='left')
        _format_cell(row.cells[6], "/", font_size=Pt(9), align='center')
    
    # 按物料分组垂直合并前三列：对每组物料内垂直合并
    # 先找出每个物料的起始行和条数
    mat_spans = []
    for ri, ind in enumerate(indicators):
        if ind.get("_first_of_mat", False):
            mat_spans.append({"start": ri, "code": ind.get("_material", {}).get("物料编码", "")})
    # 给每个span计算结束行
    for i in range(len(mat_spans)):
        if i + 1 < len(mat_spans):
            mat_spans[i]["end"] = mat_spans[i+1]["start"] - 1
        else:
            mat_spans[i]["end"] = n - 1
    # 对每个物料单独垂直合并
    for ms in mat_spans:
        if ms["end"] > ms["start"]:
            for col_idx in range(3):
                merge_cells_vertically(new_table, col_idx, ms["start"] + 1, ms["end"] + 1)

    # 注释行（整行合并成1个单元格）
    note_row = new_table.rows[1+n]
    note_cell = note_row.cells[0]
    for mc in range(1, 7):
        try:
            note_cell.merge(note_row.cells[mc])
            # 合并后清空被合并cell的文本
            if mc < len(note_row.cells):
                note_row.cells[mc].text = ""
        except:
            break
    _format_cell(note_cell, note_text, font_size=Pt(8.5), align='left')

    # 签章行：甲方(2.5/7) | 乙方(2.5/7) | 日期(2/7)
    # 用 7 列 grid 实现：甲方合并 0-2(3格), 乙方合并 3-4(2格), 日期合并 5-6(2格)
    sig1 = new_table.rows[1+n+1]
    sig2 = new_table.rows[1+n+2]
    
    # == 第1行 ==
    # 甲方：cells[0-2] 合并
    sig1.cells[0].merge(sig1.cells[1])
    sig1.cells[0].merge(sig1.cells[2])
    _format_cell(sig1.cells[0], f"甲方：{company_name}", font_size=Pt(10), bold=True, align='center')
    
    # 乙方：cells[3-4] 合并
    sig1.cells[3].merge(sig1.cells[4])
    _format_cell(sig1.cells[3], f"乙方：{supplier_name}", font_size=Pt(10), bold=True, align='center')
    
    # 日期：cells[5-6] 合并（占2格，比原来宽一倍）
    sig1.cells[5].merge(sig1.cells[6])
    _format_cell(sig1.cells[5], "日期： 年 月 日", font_size=Pt(9), align='center')
    
    # == 第2行 ==
    # 甲方签章
    sig2.cells[0].merge(sig2.cells[1])
    sig2.cells[0].merge(sig2.cells[2])
    _format_cell(sig2.cells[0], "（盖章）", font_size=Pt(10), bold=True, align='center')
    
    # 乙方签章
    sig2.cells[3].merge(sig2.cells[4])
    _format_cell(sig2.cells[3], "（盖章）", font_size=Pt(10), bold=True, align='center')
    
    # 日期 Row9 留空（不填内容，不与上垂直合并）
    try:
        sig2.cells[5].merge(sig2.cells[6])
    except:
        pass

    # 列宽（7列: 货号/品名/规格/检验项目/标准值/测试方法/备注）
    # 签章行比例：甲方2/5, 空白1/5, 乙方2/5 → 对应 cols 0-1(2份) + 2-3(1份) + 4-5(2份)
    col_widths = [2.0, 3.0, 2.0, 2.5, 5.0, 5.0, 2.5]
    _set_col_widths(new_table, col_widths)
    _set_table_full_width(new_table)
    
    logger.info(f"重绘质量标准表: {n} 个指标")


def append_materials_to_table(doc, new_materials):
    """追加新物料到原辅料明细表（追加到末尾）"""
    if len(doc.tables) < 3:
        logger.warning("文档表格不足，无法追加")
        return
    table = doc.tables[2]
    headers_count = len(table.rows[0].cells)
    existing_data_rows = sum(1 for r in table.rows if r.cells[0].text.strip().isdigit())
    start_num = existing_data_rows + 1
    for mi, mat in enumerate(new_materials):
        row = table.add_row()
        data = [
            str(start_num + mi).zfill(2),
            mat["物料名称"],
            mat["规格浓度"], mat["主计量单位"], mat["货号克隆号CAS"],
            mat["品牌产地"], mat["是否危化品"], mat["物料属性"],
            mat["储存条件"], mat["保质期"], mat["开瓶次数"], mat["冻融次数"],
        ]
        for ci, val in enumerate(data[:headers_count]):
            _format_cell(row.cells[ci], val, font_size=Pt(9), bold=False,
                         align="left" if len(str(val)) > 15 else "center")
    logger.info(f"追加 {len(new_materials)} 行原辅料")


def _duplicate_row(table, ref_row_idx):
    """在指定行前复制一行新行"""
    from lxml import etree
    ref_row = table.rows[ref_row_idx]
    new_tr = etree.SubElement(table._tbl, qn('w:tr'))
    # 复制单元格结构
    for cell in ref_row.cells:
        tc = etree.SubElement(new_tr, qn('w:tc'))
        tc_p = etree.SubElement(tc, qn('w:p'))
        tc_r = etree.SubElement(tc_p, qn('w:r'))
        tc_t = etree.SubElement(tc_r, qn('w:t'))
        tc_t.text = ""
    ref_row._tr.addprevious(new_tr)
    # 重新解析表格获取新行对象
    new_row = table.rows[ref_row_idx]
    return new_row


def append_indicators_to_table(doc, new_indicators, materials, supplier_name, company_name):
    """追加新质量标准（插入到备注行之前）"""
    if len(doc.tables) < 4:
        logger.warning("文档表格不足，无法追加")
        return
    
    table = doc.tables[3]
    # 找到备注行
    note_row_idx = -1
    for ri, row in enumerate(table.rows):
        txt = "".join(c.text.strip() for c in row.cells)
        if "以上产品需满足" in txt:
            note_row_idx = ri
            break
    if note_row_idx < 0:
        note_row_idx = len(table.rows)
    
    n = len(new_indicators)
    
    for idx, ind in enumerate(new_indicators):
        mat = ind.get("_material", materials[0] if materials else {})
        insert_pos = note_row_idx + idx
        new_tr = _duplicate_row(table, min(insert_pos, len(table.rows)-1))
        data = [
            extract_part_number(mat.get("货号克隆号CAS", "")) if idx == 0 else "",
            mat.get("物料名称", "") if idx == 0 else "",
            extract_spec(mat.get("规格浓度", "")) if idx == 0 else "",
            ind["检验项目"],
            ind["标准值"],
            ind["测试方法"],
            "/",
        ]
        for ci, val in enumerate(data):
            _format_cell(new_tr.cells[ci], val, font_size=Pt(9), bold=False,
                         align="left" if len(str(val)) > 15 else "center")
        logger.info(f"追加 {n} 个质量指标")


def _format_cell(cell, text, font_size=Pt(9), bold=False, align="center"):
    """格式化单元格"""
    cell.text = ""
    p = cell.paragraphs[0]
    for extra_p in cell.paragraphs[1:]:
        extra_p._element.getparent().remove(extra_p._element)
    if align == "center":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    elif align == "left":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    pf.line_spacing = Pt(14)
    run = p.add_run(str(text))
    _set_cell_font(run, size=font_size, bold=bold)
def _delete_table(doc, table_index):
    """从文档中删除指定索引的表格"""
    tables = doc.tables
    if table_index < len(tables):
        tbl = tables[table_index]._tbl
        tbl.getparent().remove(tbl)


def generate_qa_agreement(
    raw_material_docx,
    quality_std_docx,  # 单个文件路径 或 文件路径列表
    template_docx,
    supplier_name="待定供应商",
    platform_name="",
    oa_serial="QA-DOC-2026-XXXX",
    output_path=None,
    material_filter=None,  # 按物料编码过滤，如 ["9000000001"]
    incremental=False,  # True=追加新增物料，False=重新生成
    party_a_contact=None,  # 钉钉AI表格人工确认的甲方联系人（优先于平台默认）
    party_a_company=None,  # 钉钉AI表格人工确认的甲方公司（优先于平台默认）
    party_b_info=None,  # 乙方（供应商）联系人/电话/邮箱等，回填协议第二页
):
    logger.info("读取模板...")
    
    # 增量模式：如果已存在协议文件，打开它并追加新物料
    doc = None
    if incremental and output_path and os.path.exists(output_path):
        logger.info(f"增量模式: 打开已有协议 {output_path}")
        doc = Document(output_path)
    else:
        doc = Document(template_docx)
        incremental = False  # 没有已有文件时转为全新生成

    logger.info("解析原辅料明细表...")
    materials = parse_raw_material(raw_material_docx, material_filter=material_filter)

    logger.info("解析企业质量标准...")
    # 支持传入多个质量标准文件
    std_files = quality_std_docx if isinstance(quality_std_docx, list) else [quality_std_docx]
    
    # 将质量标准文件匹配到对应物料：通过文件名中的物料编码
    def _match_std_to_material(std_path, mats):
        """匹配质量标准文件到物料：编码>货号>物料名称（排除其他申请人的文件）"""
        base_fn = os.path.basename(std_path).lower().replace("（","(").replace("）",")")
        # 排除其他申请人的标准文件（避免跨邮件匹配）
        _exclude_patterns = ["张三", "李四", "王五"]
        for ex in _exclude_patterns:
            if ex in os.path.basename(std_path):
                return None
        for m in mats:
            code = m["物料编码"]
            if code in std_path:
                part_no = extract_part_number(m.get("货号克隆号CAS", ""))
                logger.info(f"    ✅ 编码匹配: {code} → {os.path.basename(std_path)}")
                return m
        for m in mats:
            part_no = extract_part_number(m.get("货号克隆号CAS", ""))
            if part_no and part_no in std_path:
                logger.info(f"    ✅ 货号匹配: {part_no} → {os.path.basename(std_path)}")
                return m
        # 按物料名称模糊匹配
        for m in mats:
            cand_names = [m.get("物料名称", ""), m.get("物料中文名称", "")]
            for cand in cand_names:
                mat_name = str(cand or "").lower().replace("（", "(").replace("）", ")").strip()
                norm_name = re.sub(r"\s+", "", mat_name).replace("β", "beta")
                key = std_filename_key(base_fn)
                if mat_name and key and len(key) >= 2 and (key in norm_name or norm_name in key):
                    logger.info(f"    ✅ 名称匹配: {m['物料编码']} {mat_name[:30]} → {os.path.basename(std_path)}")
                    return m
        logger.warning(f"    ⚠️ 未匹配到物料: {os.path.basename(std_path)}")
        return None
    
    # 每个物料解析自己的质量标准，保留物料归属信息
    indicators = []
    for sf in std_files:
        mat = _match_std_to_material(sf, materials)
        if mat is None:
            continue  # 跳过未匹配的文件（如其他申请人的标准文件）
        inds = parse_quality_standard(sf)
        # 统一格式：转为list of dicts
        if isinstance(inds, dict):
            inds = [{"检验项目": k, **v} for k, v in inds.items()]
        for i, ind in enumerate(inds):
            # 去重：相同物料内去重（same material + same inspection item = skip）
            dup = False
            for existing in indicators:
                if existing.get("_material", {}).get("物料编码") == (mat or {}).get("物料编码") and existing["检验项目"] == ind["检验项目"]:
                    dup = True
                    break
            if not dup:
                ind["_material"] = mat or (materials[0] if materials else {})
                ind["_first_of_mat"] = (i == 0)
                indicators.append(ind)
    
    # 补上首行标记：确保每组物料的第一个指标被标为_first_of_mat
    seen_codes = set()
    for ind in indicators:
        code = ind["_material"].get("物料编码", "")
        if code not in seen_codes:
            ind["_first_of_mat"] = True
            seen_codes.add(code)
    
    logger.info(f"解析后共 {len(indicators)} 个质量指标（{len(seen_codes)} 个物料）")

    if not platform_name:
        fname = os.path.basename(raw_material_docx)
        parts = fname.split("-")
        if len(parts) >= 4:
            platform_name = parts[1]
            logger.info(f"平台: {platform_name}")

    # 甲方公司/联系人：优先采用钉钉AI表格人工确认的“甲方公司/甲方联系人”，否则按平台默认
    company_name, company_address, reviewer, phone, email = get_party_a_info(
        platform_name, party_a_contact, party_a_company
    )
    logger.info(f"甲方: {company_name} | 审核人: {reviewer}")

    today = date.today().strftime("%Y年%m月%d日")

    # 替换甲方公司名称（Table 0: 买方）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "买方：示例甲公司" in cell.text:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if "示例甲公司" in run.text:
                                run.text = run.text.replace("示例甲公司", company_name)
                if "【甲方】示例甲公司" in cell.text:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if "示例甲公司" in run.text:
                                run.text = run.text.replace("示例甲公司", company_name)
                if "住所：YY省YY市YY区YY路16号" in cell.text:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if "YY省YY市YY区YY路16号" in run.text:
                                run.text = run.text.replace("YY省YY市YY区YY路16号", company_address)

    replace_placeholder(doc, "{1}", supplier_name)
    replace_placeholder(doc, "{2}", reviewer)
    replace_placeholder(doc, "{3}", phone)
    replace_placeholder(doc, "{4}", email)
    replace_placeholder(doc, "{5}", today)
    replace_placeholder(doc, "{6}", normalize_oa_serial(oa_serial))
    fill_party_b_info(doc, party_b_info)

    # 增量模式：合并新旧物料后整体重绘
    if incremental:
        logger.info("增量模式: 合并新旧物料后整体重绘...")
        existing_codes = set()
        if len(doc.tables) >= 3:
            for row in doc.tables[2].rows[1:]:
                txt = row.cells[1].text.strip() if len(row.cells) > 1 else ""
                if txt.isdigit():
                    existing_codes.add(txt)
        logger.info(f"已有物料: {existing_codes}")
        new_filtered = [m for m in materials if m["物料编码"] not in existing_codes]
        if not new_filtered:
            logger.info("所有物料已存在，无需追加")
            if output_path:
                doc.save(output_path)
                return output_path
            return ""
        logger.info(f"新增物料: {[m['物料编码'] for m in new_filtered]}")
        old_materials = []
        if len(doc.tables) >= 3:
            for row in doc.tables[2].rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                txt = cells[1] if len(cells) > 1 else ""
                if txt.isdigit() and txt not in {m["物料编码"] for m in new_filtered}:
                    old_materials.append({
                        "序号": cells[0] if cells else "",
                        "物料编码": txt,
                        "物料名称": cells[2] if len(cells) > 2 else txt,
                        "物料中文名称": cells[3] if len(cells) > 3 else (cells[2] if len(cells) > 2 else txt),
                        "规格浓度": cells[4] if len(cells) > 4 else "",
                        "主计量单位": cells[5] if len(cells) > 5 else "",
                        "货号克隆号CAS": cells[6] if len(cells) > 6 else "",
                        "品牌产地": cells[7] if len(cells) > 7 else "",
                        "是否危化品": cells[8] if len(cells) > 8 else "否",
                        "物料属性": cells[9] if len(cells) > 9 else "",
                        "储存条件": cells[10] if len(cells) > 10 else "",
                        "保质期": cells[11] if len(cells) > 11 else "",
                        "开瓶次数": cells[12] if len(cells) > 12 else "/",
                        "冻融次数": cells[13] if len(cells) > 13 else "/",
                    })
        all_materials = old_materials + new_filtered
        old_indicators = []
        if len(doc.tables) >= 4:
            for row in doc.tables[3].rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if not cells or not cells[0].isdigit():
                    continue
                if len(cells) >= 8:
                    name_idx, pi, sv, tm = 2, 4, 5, 6
                elif len(cells) >= 7:
                    name_idx, pi, sv, tm = 1, 3, 4, 5
                else:
                    continue
                old_indicators.append({
                    "物料编码": cells[0],
                    "物料名称": cells[name_idx],
                    "检验项目": cells[pi],
                    "标准值": cells[sv],
                    "测试方法": cells[tm],
                })
        all_indicators = indicators + [ind for ind in old_indicators if ind["检验项目"] not in {i["检验项目"] for i in indicators}]
        logger.info(f"合并后: {len(all_materials)} 个物料, {len(all_indicators)} 个质量指标")
        materials = all_materials
        indicators = all_indicators
        rebuild_material_table(doc, materials)
        groups = _group_indicators_by_material(indicators, materials)
        build_multi_quality_table(doc, groups, supplier_name, company_name)
    else:
        rebuild_material_table(doc, materials)
        groups = _group_indicators_by_material(indicators, materials)
        if groups:
            build_multi_quality_table(doc, groups, supplier_name, company_name)
    if output_path is None:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        tag = "平台C" if "示例丙公司" in company_name else "生物"
        output_path = os.path.join(OUTPUT_DIR, build_output_filename(oa_serial, tag, supplier_name))

    doc.save(output_path)
    logger.info(f"✅ 已生成: {output_path}")
    return output_path


if __name__ == "__main__":
    logger.info("=" * 50)
    logger.info("测试运行 v3 重绘策略")
    logger.info("=" * 50)

    result = generate_qa_agreement(
        raw_material_docx=TEST_RAW_MATERIAL,
        quality_std_docx=TEST_QUALITY_STD,
        template_docx=TEMPLATE_DOCX,
        supplier_name="测试供应商（供应商D）",
        platform_name="平台G",
        oa_serial="QA-DOC-2026-0001",
    )
    print(f"\n输出: {result}")
