#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
质量保证协议自动生成系统 v2
解析原辅料明细表和企业质量标准 -> 填充到质量保证协议模板
"""

import os
import re
import logging
from datetime import date
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Cm, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def set_cell_text(tc, text):
    """安全设置Word表格单元格文本（清空所有run后设置在第一个run）"""
    # 找到或创建paragraph
    ps = tc.findall(qn('w:p'))
    if not ps:
        p = parse_xml(f'<w:p {nsdecls("w")}/>')
        tc.append(p)
        ps = [p]
    p = ps[0]
    
    # 清空所有run的文本
    all_runs = p.findall(qn('w:r'))
    for r in all_runs:
        for t_elem in r.findall(qn('w:t')):
            t_elem.text = ""
    
    # 找到或创建run
    if not all_runs:
        run = parse_xml(f'<w:r {nsdecls("w")}><w:t {nsdecls("w")}></w:t></w:r>')
        p.append(run)
        all_runs = [run]
    
    # 在第一个run上设置文本
    t_elem = all_runs[0].find(qn('w:t'))
    if t_elem is None:
        t_elem = parse_xml(f'<w:t {nsdecls("w")}></w:t>')
        all_runs[0].append(t_elem)
    t_elem.text = text
    t_elem.set(qn('xml:space'), 'preserve')



from docx.oxml import parse_xml
import pandas as pd

# ============================================================
# 配置
# ============================================================
BASE_DIR = r".\agreement-automation"

TEST_RAW_MATERIAL = os.path.join(BASE_DIR, "新增原辅料明细表-平台G-9000000017-edta2NA-张三-test.docx")
TEST_QUALITY_STD  = os.path.join(BASE_DIR, "企业质量标准-平台G-9000000017-edta2NA-张三-test.docx")
CONTACT_XLSX      = os.path.join(BASE_DIR, "平台联系人的对应关系.xlsx")
TEMPLATE_DOCX     = os.path.join(BASE_DIR, "质量保证协议模板.docx")
OUTPUT_DIR        = os.path.join(BASE_DIR, "output")

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name)


def extract_spec(concise_text):
    """从 "规格：500 mL/瓶; 浓度：1M" 或 "规格：500 mL/瓶\n浓度：1M" 中提取规格 "500 mL/瓶" """
    # 先统一分隔符：将 \n 转换为 ;
    text = concise_text.replace("\n", ";").replace("\r", "")
    m = re.search(r"规格[：:]\s*([^;；]+)", text)
    if m:
        return m.group(1).strip()
    # 降级：返回第一个分号前的内容
    parts = concise_text.split(";")
    first = parts[0].strip()
    if "：" in first:
        return first.split("：")[-1].strip()
    return concise_text.strip()


def parse_storage_shelf(cell_text):
    text = cell_text.replace("\n", " ").replace("\r", " ")
    m = re.search(r"储存条件[：:]\s*(.*?)\s*保质期[：:]\s*(.*)", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    logger.warning(f"无法解析储存条件/保质期: {cell_text[:50]}")
    return cell_text.strip(), "/"


# ============================================================
# 1. 解析新增原辅料明细表
# ============================================================

def parse_raw_material(docx_path):
    doc = Document(docx_path)
    target_table = None
    for table in doc.tables:
        first_row_text = "".join(cell.text.strip() for cell in table.rows[0].cells)
        if "物料编码" in first_row_text and "物料名称" in first_row_text:
            target_table = table
            break
    if target_table is None:
        raise ValueError("未找到特征表头[物料编码][物料名称]，请检查文档格式")

    materials = []
    for ri, row in enumerate(target_table.rows):
        if ri == 0:
            continue
        cells = [cell.text.strip() for cell in row.cells]
        if not cells[0] or cells[0] == "":
            continue

        storage_text = cells[11] if len(cells) > 11 else ""
        storage_cond, shelf_life = parse_storage_shelf(storage_text)

        appearance_text = cells[9] if len(cells) > 9 else ""
        is_dangerous = "否"
        if "危化品" in appearance_text and "是否危化品：是" in appearance_text:
            is_dangerous = "是"
        elif "危化品" in appearance_text and "否" not in appearance_text:
            is_dangerous = "是"

        # Clean up multi-line cell values - flatten for compact display
        spec_concise = cells[4].replace("\n", "; ") if len(cells) > 4 else ""
        brand_concise = cells[7].replace("\n", "; ") if len(cells) > 7 else ""
        cas_concise = cells[6].replace("\n", "; ") if len(cells) > 6 else ""

        material = {
            "物料编码": cells[0],
            "物料名称": cells[1],
            "物料中文名称": cells[2],
            "规格浓度": spec_concise,
            "主计量单位": cells[5] if len(cells) > 5 else "",
            "货号克隆号CAS": cas_concise,
            "品牌产地": brand_concise,
            "是否危化品": is_dangerous,
            "物料属性": cells[10] if len(cells) > 10 else "",
            "储存条件": storage_cond,
            "保质期": shelf_life,
            "开瓶次数": cells[13] if len(cells) > 13 else "/",
            "冻融次数": cells[14] if len(cells) > 14 else "/",
        }
        materials.append(material)

    logger.info(f"解析到 {len(materials)} 条原辅料")
    return materials


# ============================================================
# 2. 解析企业质量标准
# ============================================================

def parse_quality_standard(docx_path):
    doc = Document(docx_path)
    items = {}
    current_section = None
    current_name = None
    current_body = []
    section_pattern = re.compile(r"^([12])\.(\d+)\s*(.*)")

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        m = section_pattern.match(text)
        if m:
            if current_name and current_body:
                body_text = "\n".join(current_body).strip()
                if current_section == "1":
                    if current_name not in items:
                        items[current_name] = {"标准值": "", "测试方法": "/"}
                    items[current_name]["标准值"] = body_text
                elif current_section == "2":
                    if current_name not in items:
                        items[current_name] = {"标准值": "/", "测试方法": ""}
                    items[current_name]["测试方法"] = body_text
            current_section = m.group(1)
            current_name = m.group(3).strip()
            current_body = []
        else:
            if current_name:
                current_body.append(text)

    if current_name and current_body:
        body_text = "\n".join(current_body).strip()
        if current_section == "1":
            if current_name not in items:
                items[current_name] = {"标准值": "", "测试方法": "/"}
            items[current_name]["标准值"] = body_text
        elif current_section == "2":
            if current_name not in items:
                items[current_name] = {"标准值": "/", "测试方法": ""}
            items[current_name]["测试方法"] = body_text

    indicators = []
    for p in doc.paragraphs:
        text = p.text.strip()
        m = section_pattern.match(text)
        if m and m.group(1) == "1":
            name = m.group(3).strip()
            if name in items:
                indicators.append({
                    "检验项目": name,
                    "标准值": items[name].get("标准值", "/"),
                    "测试方法": items[name].get("测试方法", "/"),
                })

    logger.info(f"解析到 {len(indicators)} 个质量指标")
    return indicators


# ============================================================
# 3. 读取平台联系人
# ============================================================

def get_platform_contact(platform_name):
    df = pd.read_excel(CONTACT_XLSX, engine="openpyxl")
    row = df[df["平台"] == platform_name]
    if row.empty:
        logger.warning(f"未找到平台 [{platform_name}] 的联系人信息")
        return "请填写", "请填写", "请填写"
    r = row.iloc[0]
    phone = r["电话"]
    if pd.isna(phone):
        phone = "请填写"
    else:
        phone = str(int(phone))
    return r["审核人"], phone, r["邮箱地址"]


# ============================================================
# 4. 占位符替换（Run级别，保留样式）
# ============================================================

def replace_placeholder(document, placeholder, new_text):
    """在文档所有位置替换占位符（Run级别保留样式）"""
    new_text = str(new_text)
    
    # 正文段落
    for p in document.paragraphs:
        for run in p.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, new_text)
    
    # 表格
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, new_text)
    
    # 页眉页脚
    for section in document.sections:
        for p in section.header.paragraphs:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)


# ============================================================
# 5. 设置表格样式（紧凑美观）
# ============================================================

def set_table_style(table, font_size=9):
    """设置表格紧凑样式"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>') 
    
    # 设置表格自动适应
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
    tblPr.append(tblW)
    
    # 设置表格边框为细线
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)
    
    # 设置所有单元格字体大小和段落间距
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = Pt(12)
                for run in p.runs:
                    run.font.size = Pt(font_size)
    
    # 设置表格行高为最小值
    for row in table.rows:
        tr = row._tr
        trPr = tr.trPr if tr.trPr is not None else parse_xml(f'<w:trPr {nsdecls("w")}/>')
        # 清除原有的trPr子元素
        for child in list(trPr):
            trPr.remove(child)
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="280" w:hRule="atLeast"/>'))


# ============================================================
# 6. 填充原辅料明细表（Table 2）
# ============================================================

def fill_material_table(template_doc, materials):
    target_table = None
    for table in template_doc.tables:
        first_row_text = "".join(cell.text.strip() for cell in table.rows[0].cells)
        if "物料名称" in first_row_text and "规格/浓度" in first_row_text:
            target_table = table
            break
    if target_table is None:
        raise ValueError("未在模板中找到原辅料明细表")

    # 删除模板中的示例行（如有）
    if len(target_table.rows) > 1:
        tbl = target_table._tbl
        row_to_remove = target_table.rows[1]
        tbl.remove(row_to_remove._tr)

    # 设置表格紧凑样式
    set_table_style(target_table, font_size=8)

    # 逐行插入数据
    for idx, mat in enumerate(materials):
        row = target_table.add_row()
        cells = row.cells
        cells[0].text = str(idx + 1).zfill(2)
        cells[1].text = mat["物料名称"]
        cells[2].text = mat["物料中文名称"] if mat["物料中文名称"] and mat["物料中文名称"] != "/" else mat["物料名称"]
        cells[3].text = mat["规格浓度"]
        cells[4].text = mat["主计量单位"]
        cells[5].text = mat["货号克隆号CAS"]
        cells[6].text = mat["品牌产地"]
        cells[7].text = mat["是否危化品"]
        cells[8].text = mat["物料属性"]
        cells[9].text = mat["储存条件"]
        cells[10].text = mat["保质期"]
        cells[11].text = mat["开瓶次数"]
        cells[12].text = mat["冻融次数"]

    logger.info(f"已填充 {len(materials)} 条原辅料到表1")


# ============================================================
# 7. 填充产品质量标准表（Table 3）
# ============================================================

def fill_quality_standard_table(template_doc, indicators, material_code="", material_name="", spec_raw="", template_docx_path=""):
    target_table = None
    for table in template_doc.tables:
        for ri in range(min(2, len(table.rows))):
            row_text = "".join(cell.text.strip() for cell in table.rows[ri].cells)
            if "检验项目" in row_text and "标准值" in row_text:
                target_table = table
                break
        if target_table:
            break
    if target_table is None:
        raise ValueError("未在模板中找到产品质量标准表")

    # 定位注释行和签章行
    note_row_xml = None
    sig_rows_xml = []
    for ri, row in enumerate(target_table.rows):
        row_text = "".join(cell.text.strip() for cell in row.cells)
        if "以上产品需满足" in row_text:
            note_row_xml = row._tr
        if ri >= 2 and "甲方" in row_text and "乙方" in row_text:
            sig_rows_xml.append(row._tr)

    # 删除所有数据行（保留表头Row0-1）
    tbl = target_table._tbl
    for ri in range(len(target_table.rows) - 1, 1, -1):
        row = target_table.rows[ri]
        if row._tr not in sig_rows_xml and row._tr is not note_row_xml:
            tbl.remove(row._tr)

    # 先移除注释行和签章行（暂存）
    if note_row_xml is not None:
        note_row_xml.getparent().remove(note_row_xml)
    for sig_xml in sig_rows_xml:
        sig_xml.getparent().remove(sig_xml)

    # 设置紧凑样式
    set_table_style(target_table, font_size=8)

    # 提取纯规格
    spec_clean = extract_spec(spec_raw)

    # 从原始模板文件取参考数据行结构
    # 重新读取模板获取Row2（保留原始gridSpan结构）
    ref_doc = Document(template_docx_path)
    ref_t3 = None
    for t in ref_doc.tables:
        for ri in range(min(2, len(t.rows))):
            rt = "".join(c.text.strip() for c in t.rows[ri].cells)
            if "检验项目" in rt and "标准值" in rt:
                ref_t3 = t
                break
        if ref_t3:
            break
    ref_tr_xml = deepcopy(ref_t3.rows[2]._tr) if ref_t3 else deepcopy(target_table.rows[1]._tr)
    
    # 清空参考行的所有文本，保留gridSpan结构
    for tc in ref_tr_xml.findall(qn('w:tc')):
        for p in tc.findall(qn('w:p')):
            for r in p.findall(qn('w:r')):
                for t in r.findall(qn('w:t')):
                    t.text = ""
            # 确保至少有一个run
            p_elem = p
            runs = p.findall(qn('w:r'))
            if not runs:
                run = parse_xml(f'<w:r {nsdecls("w")}><w:t {nsdecls("w")}></w:t></w:r>')
                p.append(run)
            else:
                for r in runs:
                    t_elem = r.find(qn('w:t'))
                    if t_elem is None:
                        t_elem = parse_xml(f'<w:t {nsdecls("w")}></w:t>')
                        r.append(t_elem)
                    t_elem.text = ""

    # 为每个质量指标克隆一行并填充
    for idx, ind in enumerate(indicators):
        new_tr = deepcopy(ref_tr_xml)
        
        # 获取该行的tc元素列表
        tcs = new_tr.findall(qn('w:tc'))
        
        # 填充内容到对应tc（按grid索引）
        # tc[0] = cells[0] = 货号(grid 0-1)
        # tc[2] = cells[2] = 品名(grid 2)
        # tc[3] = cells[3] = 规格(grid 3)
        # tc[4] = cells[4] = 检验项目(grid 4-5)
        # tc[6] = cells[6] = 标准值(grid 6-7)
        # tc[8] = cells[8] = 测试方法(grid 8-9)
        
        # 原模板数据行只有7个tc: [0]=货号, [1]=品名, [2]=规格, [3]=检验项目, [4]=标准值, [5]=测试方法, [6]=备注
        set_cell_text(tcs[3], ind["检验项目"])
        set_cell_text(tcs[4], ind["标准值"])
        set_cell_text(tcs[5], ind["测试方法"])
        
        if idx == 0:
            set_cell_text(tcs[0], material_code)
            set_cell_text(tcs[1], material_name)
            set_cell_text(tcs[2], spec_clean)
        
        # 插入到tbl末尾
        tbl.append(new_tr)

    # 对数据行应用垂直合并（货号/品名/规格跨行）
    data_start = 2  # Row2开始
    if len(indicators) > 0:
        for col_idx in [0, 1, 2]:  # 需要垂直合并的列: 货号(0), 品名(1), 规格(2)
            for ri in range(len(indicators)):
                row = target_table.rows[data_start + ri]
                tcs = row._tr.findall(qn('w:tc'))
                if col_idx < len(tcs):
                    tc = tcs[col_idx]
                    tcPr = tc.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = parse_xml(f'<w:tcPr {nsdecls("w")}/>')
                        tc.insert(0, tcPr)
                    # 移除旧的vMerge
                    old_vm = tcPr.find(qn('w:vMerge'))
                    if old_vm is not None:
                        tcPr.remove(old_vm)
                    # 设置vMerge
                    if ri == 0:
                        vm_elem = parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>')
                        # 设置垂直居中
                        vAlign = tcPr.find(qn('w:vAlign'))
                        if vAlign is not None:
                            tcPr.remove(vAlign)
                        va_elem = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>')
                        tcPr.append(va_elem)
                    else:
                        vm_elem = parse_xml(f'<w:vMerge {nsdecls("w")}/>')
                    tcPr.append(vm_elem)

    # 重新添加注释行和签章行
    if note_row_xml is not None:
        tbl.append(note_row_xml)
    for sig_xml in sig_rows_xml:
        tbl.append(sig_xml)

    logger.info(f"已填充 {len(indicators)} 个质量指标到表2（垂直合并货号/品名/规格）")
# ============================================================
# 8. 主流程
# ============================================================

def generate_qa_agreement(
    raw_material_docx,
    quality_std_docx,
    template_docx,
    supplier_name="待定供应商",
    platform_name="",
    oa_serial="QA-DOC-2026-XXXX",
    output_path=None,
):
    logger.info("正在读取模板...")
    doc = Document(template_docx)

    logger.info("正在解析原辅料明细表...")
    materials = parse_raw_material(raw_material_docx)

    logger.info("正在解析企业质量标准...")
    indicators = parse_quality_standard(quality_std_docx)

    # 从文件名中提取平台名称
    if not platform_name:
        fname = os.path.basename(raw_material_docx)
        parts = fname.split("-")
        if len(parts) >= 4:
            platform_name = parts[1]
            logger.info(f"从文件名提取平台: {platform_name}")
    
    if platform_name:
        reviewer, phone, email = get_platform_contact(platform_name)
        logger.info(f"平台 [{platform_name}] 审核人: {reviewer}")
    else:
        reviewer, phone, email = "请填写", "请填写", "请填写"

    today = date.today().strftime("%Y年%m月%d日")

    # ---- 4. 替换占位符 ----
    # {1} 供应商名称
    replace_placeholder(doc, "{1}", supplier_name)
    
    # {2} 联系人（仅甲方列，但replace_placeholder只替换Run中的文本，
    #    因此{2}只在甲方联系人中出现，乙方的"联系人："没有占位符）
    replace_placeholder(doc, "{2}", reviewer)
    
    # {3} 联系电话
    replace_placeholder(doc, "{3}", phone)
    
    # {4} 电子邮件
    replace_placeholder(doc, "{4}", email)
    
    # {5} 签订日期
    replace_placeholder(doc, "{5}", today)
    
    # {6} 页眉流水号
    replace_placeholder(doc, "{6}", oa_serial)

    # ---- 5. 填充表1：原辅料明细 ----
    fill_material_table(doc, materials)

    # ---- 6. 填充表2：产品质量标准 ----
    if materials:
        first_mat = materials[0]
        fill_quality_standard_table(
            doc, indicators,
            material_code=first_mat["物料编码"],
            material_name=first_mat["物料中文名称"] or first_mat["物料名称"],
            spec_raw=first_mat["规格浓度"],
            template_docx_path=template_docx,
        )

    # ---- 7. 保存 ----
    if output_path is None:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        n = len(materials)
        company = "平台C" if "平台C" in platform_name else "生物"
        safe_supplier = sanitize_filename(supplier_name)
        output_path = os.path.join(OUTPUT_DIR, f"{oa_serial}-【{company}】-{safe_supplier}-{n}.docx")

    doc.save(output_path)
    logger.info(f"✅ 质量保证协议已生成: {output_path}")
    return output_path


# ============================================================
# 9. 测试执行
# ============================================================

if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("使用测试文件运行 v2...")
    logger.info("=" * 60)

    result = generate_qa_agreement(
        raw_material_docx=TEST_RAW_MATERIAL,
        quality_std_docx=TEST_QUALITY_STD,
        template_docx=TEMPLATE_DOCX,
        supplier_name="测试供应商（供应商D）",
        platform_name="平台G",
        oa_serial="QA-DOC-2026-0001",
    )
    print(f"\n输出文件: {result}")
