#!/usr/bin/env python3
"""
Word Document Extractor - Zero Third-Party Dependencies
Uses ONLY Python stdlib: zipfile + xml.etree.ElementTree.

Features:
  1. Full extraction (paragraphs, tables, images, metadata)
  2. Signature-based table extraction with smart field splitting
  3. Quality standard document parsing (1.x indicators + 2.x test methods)
  4. Superscript/subscript preservation as Unicode characters
"""

import sys
import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/'
            'wordprocessingDrawing'),
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
}

def tag(name):
    return f'{{{NS["w"]}}}{name}'

def find_recursive(parent, *tags):
    el = parent
    for t in tags:
        el = el.find(tag(t))
        if el is None:
            return None
    return el

def findall_recursive(parent, tag_name):
    return parent.findall(f'.//{tag(tag_name)}')

_SUPERSCRIPT_MAP = {
    '0': '\u2070', '1': '\u00B9', '2': '\u00B2', '3': '\u00B3', '4': '\u2074',
    '5': '\u2075', '6': '\u2076', '7': '\u2077', '8': '\u2078', '9': '\u2079',
    '+': '\u207A', '-': '\u207B', '=': '\u207C', '(': '\u207D', ')': '\u207E',
    'n': '\u207F', 'i': '\u2071', 'l': '\u02B0',
}
_SUBSCRIPT_MAP = {
    '0': '\u2080', '1': '\u2081', '2': '\u2082', '3': '\u2083', '4': '\u2084',
    '5': '\u2085', '6': '\u2086', '7': '\u2087', '8': '\u2088', '9': '\u2089',
    '+': '\u208A', '-': '\u208B', '=': '\u208C', '(': '\u208D', ')': '\u208E',
    'a': '\u2090', 'e': '\u2091', 'o': '\u2092', 'x': '\u2093', 'h': '\u2095',
    'k': '\u2096', 'l': '\u2097', 'm': '\u2098', 'n': '\u2099', 'p': '\u209A',
    's': '\u209B', 't': '\u209C',
}

def _convert_super_sub(text, vert_align_val):
    if not text:
        return text
    mapping = _SUPERSCRIPT_MAP if vert_align_val == 'superscript' else _SUBSCRIPT_MAP
    result = []
    for ch in text:
        result.append(mapping.get(ch, ch))
    return ''.join(result)

def get_text_preserve_format(parent):
    """Concatenate all w:t text nodes, preserving super/subscript as Unicode."""
    parts = []
    for r_el in parent.iter(tag('r')):
        rPr = r_el.find(tag('rPr'))
        vert_align = None
        if rPr is not None:
            va = rPr.find(tag('vertAlign'))
            if va is not None:
                vert_align = va.get(tag('val'))
        for t in r_el.findall(tag('t')):
            if t.text:
                if vert_align in ('superscript', 'subscript'):
                    parts.append(_convert_super_sub(t.text, vert_align))
                else:
                    parts.append(t.text)
    return ''.join(parts)

def get_text(parent):
    """Concatenate all w:t text nodes. Uses format-preserving version."""
    return get_text_preserve_format(parent)

def attr_val(el, attr_local):
    return el.get(tag(attr_local)) if el is not None else None

# === FIELD SPLITTING ===

def split_storage_shelf(text):
    """Split combined storage/shelf cell: "室温储存/36个月" -> ("室温储存", "36个月")"""
    if not text:
        return "", "/"
    text = text.replace("\n", " ").replace("\r", " ").strip()
    # Try labeled format
    m = re.search(r'储存条件[：:]\s*(.*?)\s*保质期[：:]\s*(.*)', text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # Try only 保质期 label
    m = re.search(r'(.+?)保质期[：:]\s*(.*)', text)
    if m:
        cond = m.group(1).strip().rstrip('，,；;')
        cond = re.sub(r'^储存条件[：:]\s*', '', cond)
        return cond, m.group(2).strip()
    # Fall back: split on last "/"
    parts = text.rsplit("/", 1)
    if len(parts) == 2:
        left, right = parts[0].strip(), parts[1].strip()
        if re.match(r'^[\d≤≤]', right) or '个月' in right or '年' in right:
            return left, right
        return text, "/"
    return text, "/"

def split_appearance_danger(text):
    """Split combined appearance/danger: "白色粉末，否" -> ("白色粉末", "否")"""
    if not text:
        return "", "否"
    text = re.sub(r"[\n\r]+", " ", text).strip()
    text = re.sub(r"\s+", " ", text)
    # Look for explicit label
    m = re.search(r'(.+?)是否危化品[：:]\s*(是|否)', text)
    if m:
        return m.group(1).strip().rstrip('，,；;、 '), m.group(2)
    # Find last 是/否 at end
    m = re.search(r'^(.+?)[，,；;、\s]*(是|否)\s*$', text)
    if m:
        return m.group(1).strip().rstrip('，,；;、 '), m.group(2)
    if text.endswith('是'):
        return text[:-1].strip().rstrip('，,；;、 '), '是'
    if text.endswith('否'):
        return text[:-1].strip().rstrip('，,；;、 '), '否'
    return text, "否"
# === DOCX PARSER ===

class DocxParser:
    def __init__(self, path):
        self.path = Path(path)
        self.zf = zipfile.ZipFile(path)
        self.xml = self._read_xml('word/document.xml')
        self.root = self.xml.getroot()
        self.body = self.root.find(tag('body'))
        self.styles_xml = self._read_xml('word/styles.xml')

    def _read_xml(self, inner_path):
        try:
            data = self.zf.read(inner_path)
            return ET.parse(io_bytes(data)) if isinstance(data, bytes) else ET.fromstring(data)
        except (KeyError, ET.ParseError):
            return ET.fromstring('<root/>')

    def _ns_text(self, xmlns, elem):
        return elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

    def extract_metadata(self):
        meta = {}
        try:
            cp_xml = self.zf.read('docProps/core.xml')
            cp_root = ET.fromstring(cp_xml)
            for prefix, uri in [('cp', NS['cp']), ('dc', NS['dc']), ('dcterms', NS['dcterms'])]:
                for child in cp_root.iter():
                    if child.tag.startswith(f'{{{uri}'):
                        local = child.tag.split('}')[-1]
                        val = child.text or ''
                        if val.strip():
                            meta[local] = val.strip()
        except (KeyError, ET.ParseError):
            pass
        return meta

    def _parse_sect_pr(self, sect_pr):
        info = {}
        if sect_pr is None:
            return info
        pg_sz = sect_pr.find(tag('pgSz'))
        if pg_sz is not None:
            info['page_width'] = pg_sz.get(tag('w'), '')
            info['page_height'] = pg_sz.get(tag('h'), '')
            info['orientation'] = pg_sz.get(tag('orient'), '')
        pg_mar = sect_pr.find(tag('pgMar'))
        if pg_mar is not None:
            for side in ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'):
                info[f'margin_{side}'] = pg_mar.get(tag(side), '')
        return info

    def _get_style_name(self, style_id):
        if not self.styles_xml or not style_id:
            return style_id
        try:
            root = self.styles_xml if hasattr(self.styles_xml, "find") else ET.fromstring(self.styles_xml)
        except Exception:
            return style_id
        for sty in root.iter(tag('style')):
            if sty.get(tag('styleId'), '') == style_id:
                name_el = sty.find(tag('name'))
                if name_el is not None:
                    return name_el.get(tag('val'), style_id)
        return style_id

    def parse_paragraph(self, p_el):
        pPr = p_el.find(tag('pPr'))
        runs = p_el.findall(tag('r'))
        style_id = None
        is_heading = False
        heading_level = 0
        if pPr is not None:
            p_style = pPr.find(tag('pStyle'))
            if p_style is not None:
                style_id = p_style.get(tag('val'), '')
                if style_id:
                    m = re.match(r'^(?:Heading|heading|TOC)\s*(\d+)$', style_id)
                    if m:
                        is_heading = True
                        heading_level = int(m.group(1))
                    elif style_id.startswith('Heading') or style_id.startswith('heading'):
                        is_heading = True
                        m2 = re.search(r'(\d+)', style_id)
                        if m2:
                            heading_level = int(m2.group(1))
                        else:
                            heading_level = 1

        alignment = None
        if pPr is not None:
            jc = pPr.find(tag('jc'))
            if jc is not None:
                alignment = jc.get(tag('val'), '')

        run_list = []
        for r_el in runs:
            rPr = r_el.find(tag('rPr'))
            run_text = get_text(r_el)
            vert_align = None
            if rPr is not None:
                va = rPr.find(tag('vertAlign'))
                if va is not None:
                    vert_align = va.get(tag('val'))
            run = {
                'text': run_text,
                'bold': rPr is not None and rPr.find(tag('b')) is not None,
                'italic': rPr is not None and rPr.find(tag('i')) is not None,
                'underline': rPr is not None and rPr.find(tag('u')) is not None,
            }
            if vert_align:
                run['vert_align'] = vert_align
            if rPr:
                r_fonts = rPr.find(tag('rFonts'))
                if r_fonts is not None:
                    run['font'] = r_fonts.get(tag('ascii'), '')
                sz = rPr.find(tag('sz'))
                if sz is not None:
                    run['font_size_half_pt'] = sz.get(tag('val'), '')
                color = rPr.find(tag('color'))
                if color is not None:
                    run['color'] = color.get(tag('val'), '')
            run_list.append(run)

        text = get_text(p_el)
        entry = {
            'type': 'heading' if is_heading else 'paragraph',
            'text': text,
            'style_id': style_id or '',
        }
        if is_heading:
            entry['level'] = heading_level
        if alignment:
            entry['alignment'] = alignment
        entry['runs'] = run_list
        return entry

    def parse_table(self, tbl_el):
        rows = tbl_el.findall(tag('tr'))
        table_data = []
        for tr in rows:
            cells = tr.findall(tag('tc'))
            row = []
            for tc in cells:
                cell_text = get_text(tc).strip()
                inner_paras = tc.findall(tag('p'))
                para_texts = [get_text(p).strip() for p in inner_paras]
                row.append({
                    'text': cell_text,
                    'paragraphs': [t for t in para_texts if t],
                })
            table_data.append(row)
        return table_data

    def extract_images(self, output_dir):
        image_dir = Path(output_dir) / 'images'
        image_dir.mkdir(parents=True, exist_ok=True)
        images = []
        idx = 0
        for name in self.zf.namelist():
            if name.startswith('word/media/'):
                idx += 1
                data = self.zf.read(name)
                ext = Path(name).suffix.lstrip('.') or 'bin'
                filename = f'image_{idx:03d}.{ext}'
                out_path = image_dir / filename
                out_path.write_bytes(data)
                images.append({
                    'index': idx,
                    'file': str(out_path.resolve()),
                    'filename': filename,
                    'size_bytes': len(data),
                })
        return images

    def walk_body(self):
        for child in self.body:
            local = self._ns_text(NS['w'], child)
            if local == 'p':
                yield self.parse_paragraph(child)
            elif local == 'tbl':
                yield {
                    'type': 'table',
                    'rows': self.parse_table(child),
                }

    def summarize(self, content):
        paras = sum(1 for e in content if e['type'] == 'paragraph')
        headings = sum(1 for e in content if e['type'] == 'heading')
        tables = sum(1 for e in content if e['type'] == 'table')
        chars = sum(len(e.get('text', '')) for e in content if 'text' in e)
        outline = []
        for e in content:
            if e['type'] == 'heading':
                indent = '  ' * (e.get('level', 1) - 1)
                outline.append(f'{indent}{e["text"]}')
        return {
            'paragraphs': paras,
            'headings': headings,
            'tables': tables,
            'characters': chars,
            'heading_outline': '\n'.join(outline) if outline else None,
        }

    def get_paragraphs_text(self):
        """Return list of (index, text) for all non-empty paragraphs."""
        result = []
        idx = 0
        for child in self.body:
            local = self._ns_text(NS['w'], child)
            if local == 'p':
                text = get_text(child).strip()
                if text:
                    result.append((idx, text))
                idx += 1
        return result
# === SIGNATURE-BASED TABLE EXTRACTION ===

_COMBINED_FIELD_PATTERNS = {
    'split_storage_shelf': ['储存条件/保质期', '储存条件／保质期'],
    'split_appearance_danger': [
        '外观性状，是否危化品', '外观性状,是否危化品',
        '外观性状/是否危化品', '外观性状／是否危化品',
        '外观性状是否危化品',
    ],
}

def _normalize_header(header_text):
    return header_text.replace(' ', '').replace('\n', '').strip()

def _try_split_field(header_name, cell_value):
    """Check if header indicates combined field, split if so."""
    norm_header = _normalize_header(header_name)
    for pattern in _COMBINED_FIELD_PATTERNS['split_storage_shelf']:
        if pattern in norm_header:
            storage, shelf = split_storage_shelf(cell_value)
            return {'储存条件': storage, '保质期': shelf}
    for pattern in _COMBINED_FIELD_PATTERNS['split_appearance_danger']:
        if pattern in norm_header:
            appearance, danger = split_appearance_danger(cell_value)
            return {'外观性状': appearance, '是否危化品': danger}
    return None

def extract_table_by_signature(file_path, signature_keys):
    """Locate table by header signature, extract with smart field splitting."""
    if not os.path.exists(file_path):
        return {'status': 'error', 'message': f'文件不存在: {file_path}'}
    ext = os.path.splitext(file_path)[1].lower()
    if ext != '.docx':
        return {'status': 'error', 'message': f'不支持的格式: {ext}。仅支持 .docx。'}
    try:
        parser = DocxParser(file_path)
    except zipfile.BadZipFile:
        return {'status': 'error', 'message': '文件不是有效的 .docx (ZIP) 格式。'}
    target_table = None
    for child in parser.body:
        local = parser._ns_text(NS['w'], child)
        if local == 'tbl':
            rows = child.findall(tag('tr'))
            if not rows:
                continue
            header_cells = rows[0].findall(tag('tc'))
            headers_raw = []
            for tc in header_cells:
                txt = get_text(tc).replace('\n', '').replace(' ', '').strip()
                if txt:
                    headers_raw.append(txt)
            seen = set()
            headers_unique = []
            for h in headers_raw:
                if h not in seen:
                    seen.add(h)
                    headers_unique.append(h)
            header_joined = ''.join(headers_unique)
            if all(key.replace(' ', '') in header_joined for key in signature_keys):
                target_table = child
                break
    if target_table is None:
        return {'status': 'error', 'message': f'提取失败: 未找到包含 {signature_keys} 的表格。'}
    all_rows = target_table.findall(tag('tr'))
    header_cells = all_rows[0].findall(tag('tc'))
    col_map = {}
    for i, tc in enumerate(header_cells):
        clean = get_text(tc).replace('\n', '').replace(' ', '').strip()
        if clean and clean not in col_map:
            col_map[clean] = i
    extracted = []
    for tr in all_rows[1:]:
        cells = tr.findall(tag('tc'))
        first_text = get_text(cells[0]).strip() if cells else ''
        if not first_text or first_text in ('/', ''):
            continue
        row_data = {}
        for col_name, col_idx in col_map.items():
            if col_idx < len(cells):
                val = get_text(cells[col_idx]).strip()
                val = '' if val == '/' else val
            else:
                val = ''
            split_result = _try_split_field(col_name, val)
            if split_result:
                row_data.update(split_result)
            else:
                row_data[col_name] = val
        extracted.append(row_data)
    return {'status': 'success', 'data': extracted}
# === QUALITY STANDARD PARSER ===

def _save_item(item, target_list, sections_dict):
    if not item:
        return
    body = '\n'.join(item['body_lines']).strip()
    item['body'] = body
    target_list.append(item)
    if item['sec'] == '1' and not item.get('minor'):
        key = f"1.{item['major']}"
        sections_dict[key] = {'name': item['name'], 'body': body}

def _resolve_refs(text, sections_dict):
    """Resolve cross-references like "应符合1.3的要求"."""
    if not text or text == '/':
        return text
    resolved = text
    for m in re.finditer(r'(?:符合|应符合)\s*(\d+\.\d+)\s*(?:的)?要求', resolved):
        ref_key = m.group(1)
        if ref_key in sections_dict:
            ref_body = sections_dict[ref_key]['body']
            ref_body = ref_body.rstrip('。\u3002')
            resolved = resolved.replace(m.group(0), ref_body)
    return resolved

def parse_quality_standard(file_path):
    """Parse enterprise quality standard document (.docx)."""
    if not os.path.exists(file_path):
        return {'status': 'error', 'message': f'文件不存在: {file_path}'}
    ext = os.path.splitext(file_path)[1].lower()
    if ext != '.docx':
        return {'status': 'error', 'message': f'不支持的格式: {ext}'}
    try:
        parser = DocxParser(file_path)
    except zipfile.BadZipFile:
        return {'status': 'error', 'message': '文件不是有效的 .docx'}
    para_texts = parser.get_paragraphs_text()
    pat_item = re.compile(r'^([12])\.(\d+)(?:\.(\d+))?\s*(.*)')
    pat_header = re.compile(r'^(质量指标|检验方法)\s*$')
    current_section = None
    current_item = None
    items_1x = []
    items_2x = []
    sections_1x = {}
    for idx, text in para_texts:
        hdr_m = pat_header.match(text)
        if hdr_m:
            hdr_name = hdr_m.group(1)
            if current_item:
                _save_item(current_item, items_1x if current_item['sec'] == '1' else items_2x, sections_1x)
                current_item = None
            current_section = '1' if hdr_name == '质量指标' else '2'
            continue
        m = pat_item.match(text)
        if m:
            sec, major, minor, name = m.group(1), m.group(2), m.group(3), m.group(4).strip()
            if current_item:
                _save_item(current_item, items_1x if current_item['sec'] == '1' else items_2x, sections_1x)
            current_item = {'sec': sec, 'major': major, 'minor': minor, 'name': name, 'body_lines': []}
            current_section = sec
            continue
        if current_item and current_section:
            cleaned = text
            if current_item['sec'] == '1' and cleaned.strip() == '检验方法':
                _save_item(current_item, items_1x, sections_1x)
                current_item = None
                current_section = '2'
                continue
            if current_item['sec'] == '1':
                cleaned = re.sub(r'\s*检验方法\s*$', '', cleaned).strip()
                if not cleaned:
                    continue
            # Stop collecting body when we hit non-method section headers
            _stop_words = ['抽样规则', '抽样', '检验工具', '判定', '术语', '相关记录',
                           '附录', '样本量', '样本字码', 'Ac ', 'Re ']
            if current_item['sec'] == '2' and any(cleaned.startswith(w) for w in _stop_words):
                _save_item(current_item, items_2x if current_item['sec'] == '2' else items_1x, sections_1x)
                current_item = None
                current_section = None
                continue
            current_item['body_lines'].append(cleaned)
    if current_item:
        _save_item(current_item, items_1x if current_item['sec'] == '1' else items_2x, sections_1x)
    # Build indicators from 1.x
    indicators = []
    for item in items_1x:
        name = item['name']
        body = '\n'.join(item['body_lines']).strip()
        if not body:
            body = '/'
        if item.get('minor'):
            parent_key = f"1.{item['major']}"
            parent = sections_1x.get(parent_key, {})
            parent_name = parent.get('name', '')
            if parent_name and parent_name not in body:
                display_name = f"{parent_name} - {item['major']}.{item['minor']} {body[:80]}"
            else:
                display_name = f"{parent_name} {body[:80]}" if parent_name else body[:100]
            indicators.append({'检验项目': display_name.strip(), '标准值': body, '测试方法': '/'})
        else:
            item_key = f"1.{item['major']}"
            sections_1x[item_key] = {'name': name, 'body': body}
            indicators.append({'检验项目': name, '标准值': body, '测试方法': '/'})
    # Match 2.x test methods to 1.x
    name_to_indicators = {}
    for ind in indicators:
        key = ind['检验项目']
        if key not in name_to_indicators:
            name_to_indicators[key] = []
        name_to_indicators[key].append(ind)
    num_to_name = {}
    for item in items_1x:
        if not item.get('minor'):
            num_to_name[f"1.{item['major']}"] = item['name']
            num_to_name[item['major']] = item['name']
    for item in items_2x:
        item_name = item['name']
        method_body = '\n'.join(item['body_lines']).strip()
        if not method_body:
            method_body = item_name if item_name else '/'
        # Strategy 1: Direct name match
        if item_name in name_to_indicators:
            for ind in name_to_indicators[item_name]:
                if ind['测试方法'] == '/' or not ind['测试方法']:
                    ind['测试方法'] = _resolve_refs(method_body, sections_1x)
                else:
                    resolved = _resolve_refs(method_body, sections_1x)
                    if resolved not in ind['测试方法']:
                        ind['测试方法'] += '\n' + resolved
            continue
        # Strategy 2: Combined test method
        matched_any = False
        if '供方检验报告' in method_body or '供方检验' in method_body:
            for ind in indicators:
                ind_name = ind['检验项目']
                if ind_name and len(ind_name) >= 1 and ind_name in (item_name + ' ' + method_body):
                    ind['测试方法'] = method_body
                    matched_any = True
            range_m = re.search(r'符合\s*(\d+\.\d+)\s*[-~]\s*(\d+\.\d+)\s*(?:的)?要求', method_body)
            if range_m:
                start_major = int(range_m.group(1).split('.')[1])
                end_major = int(range_m.group(2).split('.')[1])
                for mi in range(start_major, end_major + 1):
                    ref_key = f"1.{mi}"
                    ref_name = num_to_name.get(str(mi)) or num_to_name.get(ref_key)
                    if ref_name and ref_name in name_to_indicators:
                        for ind in name_to_indicators[ref_name]:
                            if ind['测试方法'] == '/' or not ind['测试方法']:
                                ind['测试方法'] = method_body
                                matched_any = True
        if not matched_any:
            for ind in indicators:
                ind_name = ind['检验项目']
                if ind_name and len(ind_name) >= 1 and (ind_name in item_name or item_name in ind_name):
                    if ind['测试方法'] == '/' or not ind['测试方法']:
                        ind['测试方法'] = _resolve_refs(method_body, sections_1x)
                        matched_any = True
                    break
    # Post-process: resolve cross-refs
    for ind in indicators:
        tm = ind.get('测试方法', '/')
        if tm and tm != '/':
            ind['测试方法'] = _resolve_refs(tm, sections_1x)
        sv = ind.get('标准值', '/')
        if sv and sv != '/':
            ind['标准值'] = _resolve_refs(sv, sections_1x)
    return {'status': 'success', 'data': indicators, 'meta': {
        'file': file_path,
        'total_indicators': len(indicators),
        'indicators_with_methods': sum(1 for i in indicators if i['测试方法'] and i['测试方法'] != '/'),
    }}
# === FULL EXTRACTION ===

def full_extract(file_path, output_dir=None):
    parser = DocxParser(file_path)
    out = Path(output_dir) if output_dir else parser.path.parent
    content = list(parser.walk_body())
    metadata = parser.extract_metadata()
    sections = []
    last_sect = parser.body.find(tag('sectPr'))
    if last_sect is not None:
        sections.append(parser._parse_sect_pr(last_sect))
    images = parser.extract_images(out)
    for img in images:
        content.append({
            'type': 'image', 'index': img['index'], 'file': img['file'],
            'filename': img['filename'], 'size_bytes': img['size_bytes'],
        })
    summary = parser.summarize(content)
    return {'file': str(parser.path.resolve()), 'type': 'docx', 'metadata': metadata,
            'sections': sections, 'content': content, 'images': images, 'summary': summary}

def format_markdown(result):
    lines = []
    title = result.get('metadata', {}).get('title') or Path(result['file']).stem
    lines.append(f'# {title}\n')
    meta = result.get('metadata', {})
    if meta:
        lines.append('## 文档信息\n')
        for k, v in meta.items():
            if v:
                lines.append(f'- **{k}**: {v}')
        lines.append('')
    s = result.get('summary', {})
    lines.append(f'**统计**: {s.get("paragraphs", 0)} 段落, {s.get("headings", 0)} 标题, {s.get("tables", 0)} 表格, {s.get("characters", 0)} 字符\n')
    outline = s.get('heading_outline')
    if outline:
        lines.append('## 目录结构\n```')
        lines.append(outline)
        lines.append('```\n')
    lines.append('## 正文\n')
    for elem in result.get('content', []):
        t = elem.get('type', '')
        if t in ('paragraph', 'heading'):
            text = elem.get('text', '')
            if not text:
                continue
            if t == 'heading':
                level = elem.get('level', 1)
                lines.append(f'{"#" * (level + 1)} {text}')
            else:
                lines.append(text)
            bold_runs = [r for r in elem.get('runs', []) if r.get('bold') and r['text'].strip()]
            if bold_runs:
                hints = [r['text'][:30] for r in bold_runs[:2]]
                lines.append(f'> 格式提示: 含加粗文字 - {"、".join(hints)}')
            lines.append('')
        elif t == 'table':
            rows = elem.get('rows', [])
            if not rows:
                continue
            col_count = max((len(r) for r in rows), default=0)
            if col_count == 0:
                continue
            lines.append('### 表格\n')
            lines.append('| ' + ' | '.join('列' + str(i+1) for i in range(col_count)) + ' |')
            lines.append('| ' + ' | '.join(['---'] * col_count) + ' |')
            for row in rows:
                cells = [c.get('text', '').replace('\n', ' ') for c in row]
                while len(cells) < col_count:
                    cells.append('')
                lines.append('| ' + ' | '.join(cells[:col_count]) + ' |')
            lines.append('')
        elif t == 'image':
            f = elem.get('file', '')
            if f:
                lines.append(f'![]({f})\n')
    return '\n'.join(lines)

def io_bytes(data):
    import io
    return io.BytesIO(data)

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Word .docx extractor')
    parser.add_argument('input', help='Path to .docx file')
    parser.add_argument('--by-signature', '-s', help='Extract table by signature keys')
    parser.add_argument('--quality-standard', '-q', action='store_true', help='Parse quality standard document')
    parser.add_argument('--output-dir', '-o', help='Output directory')
    parser.add_argument('--format', '-f', choices=['json', 'markdown', 'both'], default='json')
    args = parser.parse_args()
    if args.quality_standard:
        result = parse_quality_standard(args.input)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    if args.by_signature:
        keys = [k.strip() for k in args.by_signature.split(',') if k.strip()]
        result = extract_table_by_signature(args.input, keys)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    result = full_extract(args.input, args.output_dir)
    out_base = Path(args.output_dir) if args.output_dir else Path(args.input).parent
    out_base.mkdir(parents=True, exist_ok=True)
    stem = Path(args.input).stem
    if args.format in ('json', 'both'):
        jp = out_base / f'{stem}_extracted.json'
        jp.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str))
        print(f'JSON -> {jp}')
    if args.format in ('markdown', 'both'):
        mp = out_base / f'{stem}_extracted.md'
        mp.write_text(format_markdown(result), encoding='utf-8')
        print(f'Markdown -> {mp}')
    if args.format == 'json':
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))

if __name__ == '__main__':
    main()
