#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""quality agreement email monitor"""

import subprocess, json, os, sys, time, hashlib, base64, hmac, shutil, re, zipfile, tempfile
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qa_config_loader import load_qa_config
from qa_upgrade import detect_change_type

_cfg = load_qa_config()
QA_EMAIL_DIR = _cfg.get("qa_dir", "")
STATE_FILE = _cfg.get("state_file", "")
DT_APP_KEY = _cfg.get("app_key", "")
DT_APP_SECRET = _cfg.get("app_secret", "")
DT_OPERATOR_ID = _cfg.get("operator_id", "")
DT_BASE_ID = _cfg.get("base_id", "")
DT_SHEET2_ID = _cfg.get("sheet2_id", "")
DT_WEBHOOK = _cfg.get("webhook", "")
DT_ROBOT_SECRET = _cfg.get("robot_secret", "")
BASE_RECORD_URL = _cfg.get("base_record_url", "")

if os.name != "nt":
    if not os.path.exists(QA_EMAIL_DIR):
        QA_EMAIL_DIR = "/app/email-database/QA"
    if not STATE_FILE or not os.path.exists(os.path.dirname(STATE_FILE)):
        STATE_FILE = "/app/email-database/.monitor_state.json"


def extract_archive(filepath, extract_to=None):
    """解压压缩包，返回包含 .docx/.doc/.xlsx 文件的列表"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in ('.zip', '.rar'):
        return [filepath]  # 不是压缩包，直接返回原文件
    
    if extract_to is None:
        extract_to = tempfile.mkdtemp(prefix="qa_extract_")
    
    extracted = []
    try:
        if ext == '.zip':
            with zipfile.ZipFile(filepath, 'r') as zf:
                zf.extractall(extract_to)
        elif ext == '.rar':
            try:
                subprocess.run(['unrar', 'x', '-o+', filepath, extract_to + '/'],
                             capture_output=True, timeout=30)
            except FileNotFoundError:
                try:
                    subprocess.run(['7z', 'x', filepath, '-o' + extract_to, '-y'],
                                 capture_output=True, timeout=30)
                except FileNotFoundError:
                    tar_candidates = [r'C:\Windows\System32\tar.exe', 'bsdtar']
                    extracted_ok = False
                    for tar in tar_candidates:
                        try:
                            subprocess.run([tar, '-xf', filepath, '-C', extract_to],
                                         capture_output=True, timeout=60)
                            extracted_ok = True
                            break
                        except FileNotFoundError:
                            continue
                    if not extracted_ok:
                        log("未找到unrar/7z/tar，无法解压.rar文件: " + filepath)
                        return [filepath]
        
        # 递归搜索所有docx/doc文件
        for root, dirs, files in os.walk(extract_to):
            for f in files:
                if f.endswith('.docx') or f.endswith('.doc') or f.endswith('.xlsx'):
                    extracted.append(os.path.join(root, f))
        
        log(f"解压 {os.path.basename(filepath)} → {len(extracted)} 个文档")
    except Exception as e:
        log(f"解压失败 {filepath}: {e}")
        return [filepath]
    
    return extracted if extracted else [filepath]


def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")


def run_agently(args):
    arg_str = " ".join(args)
    arg_str = arg_str.replace('"', '\\"')
    home_dir = os.path.expanduser("~")
    try:
        if os.name == 'nt':
            cmd = 'cmd.exe /c "agently-cli ' + arg_str + '"'
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30, shell=True, cwd=home_dir)
        else:
            cmd = ['agently-cli'] + args
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30, cwd=home_dir)
        if result.returncode == 0 and result.stdout:
            text = result.stdout
            brace_idx = text.find("{")
            if brace_idx >= 0:
                json_text = text[brace_idx:]
                end_idx = json_text.rfind("}")
                if end_idx >= 0:
                    json_text = json_text[:end_idx+1]
                    try:
                        return json.loads(json_text)
                    except:
                        return {"ok": False, "error": json_text[:200]}
            return {"ok": False, "error": text[:200]}
        return {"ok": False, "error": result.stderr[:200] if result.stderr else "exit=" + str(result.returncode)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def get_dingtalk_token():
    import requests
    url = "https://oapi.dingtalk.com/gettoken?appkey=" + DT_APP_KEY + "&appsecret=" + DT_APP_SECRET
    resp = requests.get(url, timeout=10)
    data = resp.json()
    if data.get("errcode") != 0:
        raise RuntimeError("Token failed: " + str(data))
    return data["access_token"]


def send_dingtalk_actioncard(title, text, record_url):
    import requests
    url = DT_WEBHOOK
    if DT_ROBOT_SECRET and DT_ROBOT_SECRET.strip():
        timestamp = str(round(time.time() * 1000))
        sign_str = timestamp + "\\n" + DT_ROBOT_SECRET
        signature = base64.b64encode(
            hmac.new(DT_ROBOT_SECRET.encode(), sign_str.encode(), digestmod=hashlib.sha256).digest()
        ).decode()
        sep = "&" if "?" in url else "?"
        url = url + sep + "timestamp=" + timestamp + "&sign=" + signature
    payload = {
        "msgtype": "actionCard",
        "actionCard": {
            "title": title,
            "text": text,
            "btnOrientation": "0",
            "singleTitle": "\u270f\ufe0f \u70b9\u51fb\u8fd9\u91cc\u586b\u5199\u4f9b\u5e94\u5546",
            "singleURL": record_url
        }
    }
    resp = requests.post(url, json=payload, timeout=10)
    return resp.json()


def send_dingtalk_markdown(title, text):
    import requests
    url = DT_WEBHOOK
    if DT_ROBOT_SECRET and DT_ROBOT_SECRET.strip():
        timestamp = str(round(time.time() * 1000))
        sign_str = timestamp + "\\n" + DT_ROBOT_SECRET
        signature = base64.b64encode(
            hmac.new(DT_ROBOT_SECRET.encode(), sign_str.encode(), digestmod=hashlib.sha256).digest()
        ).decode()
        sep = "&" if "?" in url else "?"
        url = url + sep + "timestamp=" + timestamp + "&sign=" + signature
    payload = {"msgtype": "markdown", "markdown": {"title": title, "text": text}}
    resp = requests.post(url, json=payload, timeout=10)
    return resp.json()

def insert_dingtalk_record(token, fields_dict):
    import requests
    headers = {
        "x-acs-dingtalk-access-token": token,
        "Content-Type": "application/json",
        "operatorId": DT_OPERATOR_ID,
    }
    url = "https://api.dingtalk.com/v1.0/notable/bases/" + DT_BASE_ID + "/sheets/" + DT_SHEET2_ID + "/records"
    payload = {"records": [{"fields": fields_dict}]}
    resp = requests.post(url, headers=headers, json=payload, timeout=15)
    if resp.status_code == 200:
        records = resp.json().get("value", [])
        if records:
            return records[0].get("id", "")
    return ""


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"checked_ids": [], "last_check": None, "processed_filenames": [], "duplicates": []}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False)


def parse_filename(filename):
    name = os.path.splitext(filename)[0]
    parts = name.split("-")
    known_types = ["新增原辅料明细表", "新增原辅料明细", "企业质量标准", "质量标准", "原料企业质量标准"]
    if len(parts) >= 5 and (parts[0] in known_types or parts[0].startswith("新增原辅料明细")):
        # 文件名格式: 类型-平台-申请人-编码-名称.docx
        return {"type": parts[0], "platform": parts[1], "applicant": parts[2], "material_code": parts[3], "material_name": "-".join(parts[4:])}
    # 新增原辅料明细表格式: 新增原辅料明细表-平台-申请人-编码-名称.docx
    for kw in ["原辅料", "质量"]:
        if kw in name:
            p = {"type": kw, "platform": "", "applicant": "", "material_code": "", "material_name": ""}
            if len(parts) >= 3:
                p["platform"] = parts[1]
                p["applicant"] = parts[2]
            # 4段=有编码无名称: 类型-平台-申请人-编码
            if len(parts) >= 4:
                p["material_code"] = parts[3]
            # 5段=完整: 类型-平台-申请人-编码-名称
            if len(parts) >= 5:
                p["material_name"] = "-".join(parts[4:])
            return p
    for kw in ["原辅料", "质量"]:
        if kw in name:
            return {"type": kw, "platform": "", "material_code": "", "material_name": "", "applicant": ""}
    return None


def extract_info_from_subject(subject):
    info = {"applicant": "", "platform": "", "material_code": "", "material_name": ""}
    # 方法1：从括号中提取申请人（如"测试（张三/平台G）"）
    m = re.search(r'[（(]([^）)]+)[）)]', subject)
    if m:
        info["applicant"] = m.group(1)
    # 方法2：从【xx】-平台-申请人-N个物料 格式中提取
    # 主题格式: 【企业质量协议签订】-平台A-张三-16个物料
    pat = re.compile(r'[【\[]+[^】\]]+[】\]]?[-—]\s*([^-—]+)[-—]\s*([^-—]+?)\s*(?:[-—]\s*\d+\s*个物料)?$')
    m2 = pat.search(subject)
    if m2:
        if not info["applicant"]:
            info["applicant"] = m2.group(2).strip()
    
    platforms = ["平台E","平台F","平台A","平台G","平台B","技术研究中心","平台D","平台H","平台C"]
    for p in platforms:
        if p in subject:
            info["platform"] = p
            break
    
    # 提取物料编码（如有）
    m3 = re.search(r'(\d{8,12})', subject)
    if m3:
        info["material_code"] = m3.group(1)
    # 提取物料数量（如"16个物料"）
    m4 = re.search(r'(\d+)\s*个物料', subject)
    if m4:
        info["material_count"] = int(m4.group(1))
    
    return info


def parse_raw_material_table(docx_path):
    """从新增原辅料明细表docx中解析所有物料信息"""
    if docx_path.lower().endswith('.xlsx'):
        import openpyxl
        wb = openpyxl.load_workbook(docx_path, read_only=True, data_only=True)
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
            return []
        name_col = col("物料名称")
        cn_col = col("物料中文名称")
        cas_col = col("货号")
        spec_col = col("规格")
        materials = []
        for row in rows[1:]:
            if row is None or len(row) <= code_col:
                continue
            code = str(row[code_col] or "").strip()
            if not code:
                continue
            cas = str(row[cas_col] or "").strip() if cas_col is not None and cas_col < len(row) else ""
            m_pn = re.search(r"货号[：:]\s*([^\n;；]+)", cas)
            part_no = m_pn.group(1).strip() if m_pn else (cas if cas and "CAS号" not in cas and "克隆号" not in cas else "")
            materials.append({
                'code': code,
                'name': str(row[name_col] or "").strip() if name_col is not None and name_col < len(row) else "",
                'name_cn': str(row[cn_col] or "").strip() if cn_col is not None and cn_col < len(row) else "",
                'part_no': part_no,
                'spec': str(row[spec_col] or "").strip() if spec_col is not None and spec_col < len(row) else "",
            })
        return materials
    from docx import Document
    doc = Document(docx_path)
    for table in doc.tables:
        first = "".join(c.text.strip()[:10] for c in table.rows[0].cells)
        if "物料编码" in first:
            materials = []
            for ri, row in enumerate(table.rows):
                if ri == 0:
                    continue
                cells = [c.text.strip() for c in row.cells]
                if not cells[0] or cells[0] in ('/', ''):
                    continue
                # 提取货号
                cas_raw = cells[6] if len(cells) > 6 else ''
                part_no = ''
                m_pn = re.search(r'货号[：:]\s*([^;；\\n]+)', cas_raw)
                if m_pn:
                    part_no = m_pn.group(1).strip()
                if not part_no and cas_raw and cas_raw != '/':
                    part_no = cas_raw.split('\\n')[0].split('；')[0].split(';')[0].strip()
                materials.append({
                    'code': cells[0],
                    'name': cells[1],
                    'name_cn': cells[2] if len(cells) > 2 else "",
                    'part_no': part_no,
                    'spec': cells[4] if len(cells) > 4 else '',
                })
            return materials
    return []



def match_std_to_material(std_filename, materials):
    """Match quality standard file to material by name/code"""
    fn = std_filename.lower()
    # Remove common suffixes
    for suffix in ["原料企业质量标准", "企业质量标准", "质量标准"]:
        fn = fn.replace(suffix, "")
    fn = fn.replace(".docx", "").strip("-_. ")
    
    best = None
    best_score = 0
    for mat in materials:
        cand_names = []
        if mat.get("name"):
            cand_names.append(mat["name"].lower().replace("（", "(").replace("）", ")"))
        if mat.get("name_cn"):
            cand_names.append(mat["name_cn"].lower().replace("（", "(").replace("）", ")"))
        mc = mat["code"]
        # Name match
        for mn in cand_names:
            if mn and mn in fn:
                if len(mn) > best_score:
                    best_score = len(mn)
                    best = mat
        # Code match (higher priority)
        if mc in fn:
            if 30 > best_score:
                best_score = 30
                best = mat
    return best

def main():
    import requests
    log("=== 质量协议邮件监控 - 开始轮询 ===")
    
    state = load_state()
    checked_ids = set(state.get("checked_ids", []))
    
    result = run_agently(["message", "+list", "--dir", "inbox", "--limit", "20"])
    if not result.get("ok"):
        log("列表邮件失败: " + result.get("error", "未知"))
        return
    
    messages = result.get("data", {}).get("data", [])
    if not messages:
        log("收件箱为空")
        return
    
    new_msgs = [m for m in messages if m.get("message_id", "") and m["message_id"] not in checked_ids]
    if not new_msgs:
        log("无新邮件")
        return
    
    log("发现 " + str(len(new_msgs)) + " 封新邮件")
    
    dt_token = None
    try:
        dt_token = get_dingtalk_token()
    except Exception as e:
        log("钉钉Token失败: " + str(e))
    
    triggered_items = []
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    DUP_FILE = os.path.join(os.path.dirname(STATE_FILE), "dedup_log.txt")
    
    for msg in new_msgs:
        mid = msg["message_id"]
        downloaded_files = []
        subject = msg.get("subject", "")
        log("检查: " + subject[:60])
        
        from qa_return_handler import RETURN_KEYWORDS, is_return_email, process_returned_files
        is_qa = any(kw in subject for kw in ["质量协议", "质量保证", "质协", "新增原辅料", "质量标准"]) or any(k in subject for k in RETURN_KEYWORDS)
        if not is_qa:
            checked_ids.add(mid)
            continue
        
        read_result = run_agently(["message", "+read", "--id", mid])
        if not read_result.get("ok"):
            checked_ids.add(mid)
            continue
        
        msg_data = read_result.get("data", {})
        email_date = msg_data.get("created_at", "")
        email_date_str = email_date[:10] if email_date else today_str
        
        attachments = msg_data.get("attachments", [])
        if not attachments:
            checked_ids.add(mid)
            continue

        is_return = is_return_email(subject, [a.get("filename", "") for a in attachments])
        return_files = []
        
        # 收集该邮件中所有匹配的物料信息（去重）
        material_items = {}  # key=申请人_编码, value={...}
        processed = state.get("processed_filenames", [])
        
        # 先用主题提取后备信息（附件名解析失败时使用）
        subj_fallback = extract_info_from_subject(subject)
        has_subj_info = bool(subj_fallback.get("applicant") or subj_fallback.get("platform") or subj_fallback.get("material_code"))
        
        for att in attachments:
            fname = att.get("filename", "")
            att_id = att.get("attachment_id", "")
            if not fname or not att_id:
                continue
            
            # 下载附件（无论文件名是否能解析都下载，以便解压）
            dl_dir = "qa_download"
            dl_path = os.path.join(os.path.expanduser("~"), dl_dir)
            os.makedirs(dl_path, exist_ok=True)
            os.makedirs(QA_EMAIL_DIR, exist_ok=True)
            # 下载附件（最多重试3次，验证文件大小>0）
            saved_file = ""
            for attempt in range(3):
                dl_result = run_agently(["attachment", "+download", "--msg", mid, "--att", att_id, "--output", "./" + dl_dir])
                if dl_result.get("ok"):
                    saved = dl_result.get("data", {}).get("saved_to", "")
                    if saved and os.path.exists(saved) and os.path.getsize(saved) > 0:
                        dst = os.path.join(QA_EMAIL_DIR, fname)
                        shutil.copy2(saved, dst)
                        saved_file = dst
                        log(f"已下载({attempt+1}次): {dst} ({os.path.getsize(dst)} bytes)")
                        break
                    else:
                        log(f"下载失败(第{attempt+1}次): 文件为空或不存在")
                else:
                    log(f"下载失败(第{attempt+1}次): {dl_result.get('error', 'unknown')}")
            if not saved_file:
                log(f"⚠️ 附件下载失败(3次重试): {fname}")
                continue

            # 下载成功：如果是压缩包，解压并扫描docx；解压产物复制到 QA 目录，供后续解析/归档
            extracted = extract_archive(saved_file)
            for ef in extracted:
                if ef != saved_file:
                    dst_ef = os.path.join(QA_EMAIL_DIR, os.path.basename(ef))
                    shutil.copy2(ef, dst_ef)
                    if dst_ef.endswith('.doc'):
                        from doc_converter import convert_doc_to_docx
                        conv = convert_doc_to_docx(dst_ef)
                        if conv and conv not in downloaded_files:
                            downloaded_files.append(conv)
                    if dst_ef not in downloaded_files:
                        downloaded_files.append(dst_ef)
                elif saved_file not in downloaded_files:
                    downloaded_files.append(saved_file)

            if is_return:
                for ef in extracted:
                    dst_ef = os.path.join(QA_EMAIL_DIR, os.path.basename(ef))
                    if os.path.exists(dst_ef) and dst_ef not in return_files:
                        return_files.append(dst_ef)
                continue
            
            # 解析文件名
            parsed = parse_filename(fname)
            
            # 如果文件名解析失败，尝试用主题信息兜底
            if not parsed:
                if has_subj_info and saved_file:
                    log("文件名未匹配规范格式，使用邮件主题信息: " + subject[:60])
                    dk = subj_fallback["applicant"] + "_" + subj_fallback.get("material_code","")
                    if dk and dk not in material_items:
                        material_items[dk] = {
                            "platform": subj_fallback.get("platform", "") or "待确认",
                            "applicant": subj_fallback.get("applicant", "") or "待确认",
                            "material_code": subj_fallback.get("material_code", "") or "",
                            "material_name": subj_fallback.get("material_name", "") or "",
                        }
                        processed.append(dk)
                        log("主题后备: 平台={p}, 申请人={a}, 编码={c}".format(
                            p=subj_fallback["platform"], a=subj_fallback["applicant"], c=subj_fallback["material_code"]))
                continue  # 文件名无法解析，跳过常规处理
            
            log("匹配附件: " + fname)
            
            # 检查解析结果的关键字段是否有效
            has_valid_info = bool(parsed.get("applicant") and parsed.get("material_code"))
            
            if not has_valid_info and has_subj_info:
                # 文件名解析出的信息不全，用主题后备
                log("文件名信息不完整，使用邮件主题信息: " + subject[:60])
                dk = subj_fallback["applicant"] + "_" + subj_fallback.get("material_code","")
                if dk and dk not in material_items:
                    material_items[dk] = {
                        "platform": subj_fallback.get("platform", "") or "待确认",
                        "applicant": subj_fallback.get("applicant", "") or "待确认",
                        "material_code": subj_fallback.get("material_code", "") or "",
                        "material_name": subj_fallback.get("material_name", "") or "",
                    }
                    processed.append(dk)
                    log("主题后备: 平台={p}, 申请人={a}, 编码={c}".format(
                        p=subj_fallback["platform"], a=subj_fallback["applicant"], c=subj_fallback["material_code"]))
                continue
            
            dk = parsed.get("applicant", "") + "_" + parsed.get("material_code", "")
            
            # 查重
            if dk in processed:
                log("跳过重复: " + dk)
                dup_log = "[{d}] DUP: {key} | sub={sub}\n".format(
                    d=datetime.now().isoformat(), key=dk, sub=subject[:30]
                )
                with open(DUP_FILE, "a") as _df:
                    _df.write(dup_log)
                continue
            
            # 收集此物料信息（必须有申请人和编码才算一条独立物料）
            if dk and parsed.get("material_code") and dk not in material_items:
                material_items[dk] = {
                    "platform": parsed.get("platform", "") or "",
                    "applicant": parsed.get("applicant", "") or "待确认",
                    "material_code": parsed.get("material_code", "") or "",
                    "material_name": parsed.get("material_name", "") or "",
                }
                processed.append(dk)
        
        state["processed_filenames"] = processed

        if is_return:
            log("识别为供应商回传邮件，进入第⑧阶段处理")
            process_returned_files(return_files, subject, cfg=_cfg)
            checked_ids.add(mid)
            continue
        
        # 从新增原辅料明细表提取品牌信息（支持 docx / xlsx，多份明细合并）
        brand_map = {}
        for downloaded_file in downloaded_files:
            if "原辅料" not in downloaded_file:
                continue
            try:
                if downloaded_file.lower().endswith('.xlsx'):
                    import openpyxl as _px
                    _wb = _px.load_workbook(downloaded_file, read_only=True, data_only=True)
                    _ws = _wb.active
                    rows = list(_ws.iter_rows(values_only=True))
                    if not rows:
                        continue
                    hdr = [re.sub(r"\s+", "", str(c or "")) for c in rows[0]]
                    code_col = next((i for i, h in enumerate(hdr) if "物料编码" in h), None)
                    brand_col = next((i for i, h in enumerate(hdr) if "品牌" in h), None)
                    if code_col is None or brand_col is None:
                        continue
                    for row in rows[1:]:
                        if row is None or len(row) <= max(code_col, brand_col):
                            continue
                        code = str(row[code_col] or "").strip()
                        if not code:
                            continue
                        brand_raw = str(row[brand_col] or "").strip()
                        for sep in [';', '；', '\n']:
                            si = brand_raw.find(sep)
                            if si > 0:
                                brand_raw = brand_raw[:si]
                        brand_raw = brand_raw.strip()
                        if brand_raw and brand_raw not in ('/', '-', ''):
                            brand_map.setdefault(code, brand_raw)
                else:
                    from docx import Document
                    doc = Document(downloaded_file)
                    for table in doc.tables:
                        hdr = [c.text.strip() for c in table.rows[0].cells]

                        def col(*keys):
                            for i, h in enumerate(hdr):
                                if any(k in h for k in keys):
                                    return i
                            return None

                        code_col = col("物料编码")
                        brand_col = col("品牌")
                        if code_col is None or brand_col is None:
                            continue
                        for ri, row in enumerate(table.rows):
                            if ri == 0:
                                continue
                            cells = [c.text.strip() for c in row.cells]
                            if len(cells) <= max(code_col, brand_col):
                                continue
                            code = cells[code_col]
                            if not code or code == '/':
                                continue
                            brand_raw = cells[brand_col]
                            for p in ['品牌：', '品牌:']:
                                if brand_raw.startswith(p):
                                    brand_raw = brand_raw[len(p):]
                            for sep in [';', '；', '\n']:
                                si = brand_raw.find(sep)
                                if si > 0:
                                    brand_raw = brand_raw[:si]
                            brand_raw = brand_raw.strip()
                            if brand_raw and brand_raw not in ('/', '-', ''):
                                brand_map.setdefault(code, brand_raw)
            except Exception as e:
                log("品牌提取失败: " + str(e))
        if brand_map:
            log("品牌映射: " + str(brand_map))
            for dk in list(material_items.keys()):
                mc = material_items[dk].get("material_code", "")
                if mc in brand_map:
                    material_items[dk]["brand"] = brand_map[mc]
        

        
        # === 新增：从原辅料明细表批量解析物料（兜底路径） ===
        # 如果当前material_items为空或信息不全，尝试从原辅料明细表解析
        has_incomplete = any(
            not item.get("material_code") or not re.fullmatch(r"\d{8,12}", str(item.get("material_code")))
            for item in material_items.values()
        )
        if not material_items or has_incomplete:
            raw_mat_files = []
            std_files = []
            for df in downloaded_files:
                if df.lower().endswith('.xlsx'):
                    try:
                        import openpyxl as _px
                        _wb = _px.load_workbook(df, read_only=True, data_only=True)
                        _ws = _wb.active
                        _hdr = [re.sub(r"\s+", "", str(c or "")) for c in next(_ws.iter_rows(values_only=True), [])]
                        if any("物料编码" in h for h in _hdr):
                            raw_mat_files.append(df)
                    except Exception:
                        pass
                if df.endswith('.docx'):
                    try:
                        from docx import Document as DocReader
                        d = DocReader(df)
                        for t in d.tables:
                            first_cell = "".join(c.text.strip()[:10] for c in t.rows[0].cells)
                            if "物料编码" in first_cell:
                                raw_mat_files.append(df)
                                break
                    except:
                        pass
                if "质量" in os.path.basename(df) and df.endswith((".docx", ".doc")):
                    if df.endswith(".doc") and os.path.exists(os.path.splitext(df)[0] + ".docx"):
                        continue
                    std_files.append(df)

            if not raw_mat_files:
                log("本邮件未下载原辅料明细表，跳过批量解析")
                # 不跨邮件搜索QA_DIR，每封邮件只处理自己的附件

            if raw_mat_files:
                all_materials = []
                for rmf in raw_mat_files:
                    all_materials.extend(parse_raw_material_table(rmf))
                # 多份明细合并去重（物料编码唯一）
                seen = set()
                merged = []
                for mat in all_materials:
                    if mat.get("code") and mat["code"] not in seen:
                        seen.add(mat["code"])
                        merged.append(mat)
                all_materials = merged
                log("原辅料明细表解析: {} 个物料, {} 个质量标准文件".format(len(all_materials), len(std_files)))

                if all_materials:
                    # 用主题信息补全
                    platform = subj_fallback.get("platform", "") if has_subj_info else ""
                    applicant = subj_fallback.get("applicant", "") if has_subj_info else ""
                    if not applicant:
                        # 尝试从现有 material_items 中取
                        for item in material_items.values():
                            if item.get("applicant"):
                                applicant = item["applicant"]
                                break

                    # 清空旧的空白记录，重建
                    old_count = len(material_items)
                    material_items = {}

                    for mat in all_materials:
                        code = mat["code"]
                        # 尝试匹配质量标准
                        matched_std = None
                        for sf in std_files:
                            sf_basename = os.path.basename(sf)
                            matched = match_std_to_material(sf_basename, [mat])
                            if matched:
                                matched_std = sf_basename
                                break

                        dk = applicant + "_" + code
                        if dk not in material_items:
                            material_items[dk] = {
                                "platform": platform or "待确认",
                                "applicant": applicant or "待确认",
                                "material_code": code,
                                "material_name": mat["name"],
                                "part_no": mat["part_no"],
                                "brand": brand_map.get(code, ""),
                                "matched_std": matched_std or "",
                            }
                            if dk not in processed:
                                processed.append(dk)

                    log("原辅料明细表: 重建 {} 条物料记录（原{}条）".format(len(material_items), old_count))
                    state["processed_filenames"] = processed
        
        # 写入前查重：拉取已有记录，避免重复写入
        existing_keys = set()
        if dt_token:
            try:
                _ex_url = "https://api.dingtalk.com/v1.0/notable/bases/" + DT_BASE_ID + "/sheets/" + DT_SHEET2_ID + "/records?maxResults=100"
                _ex_r = requests.get(_ex_url, headers={"x-acs-dingtalk-access-token": dt_token, "Content-Type": "application/json", "operatorId": DT_OPERATOR_ID}, timeout=15)
                for _rec in _ex_r.json().get("records", []):
                    _f = _rec.get("fields", {})
                    _code = _f.get("物料编码", "")
                    if isinstance(_code, dict): _code = _code.get("text", "")
                    _app = _f.get("研发申请人", "")
                    if isinstance(_app, dict): _app = _app.get("text", "")
                    _plt = _f.get("业务平台", "")
                    if isinstance(_plt, dict): _plt = _plt.get("name", "")
                    _status = _f.get("节点状态", "")
                    if isinstance(_status, dict): _status = _status.get("name", "")
                    if _status == "已删除":
                        continue
                    existing_keys.add(str(_code) + "_" + str(_app) + "_" + str(_plt))
                log("查重: 表中已有 " + str(len(existing_keys)) + " 条有效记录")
            except Exception as e:
                log("查重失败(继续写入): " + str(e))
        
        # 每个物料写一条钉钉记录
        if material_items and dt_token:
            for dk, item in material_items.items():
                # 查重：如果表中已有相同 物料编码+申请人+平台 则跳过
                _ek = str(item.get("material_code", "")) + "_" + str(item.get("applicant", "")) + "_" + str(item.get("platform", ""))
                if _ek in existing_keys:
                    log("跳过(表中已存在): " + _ek)
                    continue
                fields = {
                    "任务流水号": "QA-" + datetime.now().strftime("%Y%m%d-%H%M%S"),
                    "研发申请人": item["applicant"],
                    "业务平台": item["platform"] or "待确认",
                    "物料编码": item["material_code"] or "",
                    "物料名称": item["material_name"] or "",
                    "指定品牌": item.get("brand", ""),
                    "节点状态": "待确认供应商",
                    "填写日期": today_str,
                    "接收日期": email_date_str,
                }
                if "升版" in subject or any("升版" in (a.get("filename", "")) for a in attachments):
                    fields["变更类型"] = detect_change_type({"升版原因": subject}, subject)
                    fields["升版原因"] = subject[:200]
                record_id = insert_dingtalk_record(dt_token, fields)
                if record_id:
                    record_url = BASE_RECORD_URL + record_id
                    triggered_items.append({
                        "platform": item["platform"] or "待确认",
                        "material_code": item["material_code"] or "待解析",
                        "material_name": item["material_name"] or "待解析",
                        "applicant": item["applicant"],
                        "record_url": record_url,
                        "email_date": email_date_str,
                    })
                    log("钉钉表: 写入, 平台={p}, 申请人={a}, 编码={c}".format(
                        p=item["platform"], a=item["applicant"], c=item["material_code"]
                    ))
                else:
                    log("钉钉表: 写入失败")
        
        checked_ids.add(mid)
    
    # 批量推送
    if triggered_items and dt_token:
        n = len(triggered_items)
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        if n == 1:
            item = triggered_items[0]
            text_parts = [
                "### 新物料需分配供应商",
                "",
                "**研发已提交新物料，请确认供应商全称，以便系统自动生成《质量协议》。**",
                "",
                "- **所在平台：** " + item["platform"],
                "- **物料编码：** " + item["material_code"],
                "- **物料名称：** " + item["material_name"],
                "- **研发申请人：** " + item["applicant"],
                "",
                "---",
                "**操作指引：** 点击下方按钮，在表格中填入供应商全称，并将状态改为【已确认供应商】。系统将自动接管后续流程。",
                "",
                "\u23f0 " + now
            ]
            r = send_dingtalk_actioncard(
                title="\u6025\uff01\u8bf7\u786e\u8ba4\u65b0\u7269\u6599\u4f9b\u5e94\u5546",
                text="\n\n".join(text_parts),
                record_url=item["record_url"]
            )
            ok = r.get("errcode") == 0
            log("群通知(单条): " + ("成功" if ok else "失败"))
        else:
            items_list = []
            for i, item in enumerate(triggered_items, 1):
                items_list.append(
                    str(i) + ". **" + item["material_name"] + "** | " + item["platform"] + " | " + item["applicant"]
                    + "\\n   [\u70b9\u51fb\u586b\u5199\u4f9b\u5e94\u5546](" + item["record_url"] + ")"
                )
            text_parts = [
                "### \u65b0\u7269\u6599\u9700\u5206\u914d\u4f9b\u5e94\u5546",
                "",
                "**本轮共收到 " + str(n) + " 条物料待处理：**",
                "",
            ] + items_list + [
                "",
                "---",
                "**操作指引：** 点击每条物料的链接，在表格中填入供应商全称。",
                "",
                "\u23f0 " + now
            ]
            r = send_dingtalk_markdown(
                title="\u65b0\u7269\u6599 " + str(n) + " \u6761\u5f85\u5904\u7406",
                text="\n\n".join(text_parts)
            )
            ok = r.get("errcode") == 0
            log("群通知(批量" + str(n) + "条): " + ("成功" if ok else "失败"))
    
    elif not triggered_items:
        log("新邮件不涉及质量协议")
    
    state["checked_ids"] = list(checked_ids)
    state["last_check"] = datetime.now().isoformat()
    save_state(state)
    log("完成 (已跟踪 " + str(len(checked_ids)) + " 封)")
    log("本轮触发: " + str(len(triggered_items)) + " 条物料")


if __name__ == "__main__":
    main()
