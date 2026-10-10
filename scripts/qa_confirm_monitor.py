#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检测已确认供应商，每1小时运行；生成协议并写入 data/email_tasks，由 send_email_tasks.py 发送"""

import glob, re, shutil, requests, json, os, sys, subprocess
from datetime import datetime
from docx import Document

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qa_config_loader import load_qa_config

_cfg = load_qa_config()

### CONFIG
DT_APP_KEY = _cfg.get("app_key", "")
DT_APP_SECRET=DT_APP_SECRET = _cfg.get("app_secret", "")
DT_OPERATOR_ID = _cfg.get("operator_id", "")
DT_BASE_ID = _cfg.get("base_id", "")
DT_SHEET2_ID = _cfg.get("sheet2_id", "")
QA_DIR = _cfg.get("qa_dir", "")
if os.name != "nt" and not os.path.exists(QA_DIR):
    QA_DIR = "/app/email-database/QA"
    import logging as _lg; _lg.getLogger().info("QA_DIR fallback: " + QA_DIR)
_state_f = _cfg.get("state_file", "")
if _state_f:
    STATE_BASE = os.path.dirname(_state_f)
    STATE_FILE = os.path.join(STATE_BASE, ".confirm_state.json")
else:
    STATE_BASE = ""
    STATE_FILE = ""
# Container fallback: use email-database dir
if not STATE_FILE or not os.path.exists(STATE_BASE):
    _alt = "/app/email-database" if os.name != "nt" else _cfg.get("qa_dir", "")
    if os.path.exists(_alt):
        STATE_FILE = os.path.join(_alt, ".confirm_state.json")
PROJ_DIR = "/app/agreement-automation" if os.name != "nt" else r".\agreement-automation"
DATA_DIR = "/app/data" if os.name != "nt" else r".\data"
CONTACTS_XLSX = os.path.join(PROJ_DIR, "平台联系人的对应关系.xlsx")
TEMPLATE_DOCX = os.path.join(PROJ_DIR, "质量保证协议模板.docx")

def run_agently(args):
    """执行 agently-cli 命令"""
    arg_str = " ".join(args).replace('"', '\\"')
    home_dir = os.path.expanduser("~")
    try:
        if os.name == 'nt':
            cmd = 'cmd.exe /c "agently-cli ' + arg_str + '"'
            result = __import__('subprocess').run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30, shell=True, cwd=home_dir)
        else:
            cmd = ['agently-cli'] + args
            result = __import__('subprocess').run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30, cwd=home_dir)
        if result.returncode == 0 and result.stdout:
            text = result.stdout
            brace_idx = text.find("{")
            if brace_idx >= 0:
                json_text = text[brace_idx:]
                end_idx = json_text.rfind("}")
                if end_idx >= 0:
                    json_text = json_text[:end_idx+1]
                    try:
                        return __import__('json').loads(json_text)
                    except:
                        return {"ok": False, "error": json_text[:200]}
            return {"ok": False, "error": text[:200]}
        return {"ok": False, "error": result.stderr[:200] if result.stderr else "exit=" + str(result.returncode)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def log(msg):
    t = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{t}] {msg}")

def get_val(v):
    if isinstance(v, dict):
        return v.get("name") or v.get("text") or ""
    if isinstance(v, list):
        for item in v:
            if isinstance(item, dict):
                val = item.get("name") or item.get("text") or item.get("value") or ""
                if val:
                    return val
            elif item:
                return str(item)
        return ""
    return str(v) if v else ""

def get_linked_supplier(headers, field_val):
    if not isinstance(field_val, dict):
        return str(field_val) if field_val else ""
    ids = field_val.get("linkedRecordIds", [])
    if not ids:
        return ""
    for sid in ["YQva3Vo", "9uU30VN"]:
        url = "https://api.dingtalk.com/v1.0/notable/bases/" + DT_BASE_ID + "/sheets/" + sid + "/records/" + ids[0]
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            lf = r.json().get("fields", {})
            val = lf.get("供应商全称", "")
            if isinstance(val, dict):
                return val.get("text", "")
            return str(val) if val else ""
    return ""

def get_email_by_reviewer(reviewer):
    """按甲方联系人（审核人）从 平台联系人的对应关系.xlsx 查邮箱"""
    if not reviewer:
        return ""
    try:
        import pandas as _pd
        _df = _pd.read_excel(CONTACTS_XLSX, engine="openpyxl")
        _row = _df[_df["审核人"] == reviewer]
        if not _row.empty:
            _email = str(_row.iloc[0].get("邮箱地址", "") or "")
            if _email and _email != "nan":
                return _email
    except Exception as e:
        log("按甲方联系人查邮箱失败: " + str(e))
    return ""


def safe_name(s):
    return re.sub(r'[\\/*?:"<>|]', "", s)


def load_history(supplier_dir):
    p = os.path.join(supplier_dir, "history.json")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return None


def save_history(supplier_dir, history):
    with open(os.path.join(supplier_dir, "history.json"), "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def find_raw_files(codes):
    """在 QA_DIR 中找出包含指定物料编码的原辅料明细表"""
    result = []
    if not os.path.exists(QA_DIR):
        return result
    for fn in sorted(os.listdir(QA_DIR)):
        if fn.startswith("~$") or re.search(r"_\d{14}\.(docx|xlsx)$", fn):
            continue
        if "原辅料" not in fn or not (fn.endswith(".docx") or fn.endswith(".xlsx")):
            continue
        fp = os.path.join(QA_DIR, fn)
        try:
            file_codes = set()
            if fn.endswith(".xlsx"):
                import openpyxl
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                if rows:
                    hdr = [re.sub(r"\s+", "", str(c or "")) for c in rows[0]]
                    code_col = next((i for i, h in enumerate(hdr) if "物料编码" in h), None)
                    if code_col is not None:
                        for row in rows[1:]:
                            if row and len(row) > code_col and str(row[code_col] or "").strip():
                                file_codes.add(str(row[code_col]).strip())
            else:
                doc = Document(fp)
                for t in doc.tables:
                    first = "".join(c.text.strip()[:10] for c in t.rows[0].cells)
                    if "物料编码" in first:
                        for ri in range(1, len(t.rows)):
                            c0 = t.rows[ri].cells[0].text.strip()
                            if c0:
                                file_codes.add(c0)
                        break
            if file_codes & set(codes):
                result.append(fp)
        except Exception as e:
            log("解析原辅料失败 " + fn + ": " + str(e))
    return result


def find_std_files(codes, names_by_code):
    """按物料编码/名称匹配质量标准文件，data/_party_c_std_map.json 兜底"""
    std_map = {}
    if os.path.exists(QA_DIR):
        for fn in sorted(os.listdir(QA_DIR)):
            if fn.startswith("~$") or re.search(r"_\d{14}\.(docx|xlsx)$", fn):
                continue
            if not fn.endswith((".docx", ".doc")) or "质量" not in fn:
                continue
            fp = os.path.join(QA_DIR, fn)
            if os.path.getsize(fp) <= 0:
                continue
            if fn.endswith(".doc") and os.path.exists(os.path.splitext(fp)[0] + ".docx"):
                continue
            for code in codes:
                if code in fn and code not in std_map:
                    std_map[code] = fp
        for code in codes:
            if code in std_map:
                continue
            name = names_by_code.get(code, "")
            if not name:
                continue
            try:
                from generate_qa_agreement_v3 import std_filename_key
            except Exception:
                std_filename_key = lambda f: f
            norm_name = re.sub(r"\s+", "", name.replace("（", "(").replace("）", ")").replace("β", "beta"))
            for fn in sorted(os.listdir(QA_DIR)):
                if fn.startswith("~$") or re.search(r"_\d{14}\.(docx|xlsx)$", fn):
                    continue
                if fn.endswith((".docx", ".doc")) and "质量" in fn:
                    if fn.endswith(".doc") and os.path.exists(os.path.splitext(os.path.join(QA_DIR, fn))[0] + ".docx"):
                        continue
                    key = std_filename_key(fn)
                    if key and len(key) >= 2 and (key in norm_name or norm_name in key):
                        std_map[code] = os.path.join(QA_DIR, fn)
                        break
    std_map_file = os.path.join(DATA_DIR, "_party_c_std_map.json")
    if os.path.exists(std_map_file):
        try:
            with open(std_map_file, encoding="utf-8") as f:
                extra = json.load(f)
            for code in codes:
                if code not in std_map and code in extra and os.path.exists(extra[code]):
                    std_map[code] = extra[code]
        except Exception as e:
            log("读取 _party_c_std_map.json 失败: " + str(e))
    return std_map


def collect_supplier_files(supplier_dir, history, codes, names_by_code):
    """汇总原辅料明细/质量标准：优先使用 data/{供应商}/materials 归档，缺失/新增物料回退 QA_DIR 扫描"""
    raw_files, std_map = [], {}
    if history:
        for v in history.get("versions", []):
            for rf in v.get("raw_files", []):
                p = os.path.join(supplier_dir, rf)
                if os.path.exists(p) and p not in raw_files:
                    raw_files.append(p)
            sf = v.get("std_files", {})
            if isinstance(sf, dict):
                items = sf.items()
            else:
                items = [(None, p) for p in sf]
            for code, p in items:
                fp = p if os.path.isabs(p) else os.path.join(supplier_dir, p)
                if not os.path.exists(fp):
                    continue
                if code is not None:
                    if code not in std_map:
                        std_map[code] = fp
                else:
                    base = os.path.basename(fp)
                    for c in codes:
                        if c in base or (names_by_code.get(c) and names_by_code[c] in base):
                            if c not in std_map:
                                std_map[c] = fp
                            break
    # 归档不足/新增物料 -> QA_DIR 扫描兜底
    qa_raw = find_raw_files(codes)
    if not raw_files:
        raw_files = qa_raw
    else:
        for p in qa_raw:
            if p not in raw_files:
                raw_files.append(p)
    qa_std = find_std_files(codes, names_by_code)
    for code, p in qa_std.items():
        std_map.setdefault(code, p)
    return raw_files, std_map


def send_email_via_agently(to_addr, subject, body, attachments):
    """通过 agently-cli 发送邮件"""
    log("发送邮件: " + to_addr + " | " + subject[:40])
    try:
        # 先用 +send 命令创建邮件
        send_args = ["message", "+send", "--to", to_addr, "--subject", subject, "--body", body]
        for att in attachments:
            send_args.extend(["--attach", att])
        result = run_agently(send_args)
        if result.get("ok"):
            log("✅ 邮件已发送: " + to_addr)
            return True
        else:
            log("❌ 邮件发送失败: " + str(result.get("error", "")))
            return False
    except Exception as e:
        log("❌ 邮件发送异常: " + str(e))
        return False

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"processed_ids": []}

def save_state(s):
    d = os.path.dirname(STATE_FILE)
    if d:
        os.makedirs(d, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False)

def auto_refresh_oa_serials(records):
    """检测钉钉"OA审批流水号"变化，自动刷新已生成协议的文件名/页眉/history（接进轮询）"""
    rid_map = {}
    for hp in glob.glob(os.path.join(DATA_DIR, "*", "*", "history.json")):
        supplier = os.path.basename(os.path.dirname(hp))
        try:
            with open(hp, encoding="utf-8") as f:
                h = json.load(f)
        except Exception:
            continue
        for v in h.get("versions", []):
            for rid in v.get("record_ids", []):
                rid_map.setdefault(rid, (supplier, v))
    try:
        sys.path.insert(0, PROJ_DIR)
        from refresh_oa_serial import refresh
        from generate_qa_agreement_v3 import normalize_oa_serial
    except Exception as e:
        log("OA流水号刷新模块加载失败: " + str(e))
        return
    refreshed = 0
    for rec in records:
        f = rec.get("fields", {})
        oa_val = get_val(f.get("OA审批流水号", ""))
        if not oa_val:
            continue
        item = rid_map.get(rec["id"])
        if not item:
            continue
        supplier, version = item
        current = normalize_oa_serial(version.get("oa_serial") or version.get("oa_serial_temp") or "")
        target = normalize_oa_serial(oa_val)
        if current == target:
            continue
        ok, msg = refresh(supplier, oa_val)
        log(("✅ OA流水号刷新: " if ok else "❌ OA流水号刷新失败: ") + supplier + " | " + msg)
        if ok:
            refreshed += 1
    if refreshed:
        log(f"本轮自动刷新流水号 {refreshed} 个供应商")


def main():
    log("=== 已确认供应商监控 ===")
    try:
        r = requests.get("https://oapi.dingtalk.com/gettoken?appkey=" + DT_APP_KEY + "&appsecret=" + DT_APP_SECRET, timeout=10)
        token = r.json()["access_token"]
    except Exception as e:
        log("Token失败: " + str(e))
        return

    headers = {"x-acs-dingtalk-access-token": token, "Content-Type": "application/json", "operatorId": DT_OPERATOR_ID}
    url = "https://api.dingtalk.com/v1.0/notable/bases/" + DT_BASE_ID + "/sheets/" + DT_SHEET2_ID + "/records?maxResults=100"
    r = requests.get(url, headers=headers, timeout=10)
    records = r.json().get("records", [])
    auto_refresh_oa_serials(records)
    state = load_state()
    processed = set(state.get("processed_ids", []))
    new_confirmed = []

    for rec in records:
        status = get_val(rec.get("fields", {}).get("节点状态", ""))
        if status in ("已确认供应商",) and rec["id"] not in processed:
            new_confirmed.append(rec)

    if not new_confirmed:
        log("无新增已确认")
        return

    log("发现 " + str(len(new_confirmed)) + " 条新确认")

    # 构建物料名称查询表 {编码: 名称}
    _names_by_code = {}
    for rec in records:
        f = rec.get("fields", {})
        _c = ""
        _v = f.get("物料编码", "")
        if isinstance(_v, dict): _c = _v.get("text", "")
        else: _c = str(_v) if _v else ""
        _n = ""
        _v2 = f.get("物料名称", "")
        if isinstance(_v2, dict): _n = _v2.get("text", "")
        else: _n = str(_v2) if _v2 else ""
        if _c: _names_by_code[_c] = _n
    
    supplier_groups = {}
    for rec in new_confirmed:
        rid = rec["id"]
        f = rec.get("fields", {})
        platform = get_val(f.get("业务平台", ""))
        applicant = get_val(f.get("研发申请人", ""))
        code = get_val(f.get("物料编码", ""))
        supplier = get_linked_supplier(headers, f.get("分配供应商(生物)", ""))
        if not supplier:
            supplier = get_linked_supplier(headers, f.get("分配供应商(平台C)", ""))
        change_type = get_val(f.get("变更类型", ""))
        if change_type and not supplier:
            try:
                sys.path.insert(0, PROJ_DIR)
                from material_master import load_master
                _entry = load_master().get(code, {})
                _sup = (_entry.get("suppliers") or [None])[0]
                if _sup:
                    supplier = _sup
                    log(f"升版记录 {rid}: 从主数据继承供应商 {supplier}")
            except Exception as e:
                log("升版供应商推断失败: " + str(e))
        if not supplier:
            log("跳过 " + rid + ": 未填供应商（下一轮重试）")
            continue
        party_a = get_val(f.get("甲方联系人", ""))
        party_a_company = get_val(f.get("甲方公司", ""))
        oa_rec = get_val(f.get("OA审批流水号", ""))
        log("收集: " + code + " | 供应商=" + supplier)
        if supplier not in supplier_groups:
            supplier_groups[supplier] = {
                "platform": platform, "applicant": applicant, "codes": [],
                "std_files": [], "raw_files": [], "platform_email": "",
                "record_ids": [], "party_a_contacts": [], "party_a_companies": [],
                "oa_serials": [], "change_types": {}, "supplements": {},
                "record_fields_by_code": {}, "record_ids_by_code": {},
            }
        sg = supplier_groups[supplier]
        # 获取平台联系人邮箱
        if not sg.get("platform_email"):
            _contacts_path = os.path.join(PROJ_DIR, "平台联系人的对应关系.xlsx")
            if os.path.exists(_contacts_path):
                try:
                    import pandas as _pd
                    _df = _pd.read_excel(_contacts_path, engine="openpyxl")
                    _row = _df[_df["平台"] == platform]
                    if not _row.empty:
                        _email = str(_row.iloc[0].get("邮箱地址", "") or "")
                        if _email and _email != "nan":
                            sg["platform_email"] = _email
                            log("平台邮箱: " + platform + " -> " + _email)
                except Exception as e:
                    log("读取联系人失败: " + str(e))
        sg = supplier_groups[supplier]
        sg["codes"].append(code)
        sg["record_ids"].append(rid)
        if party_a:
            sg["party_a_contacts"].append(party_a)
        if party_a_company:
            sg["party_a_companies"].append(party_a_company)
        if oa_rec:
            sg["oa_serials"].append(oa_rec)
        if change_type:
            sg["change_types"][code] = change_type
        if change_type and get_val(f.get("物料信息补充", "")):
            sg["supplements"][code] = get_val(f.get("物料信息补充", ""))
        sg["record_fields_by_code"].setdefault(code, []).append(f)
        sg["record_ids_by_code"].setdefault(code, []).append(rid)

    for supplier, sg in supplier_groups.items():
        log("\n=== 生成协议: " + supplier + " ===")
        codes = sorted(set(sg["codes"]))
        log("物料: " + ", ".join(codes))
        safe_supplier = safe_name(supplier)
        party_a_companies_early = [c for c in sg.get("party_a_companies", []) if c]
        party_a_company_early = party_a_companies_early[0] if party_a_companies_early else ""
        tag_early = "平台C" if ("示例丙公司" in party_a_company_early or (not party_a_company_early and sg.get("platform") == "平台C")) else "生物"
        company_folder = "示例丙公司" if tag_early == "平台C" else "示例甲公司"
        supplier_dir = os.path.join(DATA_DIR, company_folder, safe_supplier)
        agreements_dir = os.path.join(supplier_dir, "agreements")
        materials_dir = os.path.join(supplier_dir, "materials")
        os.makedirs(agreements_dir, exist_ok=True)
        os.makedirs(materials_dir, exist_ok=True)

        history = load_history(supplier_dir)
        old_codes = history.get("all_material_codes", []) if history else []
        all_codes = sorted(set(old_codes) | set(codes))
        new_codes = [c for c in codes if c not in old_codes]
        upgrade_codes = [c for c in codes if sg.get("change_types", {}).get(c)]
        if not new_codes and not upgrade_codes:
            log("物料均已在历史中处理过，跳过: " + ", ".join(codes))
            processed.update(sg["record_ids"])
            continue

        # 汇总原辅料明细/质量标准（历史+新增）：优先归档，回退 QA_DIR
        names_by_code = dict(_names_by_code)
        if history:
            for v in history.get("versions", []):
                for md in v.get("material_details", []):
                    if md.get("code") and not names_by_code.get(md["code"]):
                        names_by_code[md["code"]] = md.get("name", "")
        # 明细文件里的中文名称用于匹配标准文件（标准文件名通常为中文）
        match_names = dict(names_by_code)
        try:
            sys.path.insert(0, PROJ_DIR)
            from generate_qa_agreement_v3 import parse_raw_material
            for rf in find_raw_files(all_codes):
                for m in parse_raw_material(rf):
                    code = m.get("物料编码")
                    cn = (m.get("物料中文名称") or "").strip()
                    if code and cn and cn not in ("", "/"):
                        match_names[code] = cn
        except Exception as e:
            log("中文名称补全失败: " + str(e))
        raw_files, std_map = collect_supplier_files(supplier_dir, history, all_codes, match_names)
        master = None
        change_summaries = {}
        if upgrade_codes:
            from qa_upgrade import (
                apply_record_to_master, upgrade_standard, find_latest_std_file,
                update_record_fields, summary_text,
            )
            from material_master import load_master, save_master, check_missing_fields, _norm
            from generate_qa_agreement_v3 import parse_quality_standard
            master = load_master()
            pending = {}
            for code in upgrade_codes:
                fields_list = sg["record_fields_by_code"].get(code, [])
                fields = fields_list[0] if fields_list else {}
                entry = apply_record_to_master(master, code, fields)
                if _norm(entry.get("name")):
                    names_by_code[code] = entry["name"]
                new_std = find_latest_std_file(QA_DIR, code, names_by_code.get(code, ""))
                if new_std:
                    std_map[code] = new_std
                    parsed = None
                    try:
                        parsed = parse_quality_standard(new_std)
                    except Exception as e:
                        log(f"标准解析失败 {code}: {e}")
                    entry, summary = upgrade_standard(master, code, new_std, parsed)
                    change_summaries[code] = summary
                    log(f"升版 {code} {entry.get('name', '')}: " + summary_text(summary))
                else:
                    change_summaries[code] = {
                        "standard_changed": False,
                        "std_version": entry.get("std_version", 1),
                        "added": [], "removed": [], "changed": [],
                    }
                missing = check_missing_fields(entry)
                if missing:
                    pending[code] = {
                        "missing": missing,
                        "rids": sg["record_ids_by_code"].get(code, []),
                    }
            save_master(master)
            if pending:
                for code, p in pending.items():
                    text = "缺 " + "、".join(p["missing"])
                    log(f"升版物料 {code} 信息不完整，待人工补全: {text}")
                    ok = update_record_fields(_cfg, p["rids"], {"待补全字段": text})
                    log(("已回写待补全字段" if ok else "待补全字段回写失败") + f" {code}")
                continue
        if not raw_files:
            if not (master and all(c in master for c in all_codes)):
                log("未找到原辅料明细且主数据不完整，跳过（下一轮重试）")
                continue
            log("未找到原辅料明细，使用物料主数据构建")
        if not std_map:
            log("未找到质量标准文件，跳过（下一轮重试）")
            continue

        # 甲方公司/联系人：人工确认优先；同一供应商多个值时取第一个并告警
        party_a_contacts = [c for c in sg["party_a_contacts"] if c]
        party_a_contact = party_a_contacts[0] if party_a_contacts else ""
        if len(set(party_a_contacts)) > 1:
            log("同一供应商存在多个甲方联系人，使用第一个: " + party_a_contact)
        party_a_companies = [c for c in sg["party_a_companies"] if c]
        party_a_company = party_a_companies[0] if party_a_companies else ""
        if len(set(party_a_companies)) > 1:
            log("同一供应商存在多个甲方公司，使用第一个: " + party_a_company)

        oa_serials = [s for s in sg["oa_serials"] if s]
        oa_serial = oa_serials[0] if oa_serials else "QA-" + datetime.now().strftime("%Y%m%d-%H%M%S")
        version = max([v["version"] for v in (history or {}).get("versions", [])], default=0) + 1
        platform = sg["platform"] or "平台D"
        if not sg["platform"] and master:
            for code in upgrade_codes:
                if master.get(code, {}).get("platform"):
                    platform = master[code]["platform"]
                    break
        tag = "平台C" if ("示例丙公司" in party_a_company or (not party_a_company and platform == "平台C")) else "生物"
        sys.path.insert(0, PROJ_DIR)
        from generate_qa_agreement_v3 import build_output_filename
        out_file = os.path.join(agreements_dir, build_output_filename(oa_serial, tag, supplier))

        try:
            sys.path.insert(0, PROJ_DIR)
            from batch_generate_v3 import batch_generate
            party_b_info = {}
            try:
                with open(os.path.join(supplier_dir, "supplier_info.json"), encoding="utf-8") as _f:
                    _si = json.load(_f)
                _latest = _si.get("latest", {})
                party_b_info = {k: _latest.get(k, "") for k in ("name", "address", "contact", "phone", "email", "date")}
            except Exception:
                pass
            party_b_info["name"] = supplier
            party_b_info["date"] = datetime.now().strftime("%Y年%m月%d日")
            out_file = batch_generate(
                raw_material_docx=raw_files,
                supplier_name=supplier,
                platform_name=platform,
                oa_serial=oa_serial,
                output_path=out_file,
                material_filter=all_codes,
                std_map=std_map,
                name_overrides=names_by_code,
                party_a_contact=party_a_contact,
                party_a_company=party_a_company,
                master_overrides={c: master[c] for c in all_codes} if master else None,
                party_b_info=party_b_info or None,
            )
        except Exception as e:
            log("生成失败（下一轮重试）: " + str(e))
            continue
        content_hash = ""
        content_lines = []
        try:
            sys.path.insert(0, PROJ_DIR)
            from generate_qa_agreement_v3 import snapshot_agreement_text, agreement_lines
            content_hash = snapshot_agreement_text(out_file)
            content_lines = agreement_lines(out_file)
        except Exception as e:
            log("内容快照失败: " + str(e))

        # 归档原辅料明细/质量标准（按版本）
        raw_saved, std_saved = [], []
        try:
            std_ver_dir = os.path.join(materials_dir, "质量标准_v" + str(version))
            os.makedirs(std_ver_dir, exist_ok=True)
            for i, rf in enumerate(raw_files, 1):
                dst = os.path.join(
                    materials_dir,
                    "新增原辅料明细_v" + str(version) + ("_" + str(i) if len(raw_files) > 1 else "") + ".docx",
                )
                shutil.copy2(rf, dst)
                raw_saved.append(os.path.relpath(dst, supplier_dir))
            for code, sp in std_map.items():
                dst = os.path.join(std_ver_dir, os.path.basename(sp))
                shutil.copy2(sp, dst)
                std_saved.append(os.path.relpath(dst, supplier_dir))
        except Exception as e:
            log("归档失败（不影响协议）: " + str(e))

        # 更新 history.json
        if not history:
            history = {
                "supplier_name": supplier,
                "platform": platform,
                "created_at": datetime.now().isoformat(),
                "versions": [],
                "all_material_codes": [],
            }
        version_entry = {
            "version": version,
            "oa_serial": oa_serial,
            "oa_serial_temp": oa_serial,
            "record_ids": sg["record_ids"],
            "content_hash": content_hash,
            "content_lines": content_lines,
            "company_tag": tag,
            "company_folder": company_folder,
            "material_count": len(all_codes),
            "new_material_codes": new_codes,
            "raw_files": raw_saved,
            "std_files": {code: os.path.relpath(sp, supplier_dir) for code, sp in std_map.items()},
            "output_file": os.path.relpath(out_file, supplier_dir),
            "timestamp": datetime.now().isoformat(),
        }
        if upgrade_codes:
            version_entry["change_type"] = "升版"
            version_entry["change_summaries"] = {c: change_summaries.get(c, {}) for c in upgrade_codes}
        history.setdefault("versions", []).append(version_entry)
        history["all_material_codes"] = all_codes
        save_history(supplier_dir, history)
        log("✅ 协议已生成: " + out_file)

        # 写待发邮件任务（统一到 data/email_tasks，由 send_email_tasks.py 发送）
        # 收件人以钉钉"甲方联系人"为准，未填时回退平台负责人邮箱
        supplier_email = get_email_by_reviewer(party_a_contact) or sg.get("platform_email", "")
        if supplier_email:
            try:
                task_dir = os.path.join(DATA_DIR, "email_tasks")
                os.makedirs(task_dir, exist_ok=True)
                task_file = os.path.join(
                    task_dir,
                    safe_supplier + "_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".json",
                )
                email_task = {
                    "to": supplier_email,
                    "subject": "企业质量协议签订 - " + supplier,
                    "body": "您好，\n\n附件为贵司与我司签订的企业质量保证协议，请查收并盖章回传。\n\n谢谢！",
                    "attachments": [out_file],
                    "supplier": supplier,
                    "platform": platform,
                    "oa_serial": oa_serial,
                    "record_ids": sg["record_ids"],
                    "time": datetime.now().isoformat(),
                }
                with open(task_file, "w", encoding="utf-8") as _tf:
                    json.dump(email_task, _tf, ensure_ascii=False, indent=2)
                log("邮件任务已保存: " + supplier_email + " | 记录数 " + str(len(sg["record_ids"])))
            except Exception as e:
                log("邮件任务保存异常（需人工检查）: " + str(e))
        else:
            log("未找到供应商邮箱，跳过邮件（协议仍已生成）")

        processed.update(sg["record_ids"])

    state["processed_ids"] = list(processed)
    save_state(state)
    log("完成")

if __name__ == "__main__":
    main()
