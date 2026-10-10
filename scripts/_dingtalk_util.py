import json, os, requests
from qa_config_loader import load_qa_config

def load_config():
    return load_qa_config()

def get_dingtalk_token(cfg):
    r = requests.get(f"https://oapi.dingtalk.com/gettoken?appkey={cfg['app_key']}&appsecret={cfg['app_secret']}", timeout=10)
    return r.json()["access_token"]

def get_headers(cfg, token):
    return {
        "x-acs-dingtalk-access-token": token,
        "Content-Type": "application/json",
        "operatorId": cfg["operator_id"],
    }

def get_linked_supplier(headers, cfg, field_val):
    if not isinstance(field_val, dict):
        return str(field_val) if field_val else ""
    ids = field_val.get("linkedRecordIds", [])
    if not ids:
        return ""
    for sid in ["YQva3Vo", "9uU30VN"]:
        u = f"https://api.dingtalk.com/v1.0/notable/bases/{cfg['base_id']}/sheets/{sid}/records/{ids[0]}"
        try:
            r3 = requests.get(u, headers=headers, timeout=10)
            if r3.status_code == 200:
                lf = r3.json().get("fields", {})
                val = lf.get("供应商全称", "")
                if isinstance(val, dict):
                    return val.get("text", "")
                return str(val) if val else ""
        except:
            pass
    return ""

def get_val(field):
    if isinstance(field, dict):
        return field.get("name", field.get("text", ""))
    if isinstance(field, list):
        return "; ".join(str(v)[:30] for v in field)
    return str(field) if field else ""

def fetch_records(cfg):
    token = get_dingtalk_token(cfg)
    headers = get_headers(cfg, token)
    url = f"https://api.dingtalk.com/v1.0/notable/bases/{cfg['base_id']}/sheets/{cfg['sheet2_id']}/records?maxResults=100"
    r = requests.get(url, headers=headers, timeout=15)
    return r.json().get("records", []), headers
