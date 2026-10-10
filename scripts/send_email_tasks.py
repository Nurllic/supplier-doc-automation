#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发送 email_tasks 中的邮件任务；发送成功后回写钉钉状态（协议已发送）并发送群通知"""

import base64, hashlib, hmac, json, os, re, subprocess, sys, time
from datetime import datetime

import requests

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qa_config_loader import load_qa_config


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def load_config():
    return load_qa_config()


def get_new_token(cfg):
    """获取 api.dingtalk.com 域名的新版 accessToken"""
    r = requests.post(
        "https://api.dingtalk.com/v1.0/oauth2/accessToken",
        json={"appKey": cfg["app_key"], "appSecret": cfg["app_secret"]},
        timeout=10,
    )
    data = r.json()
    if "accessToken" not in data:
        raise RuntimeError("获取钉钉Token失败: " + str(data))
    return data["accessToken"]


def update_dingtalk_status(cfg, record_ids):
    """把钉钉AI表格记录节点状态更新为 协议已发送"""
    if not record_ids:
        return True
    token = get_new_token(cfg)
    headers = {
        "x-acs-dingtalk-access-token": token,
        "Content-Type": "application/json",
        "operatorId": cfg["operator_id"],
    }
    records = [{"id": rid, "fields": {"节点状态": "协议已发送"}} for rid in record_ids]
    url = f"https://api.dingtalk.com/v1.0/notable/bases/{cfg['base_id']}/sheets/{cfg['sheet2_id']}/records"
    r = requests.put(url, headers=headers, json={"records": records}, timeout=15)
    ok = r.status_code == 200
    if not ok:
        log("钉钉状态回写失败: " + r.text[:300])
    return ok


def send_group_notification(cfg, supplier, to_addr, oa_serial):
    url = cfg.get("webhook", "")
    if not url:
        return False
    secret = cfg.get("robot_secret", "")
    if secret and secret.strip():
        timestamp = str(round(time.time() * 1000))
        sign_str = timestamp + "\n" + secret
        signature = base64.b64encode(
            hmac.new(secret.encode(), sign_str.encode(), digestmod=hashlib.sha256).digest()
        ).decode()
        sep = "&" if "?" in url else "?"
        url = url + sep + "timestamp=" + timestamp + "&sign=" + signature
    text = (
        "### ✅ 企业质量协议已发送\n\n"
        f"- **供应商：** {supplier}\n"
        f"- **收件人：** {to_addr}\n"
        f"- **流水号：** {oa_serial or '—'}\n"
        f"- **时间：** {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    )
    payload = {"msgtype": "markdown", "markdown": {"title": "企业质量协议已发送", "text": text}}
    try:
        r = requests.post(url, json=payload, timeout=10)
        return r.json().get("errcode") == 0
    except Exception as e:
        log("群通知发送失败: " + str(e))
        return False


def collect_task_files():
    """扫描 data/email_tasks 和 email-database/email_tasks（兼容旧路径）"""
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dirs = [
        os.path.join(base, "data", "email_tasks"),
        os.path.join(base, "email-database", "email_tasks"),
    ]
    files = []
    for d in dirs:
        if os.path.exists(d):
            for fn in sorted(os.listdir(d)):
                if fn.endswith(".json"):
                    files.append(os.path.join(d, fn))
    return files


def run_agently(args, cwd):
    cmd = "agently-cli " + " ".join(f'"{a}"' if " " in a else a for a in args)
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                            timeout=60, shell=True, cwd=cwd)
    if result.returncode != 0:
        return {"ok": False, "error": (result.stderr or "")[:300]}
    text = result.stdout
    brace_idx = text.find("{")
    if brace_idx < 0:
        return {"ok": False, "error": text[:300]}
    end_idx = text.rfind("}")
    if end_idx < 0:
        return {"ok": False, "error": text[:300]}
    try:
        return json.loads(text[brace_idx:end_idx + 1])
    except Exception as e:
        return {"ok": False, "error": str(e)}


def send_via_agently(to_addr, subject, body, attachments):
    """agently-cli message +send 为两阶段确认：第一次返回 confirmation_token，带 token 再调一次才真正发送。
    注意：正文含换行会经 cmd.exe 破坏确认一致性，发送前需压成单行。"""
    body = re.sub(r"\s*\n+\s*", " ", str(body)).strip()
    subject = re.sub(r"\s*\n+\s*", " ", str(subject)).strip()
    args = ["message", "+send", "--to", to_addr, "--subject", subject, "--body", body]
    cwd = os.path.dirname(os.path.abspath(attachments[0])) if attachments else os.path.expanduser("~")
    for att in attachments:
        if os.path.exists(att):
            rel = os.path.relpath(os.path.abspath(att), cwd)
            args.extend(["--attachment", rel])
        else:
            log(f"  ⚠️ 附件不存在: {att}")
    first = run_agently(args, cwd)
    if not first.get("ok"):
        return False, first.get("error", "")
    data = first.get("data", {})
    token = data.get("confirmation_token", "") if data.get("confirmation_required") else ""
    if token:
        args.extend(["--confirmation-token", token])
        second = run_agently(args, cwd)
        if not second.get("ok"):
            return False, second.get("error", "")
        if not second.get("data", {}).get("queued"):
            return False, "确认未完成：第二次调用仍返回 confirmation_required，邮件未入队"
    return True, ""


def send_one(cfg, fp):
    with open(fp, encoding="utf-8") as f:
        task = json.load(f)

    to_addr = task.get("to", "")
    subject = task.get("subject", "")
    body = task.get("body", "")
    attachments = task.get("attachments", [])
    log(f"发送邮件: {to_addr} | {subject[:40]}")
    for att in attachments:
        log(f"  附件: {att} (存在={os.path.exists(att)})")

    resolved = []
    for att in attachments:
        if not os.path.exists(att):
            alt = att.replace("email-database\\output", "data\\" + task.get("supplier", "") + "\\agreements")
            alt = alt.replace("email-database/output", "data/" + task.get("supplier", "") + "/agreements")
            if os.path.exists(alt):
                att = alt
        resolved.append(att)

    ok, err = send_via_agently(to_addr, subject, body, resolved)
    if not ok:
        log(f"❌ 发送失败（任务保留待重试）: {err}")
        return False

    log("✅ 邮件已进入发送队列")
    # 先删除任务避免重复发信，再回写状态/群通知（失败仅记录日志）
    try:
        os.remove(fp)
    except OSError:
        pass

    if not update_dingtalk_status(cfg, task.get("record_ids", [])):
        log("⚠️ 邮件已发送，但钉钉状态回写失败，请人工检查记录: " + ", ".join(task.get("record_ids", [])))
    if not send_group_notification(cfg, task.get("supplier", ""), to_addr, task.get("oa_serial", "")):
        log("⚠️ 群通知发送失败")

    time.sleep(2)
    return True


def main():
    cfg = load_config()
    files = collect_task_files()
    if not files:
        print("无邮件任务")
        return
    for fp in files:
        try:
            send_one(cfg, fp)
        except Exception as e:
            log("任务处理异常（保留任务）: " + str(e))
    print("完成")


if __name__ == "__main__":
    main()
