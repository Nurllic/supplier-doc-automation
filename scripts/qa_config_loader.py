#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一加载 qa_config.json，并允许用环境变量覆盖敏感配置（密钥外置）。"""

import json
import os


def _default_path():
    if os.name != "nt":
        return "/app/scripts/qa_config.json"
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "qa_config.json")


def load_qa_config(path=None):
    path = path or os.environ.get("QA_CONFIG_PATH") or _default_path()
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    env_map = {
        "app_key": "DINGTALK_APP_KEY",
        "app_secret": "DINGTALK_APP_SECRET",
        "operator_id": "DINGTALK_OPERATOR_ID",
        "base_id": "DINGTALK_BASE_ID",
        "sheet2_id": "DINGTALK_SHEET2_ID",
        "webhook": "DINGTALK_WEBHOOK",
        "robot_secret": "DINGTALK_ROBOT_SECRET",
        "qa_dir": "DINGTALK_QA_DIR",
        "state_file": "DINGTALK_STATE_FILE",
        "base_record_url": "DINGTALK_BASE_RECORD_URL",
    }
    for key, env_name in env_map.items():
        value = os.environ.get(env_name)
        if value:
            cfg[key] = value
    return cfg
