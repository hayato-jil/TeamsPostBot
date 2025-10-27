# -*- coding: utf-8 -*-
import json
import os
from typing import Any, Dict

DEFAULT_STATE_PATH = os.environ.get("STATE_JSON", "state.json")

def load_state(path: str = DEFAULT_STATE_PATH) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {"records": {}}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"records": {}}

def save_state(state: Dict[str, Any], path: str = DEFAULT_STATE_PATH) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def record_key(no: str, name: str, email: str, join_date_iso: str) -> str:
    # シンプルなキー：No-メール-入社日
    return f"{no}::{email}::{join_date_iso}"

def get_record(state: Dict[str, Any], key: str) -> Dict[str, Any]:
    return state.setdefault("records", {}).setdefault(key, {})

def set_flag(state: Dict[str, Any], key: str, flag: str, value: bool = True) -> None:
    rec = get_record(state, key)
    rec[flag] = value
