# -*- coding: utf-8 -*-
import os
from typing import Dict

def load_template(path: str) -> str:
    if not os.path.exists(path):
        raise FileNotFoundError(f"template not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def render_template(text: str, mapping: Dict[str, str]) -> str:
    out = text
    for k, v in mapping.items():
        out = out.replace(f"{{{{{k}}}}}", v if v is not None else "")
    return out
