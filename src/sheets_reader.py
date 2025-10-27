# -*- coding: utf-8 -*-
import os
import re
from datetime import datetime
from dateutil import tz
from typing import List, Dict, Any, Optional

import gspread
from google.oauth2.service_account import Credentials

JST = tz.gettz(os.environ.get("TZ", "Asia/Tokyo"))

def _service_account() -> gspread.Client:
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", ".\\sa_credentials.json")
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return gspread.authorize(credentials)

def _parse_date(s: str) -> Optional[datetime]:
    if not s:
        return None
    s = s.strip()
    # 例: 2025年10月16日 / 2025/10/16 / 2025-10-16
    s = s.replace("　", " ").replace(",", " ").replace("（", "(").replace("）", ")")
    jp = re.match(r"^\s*(\d{4})年(\d{1,2})月(\d{1,2})日", s)
    if jp:
        y, m, d = map(int, jp.groups())
        return datetime(y, m, d, 0, 0, 0, tzinfo=JST)
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.replace(tzinfo=JST)
        except Exception:
            pass
    return None

def _last_name(full_name: str) -> str:
    if not full_name:
        return ""
    # 全角スペース優先、なければ半角スペース
    if "　" in full_name:
        return full_name.split("　", 1)[0].strip()
    if " " in full_name:
        return full_name.split(" ", 1)[0].strip()
    return full_name.strip()

def fetch_rows() -> List[Dict[str, Any]]:
    ss_id = os.environ["SPREADSHEET_ID"]
    sheet_name = os.environ.get("SHEET_NAME", "入社管理テスト用シート")
    tab_name = os.environ.get("TAB_NAME", "入社情報")

    gc = _service_account()
    sh = gc.open_by_key(ss_id)
    ws = sh.worksheet(tab_name) if tab_name else sh.sheet1

    # 想定見出し: A:No. B:氏名 C:メール D:入社日 E:電子締結完了
    values = ws.get_all_values()
    rows = []
    if not values or len(values) < 2:
        return rows

    header = values[0]
    # 簡易マップ
    col = {name: idx for idx, name in enumerate(header)}

    def val(r, name):
        idx = col.get(name)
        return r[idx].strip() if idx is not None and idx < len(r) else ""

    for r in values[1:]:
        if not any(r):
            # 完全な空行はスキップ
            continue

        no = val(r, "No.")
        full = val(r, "氏名")
        email = val(r, "メール")
        join = val(r, "入社日")
        done = val(r, "電子締結完了")

        # ここを追加：氏名・メール・入社日が全て空ならプレースホルダ行としてスキップ
        if (full == "") and (email == "") and (join == ""):
            continue

        join_dt = _parse_date(join)
        last = _last_name(full)
        rows.append({
            "no": no,
            "full_name": full,
            "last_name": last,
            "email": email,
            "join_date_raw": join,
            "join_date": join_dt,  # datetime or None (JST)
            "e_sign_done": str(done).strip().lower() in ("true", "1", "yes", "y", "済", "完了"),
        })
    return rows

def days_until(dt: Optional[datetime]) -> Optional[int]:
    if not dt:
        return None
    today = datetime.now(JST).date()
    return (dt.date() - today).days
