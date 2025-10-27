# -*- coding: utf-8 -*-
import os
import argparse
from datetime import datetime
from dateutil import tz

from sheets_reader import fetch_rows, days_until
from templates_loader import load_template, render_template
from state_store import load_state, record_key

JST = tz.gettz(os.environ.get("TZ", "Asia/Tokyo"))

def decide_actions(row):
    """
    ここではまだTeams操作はしない。
    何をやるべきか（候補）を返すだけ。
    必須機能の範囲：
      - 4〜5: チャット作成 + テンプレ①送信
      - 7   : テンプレ③ + PDF3添付（条件: 電子締結完了 + 入社日の30日前）
    """
    actions = []
    # 4-5: チャット作成＆テンプレ①（v1では常に候補、実際はstateで未実行だけに絞る予定）
    actions.append("CREATE_CHAT_AND_SEND_TPL1")

    # 7: 電子締結完了 かつ 30日前
    du = days_until(row["join_date"])
    if row["e_sign_done"] and du is not None and du <= 30:
        actions.append("SEND_TPL3_WITH_PDFS")
    return actions

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="実行せずに計画のみ表示")
    args = parser.parse_args()

    # テンプレ読み込み
    tpl1_path = os.environ.get("TPL1_PATH", "templates\\template1.txt")
    tpl3_path = os.environ.get("TPL3_PATH", "templates\\template3.txt")
    tpl1 = load_template(tpl1_path)
    tpl3 = load_template(tpl3_path)

    rows = fetch_rows()
    state = load_state()  # まだ使わないが読み込みだけ

    print("=== ドライラン: 実行候補一覧 ===")
    print(f"時刻: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S %Z')}")
    print(f"行数: {len(rows)}")
    print("")

    for r in rows:
        # キー（後でstate管理に使う）
        join_iso = r["join_date"].strftime("%Y-%m-%d") if r["join_date"] else ""
        rk = record_key(r.get("no",""), r["full_name"], r["email"], join_iso)

        # 変数置換の準備
        mapping = {
            "FULL_NAME": r["full_name"],
            "LAST_NAME": r["last_name"],
            "JOIN_DATE": join_iso,
            "EMAIL": r["email"],
        }
        t1 = render_template(tpl1, mapping)
        t3 = render_template(tpl3, mapping)

        actions = decide_actions(r)

        print(f"[No.{r.get('no','')}] {r['full_name']} / {r['email']} / 入社日:{r['join_date_raw']} / e署名:{r['e_sign_done']}")
        print(f"  key: {rk}")
        du = days_until(r["join_date"])
        print(f"  30日前まであと: {du if du is not None else 'N/A'} 日")
        if not actions:
            print("  実行なし")
        else:
            for a in actions:
                if a == "CREATE_CHAT_AND_SEND_TPL1":
                    print("  ▶ チャット作成＋テンプレ①送信 候補")
                    print("    --- テンプレ①(プレビュー) ---")
                    print(indent_block(t1, "    "))
                elif a == "SEND_TPL3_WITH_PDFS":
                    print("  ▶ テンプレ③＋PDF3添付 候補")
                    print("    --- テンプレ③(プレビュー) ---")
                    print(indent_block(t3, "    "))
                    print(f"    添付予定ディレクトリ: {os.environ.get('PDF_DIR','')}")
        print("-" * 60)

def indent_block(text: str, prefix: str) -> str:
    return "\n".join(prefix + line for line in text.splitlines())

if __name__ == "__main__":
    main()
