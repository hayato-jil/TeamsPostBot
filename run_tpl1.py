# -*- coding: utf-8 -*-
import os
from templates_loader import load_template, render_template
from sheets_reader import fetch_rows
from dateutil import tz
from teams_ui import create_group_chat_and_send_message

def main():
    rows = fetch_rows()
    if not rows:
        print("対象行がありません。")
        return
    r = rows[0]  # まずは1件だけテスト送信
    full = r["full_name"]
    last = r["last_name"]
    email = r["email"]

    admin_email = os.environ.get("ADMIN_EMAIL", "")
    chat_name = f"{full}(総務)"

    tpl1_path = os.environ.get("TPL1_PATH", "templates\\template1.txt")
    tpl1 = load_template(tpl1_path)
    msg = render_template(tpl1, {
        "FULL_NAME": full,
        "LAST_NAME": last,
        "JOIN_DATE": r["join_date"].strftime("%Y-%m-%d") if r["join_date"] else "",
        "EMAIL": email,
    })

    print(f"対象: {full} / {email}")
    print(f"チャット名: {chat_name}")
    print("送信メッセージ（プレビュー）:")
    print(msg)
    input("↑で問題なければ Enter で実行します（Teams ウィンドウが開きます）...")

    create_group_chat_and_send_message(
        admin_email=admin_email,
        target_email=email,
        chat_name=chat_name,
        message=msg
    )

if __name__ == "__main__":
    main()
