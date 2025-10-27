# -*- coding: utf-8 -*-
import os
import sys
from dotenv import load_dotenv

from templates_loader import load_template
from teams_ui import open_existing_chat_and_send_message_with_files

def main():
    load_dotenv(override=True)
    if len(sys.argv) < 2:
        print("使い方: python run_tpl3_step3.py '<フルネーム>'")
        sys.exit(1)

    full_name = sys.argv[1].strip()
    last_name = full_name.split()[0].replace("　", " ").split(" ")[0]  # 姓だけ
    chat_name = f"{full_name}(総務)"

    # テンプレ③の読み込み
    msg = load_template("template3.txt").replace("{{LAST_NAME}}", last_name)

    # 添付ファイル（固定名・固定場所）
    base = r"C:\TeamsPostBot\pdf"
    files = [
        os.path.join(base, "就業規則_20240206（閲覧用）.pdf"),
        os.path.join(base, "【20250728最新版】入社時労務関連しおり.pdf"),
        os.path.join(base, "副業・兼業について基本ガイドライン（20240724）（閲覧用).pdf"),
    ]

    print(f"対象チャット名: {chat_name}")
    print("送信メッセージ（プレビュー）:")
    print(msg)
    print("\n添付予定ファイル:")
    for f in files:
        print(" -", f)

    input("\n↑で問題なければ Enter で実行します（Teamsウィンドウが開きます）...")

    open_existing_chat_and_send_message_with_files(
        chat_name=chat_name,
        message=msg,
        file_paths=files,
    )

if __name__ == "__main__":
    main()
