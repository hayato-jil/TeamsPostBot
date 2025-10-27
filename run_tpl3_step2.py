# -*- coding: utf-8 -*-
import os
import sys

try:
    from dotenv import load_dotenv  # 任意（Use-DotEnvの保険）
    load_dotenv()
except Exception:
    pass

from templates_loader import load_template
from teams_ui import open_existing_chat_and_send_message_with_file

PDF_DIR = os.environ.get("PDF_DIR", ".\\pdf")
# Step3-2 では 1つだけ添付（テストはコレに固定）
PDF_FILE = os.environ.get(
    "PDF_FILE1",
    "【20250728最新版】入社時労務関連しおり.pdf"
)

def last_name_from_fullname(fullname: str) -> str:
    # 全角スペース優先、なければ半角スペースで切る（保険）
    if "　" in fullname:
        return fullname.split("　")[0].strip()
    if " " in fullname:
        return fullname.split(" ")[0].strip()
    return fullname.strip()

def main():
    if len(sys.argv) < 2:
        print("Usage: python run_tpl3_step2.py '＜フルネーム（全角スペース）＞'")
        sys.exit(1)

    full_name = sys.argv[1]
    chat_name = f"{full_name}(総務)"

    tpl3 = load_template("template3.txt")
    message = tpl3.replace("{{LAST_NAME}}", last_name_from_fullname(full_name))

    file_path = os.path.join(PDF_DIR, PDF_FILE)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"添付ファイルが見つかりません: {file_path}")

    print(f"対象チャット名: {chat_name}")
    print("送信メッセージ（プレビュー）:")
    print(message)
    print(f"\n添付ファイル: {file_path}\n")
    input("↑で問題なければ Enter で実行します（Teamsウィンドウが開きます）...")

    open_existing_chat_and_send_message_with_file(
        chat_name=chat_name,
        message=message,
        file_path=file_path,
    )

if __name__ == "__main__":
    main()
