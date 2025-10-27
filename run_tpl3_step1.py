# -*- coding: utf-8 -*-
import os
import sys
from dotenv import load_dotenv

from templates_loader import load_template
from teams_ui import open_existing_chat_and_send_message

def split_last_name(full_name: str) -> str:
    # 全角/半角スペースどちらでもOK
    for sep in ["　", " "]:
        if sep in full_name:
            return full_name.split(sep)[0].strip()
    # スペースが無い場合は先頭1文字だけ姓とみなす（暫定）
    return full_name[:1]

def main():
    load_dotenv()
    if len(sys.argv) >= 2:
        full_name = sys.argv[1]
    else:
        full_name = input("フルネーム（例: 山田　太郎）を入力してください: ").strip()
    if not full_name:
        print("フルネームが空です。終了します。")
        return

    chat_name = f"{full_name}(総務)"
    last_name = split_last_name(full_name)

    tpl3 = load_template("template3.txt")
    message = tpl3.replace("{{LAST_NAME}}", last_name)

    print("対象チャット名:", chat_name)
    print("送信メッセージ（プレビュー）:\n" + message + "\n")
    input("↑で問題なければ Enter で実行します（Teamsウィンドウが開きます）...")

    open_existing_chat_and_send_message(chat_name=chat_name, message=message)

if __name__ == "__main__":
    main()
