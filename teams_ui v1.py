# -*- coding: utf-8 -*-
import os
import re
import time
from typing import Optional, Sequence

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# 永続プロファイル（初回ログイン後は再利用）
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"

def _launch():
    channel = os.environ.get("PLAYWRIGHT_CHANNEL", "msedge")  # "msedge" or "chromium"
    pw = sync_playwright().start()
    browser = pw.chromium.launch_persistent_context(
        user_data_dir=PERSISTENT_DIR,
        channel=channel if channel in ("msedge", "chrome") else None,
        headless=False,
        args=["--disable-dev-shm-usage"],
    )
    page = browser.new_page()
    return pw, browser, page

def open_teams_for_login():
    """
    初回サインイン用。手動でログイン→閉じるだけ。
    """
    pw, browser, page = _launch()
    try:
        page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=120_000)
        page.wait_for_load_state("networkidle", timeout=120_000)
        print("Teams を開きました。必要ならサインインしてください。このウィンドウを閉じれば保存されます。")
        # 十分待つ（手動ログイン時間）。必要に応じて延長してください。
        time.sleep(5)
        input("ログインが完了したら Enter を押してください...")
    finally:
        browser.close()
        pw.stop()

def _get_by_label_any(page, roles_and_names):
    """
    日本語/英語UIの両対応を狙って、複数候補で取得を試みる。
    roles_and_names: [(role, regex_pattern), ...]
    """
    for role, name_pat in roles_and_names:
        try:
            return page.get_by_role(role, name=re.compile(name_pat))
        except Exception:
            continue
    return None

def _ensure_ready(page):
    # SPAのロード完了待ち
    page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=120_000)
    page.wait_for_load_state("networkidle", timeout=120_000)
    # たまに古いUIが残るので薄く遅延
    time.sleep(1.0)

def create_group_chat_and_send_message(*, admin_email: str, target_email: str,
                                       chat_name: str, message: str) -> None:
    """
    総務アカウントで、管理者＋対象者Aとのグループチャットを作成し、名称を設定して、messageを送信。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

        # 「新しいチャット」
        new_chat = _get_by_label_any(page, [
            ("button", r"新しいチャット"),
            ("button", r"New chat"),
        ])
        if new_chat is None:
            # 画面幅で「作成」UIが隠れている場合があるので、左ナビの「チャット」→表示
            chat_nav = _get_by_label_any(page, [
                ("link", r"チャット"),
                ("link", r"Chat"),
                ("button", r"チャット"),
                ("button", r"Chat"),
            ])
            if chat_nav:
                chat_nav.click()
                page.wait_for_load_state("networkidle", timeout=60_000)
                time.sleep(0.5)
                new_chat = _get_by_label_any(page, [
                    ("button", r"新しいチャット"),
                    ("button", r"New chat"),
                ])
        if not new_chat:
            raise RuntimeError("新しいチャットのボタンが見つかりませんでした。")

        new_chat.click()
        page.wait_for_timeout(500)

        # 宛先（To: / 宛先）
        to_box = None
        # combobox（宛先）
        for _ in range(2):
            to_box = _get_by_label_any(page, [
                ("combobox", r"宛先"),
                ("combobox", r"To"),
            ])
            if to_box:
                break
            page.wait_for_timeout(500)

        if not to_box:
            # 代替：テキストボックス/検索ボックス
            to_box = _get_by_label_any(page, [
                ("textbox", r"宛先|To|検索"),
            ])
        if not to_box:
            raise RuntimeError("宛先入力欄が見つかりませんでした。")

        # 宛先に admin と target を順に追加
        for addr in (admin_email, target_email):
            if not addr:
                continue
            to_box.click()
            to_box.fill("")  # クリア
            to_box.type(addr)
            page.keyboard.press("Enter")
            page.wait_for_timeout(400)

        # グループ名（チャット名）を追加（新規作成時に「グループ名を追加」ボタンがある想定）
        name_btn = _get_by_label_any(page, [
            ("button", r"グループ名を追加"),
            ("button", r"Add group name"),
            ("button", r"チャット名の編集"),
            ("button", r"Edit chat name"),
        ])
        if name_btn:
            name_btn.click()
            # 入力欄
            name_box = _get_by_label_any(page, [
                ("textbox", r"グループ名|チャット名|Name"),
            ])
            if name_box:
                name_box.fill(chat_name)
                # 保存
                save_btn = _get_by_label_any(page, [
                    ("button", r"保存"),
                    ("button", r"Save"),
                    ("button", r"適用"),
                    ("button", r"Apply"),
                ])
                if save_btn:
                    save_btn.click()
                else:
                    page.keyboard.press("Enter")

        # メッセージ入力欄を探す
        msg_box = _get_by_label_any(page, [
            ("textbox", r"メッセージ|メッセージを入力|Type a message"),
        ])
        if not msg_box:
            # Fallback: contenteditable検索
            msg_box = page.locator("[contenteditable='true']").nth(0)
        msg_box.click()
        msg_box.type(message)
        page.keyboard.press("Enter")
        page.wait_for_timeout(800)
        print("送信完了：テンプレ①を投稿しました。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
