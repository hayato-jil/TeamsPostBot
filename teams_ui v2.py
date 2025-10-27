# -*- coding: utf-8 -*-
import os
import re
import time
from typing import Optional, Sequence

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page, Locator

# 永続プロファイル（初回ログイン後は再利用）
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"

_DEFAULT_TIMEOUT = 120_000  # 120s（重いとき用に長め）
_SHORT_PAUSE = 500  # ms

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
        page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=_DEFAULT_TIMEOUT)
        page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)
        print("Teams を開きました。必要ならサインインしてください。このウィンドウを閉じれば保存されます。")
        time.sleep(5)
        input("ログインが完了したら Enter を押してください...")
    finally:
        browser.close()
        pw.stop()

# ------- 小さく堅牢化ユーティリティ -------

def _ensure_ready(page: Page):
    page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=_DEFAULT_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)
    time.sleep(1.0)

def _get_by_role_any(page: Page, candidates: Sequence[tuple[str, str]]) -> Optional[Locator]:
    for role, name_pat in candidates:
        try:
            loc = page.get_by_role(role, name=re.compile(name_pat))
            # 存在チェック（実在しないと click でタイムアウトする）
            loc.first.wait_for(timeout=2_000)
            return loc
        except Exception:
            continue
    return None

def _query_any(page: Page, selectors: Sequence[str]) -> Optional[Locator]:
    for sel in selectors:
        loc = page.locator(sel)
        try:
            loc.first.wait_for(timeout=2_000)
            return loc
        except Exception:
            continue
    return None

def _find_to_field(page: Page) -> Optional[Locator]:
    """
    宛先（To）入力欄を多段で探す。
    テナント/言語差や画面幅差に対応するため、役割/ラベル/placeholder/attribute の順でフォールバック。
    """
    # 1) role=combobox / name=宛先|To
    loc = _get_by_role_any(page, [
        ("combobox", r"宛先"),
        ("combobox", r"To"),
    ])
    if loc: return loc

    # 2) role=textbox / name=宛先|To|検索
    loc = _get_by_role_any(page, [
        ("textbox", r"宛先|To|検索|Search"),
    ])
    if loc: return loc

    # 3) attribute系（placeholder / aria-label）
    loc = _query_any(page, [
        "input[placeholder*='宛先']",
        "input[placeholder*='To']",
        "input[aria-label*='宛先']",
        "input[aria-label*='To']",
        "div[role='combobox'][aria-label*='宛先']",
        "div[role='combobox'][aria-label*='To']",
        "div[contenteditable='true'][aria-label*='宛先']",
        "div[contenteditable='true'][aria-label*='To']",
        # 一部UIで「名前、メール、グループ」系のplaceholderが使われることがある
        "input[placeholder*='名前']",
        "input[placeholder*='Name']",
    ])
    if loc: return loc

    # 4) 最後の手：新規チャットパネル上の最初の searchable combobox
    loc = _query_any(page, [
        "div[role='combobox']",
        "[data-tid='people-picker-input']",
        "[data-tid='newChat-peoplePicker']",
    ])
    if loc: return loc

    return None

def _find_message_box(page: Page) -> Locator:
    # もっとも安定しているのは contenteditable なエディタ
    loc = _query_any(page, [
        "div[contenteditable='true'][role='textbox']",
        "div[contenteditable='true']",
        # 予備（メッセージ入力欄を示す aria-label）
        "div[role='textbox'][aria-label*='メッセージ']",
        "div[role='textbox'][aria-label*='message']",
    ])
    if loc:
        return loc
    # 苦肉の策：最後にページ内の contenteditable 最初の要素
    return page.locator("[contenteditable='true']").first

def _set_chat_name_if_available(page: Page, chat_name: str):
    # 新規作成時のみ表示されることがある「グループ名を追加」等のUIを探す
    name_btn = _get_by_role_any(page, [
        ("button", r"グループ名を追加"),
        ("button", r"Add group name"),
        ("button", r"チャット名の編集"),
        ("button", r"Edit chat name"),
        ("button", r"名前を追加"),
        ("button", r"Add name"),
    ])
    if not name_btn:
        return
    try:
        name_btn.click(timeout=10_000)
        name_box = _get_by_role_any(page, [
            ("textbox", r"グループ名|チャット名|Name"),
        ])
        if name_box:
            name_box.fill(chat_name)
            save_btn = _get_by_role_any(page, [
                ("button", r"保存"),
                ("button", r"Save"),
                ("button", r"適用"),
                ("button", r"Apply"),
                ("button", r"完了"),
                ("button", r"Done"),
            ])
            if save_btn:
                save_btn.click()
            else:
                page.keyboard.press("Enter")
            page.wait_for_timeout(_SHORT_PAUSE)
    except Exception:
        # 名前設定は必須ではないため、握りつぶして続行
        pass

# ------- メイン操作 -------

def create_group_chat_and_send_message(*, admin_email: str, target_email: str,
                                       chat_name: str, message: str) -> None:
    """
    総務アカウントで、管理者＋対象者Aとのグループチャットを作成し、名称を設定して、messageを送信。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

        # 左ナビの「チャット」を明示クリック（狭幅だと必要）
        chat_nav = _get_by_role_any(page, [
            ("link", r"チャット"),
            ("link", r"Chat"),
            ("button", r"チャット"),
            ("button", r"Chat"),
        ])
        if chat_nav:
            chat_nav.click()
            page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)
            page.wait_for_timeout(_SHORT_PAUSE)

        # 「新しいチャット」を開く
        new_chat = _get_by_role_any(page, [
            ("button", r"新しいチャット"),
            ("button", r"New chat"),
            ("link",   r"新しいチャット"),
            ("link",   r"New chat"),
        ])
        if not new_chat:
            # 画面の上部にペンアイコンだけのケースがある
            icon_try = _query_any(page, [
                "[data-tid='new-chat-button']",
                "[aria-label*='新しいチャット']",
                "[aria-label*='New chat']",
            ])
            if icon_try:
                new_chat = icon_try
        if not new_chat:
            raise RuntimeError("新しいチャットの入口が見つかりませんでした。")

        new_chat.click()
        page.wait_for_timeout(_SHORT_PAUSE)

        # 宛先欄を取得（多段フォールバック）
        to_box = None
        for _ in range(6):  # 数回リトライ（SPAで出現が遅いことがある）
            to_box = _find_to_field(page)
            if to_box:
                break
            page.wait_for_timeout(1000)
        if not to_box:
            raise RuntimeError("宛先入力欄が見つかりませんでした。")

        # 宛先に admin と target を順に追加（タイプ→Enter）
        for addr in (admin_email, target_email):
            if not addr:
                continue
            to_box.click()
            try:
                to_box.fill("")  # 入力をクリア
            except Exception:
                pass  # fill不可なdiv系でも動くように
            to_box.type(addr, delay=20)
            page.keyboard.press("Enter")
            page.wait_for_timeout(700)

        # 可能ならチャット名を設定（UIがある場合のみ）
        _set_chat_name_if_available(page, chat_name)

        # メッセージ入力欄を探して送信
        msg_box = _find_message_box(page)
        msg_box.click()
        msg_box.type(message, delay=10)
        page.keyboard.press("Enter")
        page.wait_for_timeout(1000)
        print("送信完了：テンプレ①を投稿しました。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
