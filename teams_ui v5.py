# -*- coding: utf-8 -*-
import os
import re
import time
from typing import Optional, Sequence

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page, Locator

# 永続プロファイル（初回ログイン後は再利用）
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"

_DEFAULT_TIMEOUT = 150_000  # 少し延長
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

# ------- ユーティリティ -------

def _ensure_ready(page: Page):
    page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=_DEFAULT_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)
    time.sleep(1.0)

def _get_by_role_any(page: Page, candidates: Sequence[tuple[str, str]]) -> Optional[Locator]:
    for role, name_pat in candidates:
        try:
            loc = page.get_by_role(role, name=re.compile(name_pat))
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
    # 1) role=combobox / name
    loc = _get_by_role_any(page, [
        ("combobox", r"宛先"),
        ("combobox", r"To"),
    ])
    if loc: return loc
    # 2) role=textbox
    loc = _get_by_role_any(page, [
        ("textbox", r"宛先|To|検索|Search"),
    ])
    if loc: return loc
    # 3) attribute系
    loc = _query_any(page, [
        "input[placeholder*='宛先']",
        "input[placeholder*='To']",
        "input[aria-label*='宛先']",
        "input[aria-label*='To']",
        "div[role='combobox'][aria-label*='宛先']",
        "div[role='combobox'][aria-label*='To']",
        "div[contenteditable='true'][aria-label*='宛先']",
        "div[contenteditable='true'][aria-label*='To']",
        "input[placeholder*='名前']",
        "input[placeholder*='Name']",
    ])
    if loc: return loc
    # 4) people picker の汎用セレクタ
    loc = _query_any(page, [
        "[data-tid='people-picker-input']",
        "[data-tid='newChat-peoplePicker']",
        "div[role='combobox']",
    ])
    if loc: return loc
    return None

def _find_message_box(page: Page) -> Locator:
    loc = _query_any(page, [
        "div[contenteditable='true'][role='textbox']",
        "div[contenteditable='true']",
        "div[role='textbox'][aria-label*='メッセージ']",
        "div[role='textbox'][aria-label*='message']",
    ])
    if loc:
        return loc
    return page.locator("[contenteditable='true']").first

def _set_chat_name_if_available(page: Page, chat_name: str):
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
        pass  # 任意のため握りつぶし

def _recipient_chip_count(page: Page) -> int:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']",
        "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill",
        "[aria-label*='削除'] span.pill",
        "[aria-label*='Remove'] span.pill",
    ])
    if not loc:
        return 0
    try:
        return loc.count()
    except Exception:
        return 1

def _ensure_invite_if_needed(page: Page):
    """
    外部/未登録の相手で表示される「招待/追加」系ボタンがあれば押す。
    """
    btn = _get_by_role_any(page, [
        ("button", r"招待"),
        ("button", r"Invite"),
        ("button", r"追加"),
        ("button", r"Add"),
        ("button", r"参加"),
        ("button", r"Join"),
        ("button", r"送信"),  # 稀に「招待を送信」
        ("button", r"Send invite"),
    ])
    if btn:
        try:
            btn.click(timeout=5_000)
            page.wait_for_timeout(800)
        except Exception:
            pass

def _focus_to_field(page: Page, to_box: Optional[Locator]) -> Optional[Locator]:
    """
    宛先欄をクリックできない/消えた場合のフォールバック。
    people-pickerコンテナをクリック→Tabでフォーカス。
    """
    if to_box:
        try:
            to_box.click(timeout=2_000)
            return to_box
        except Exception:
            pass
    container = _query_any(page, [
        "[data-tid='people-picker']",
        "[data-tid='newChat-peoplePicker']",
        "div[role='combobox']",
    ])
    if container:
        try:
            container.click(timeout=2_000)
            page.keyboard.press("Tab")
            page.wait_for_timeout(200)
        except Exception:
            pass
    # 再取得
    return _find_to_field(page)

def _clear_field(page: Page, box: Locator):
    """
    Ctrl+A → Backspace で確実にクリア。
    """
    try:
        box.click()
        page.keyboard.press("Control+A")
        page.wait_for_timeout(50)
        page.keyboard.press("Backspace")
    except Exception:
        try:
            box.fill("")
        except Exception:
            pass

def _click_matching_option(page: Page, address: str) -> bool:
    """
    候補の option 群から、メールアドレス一致のものを優先クリック。
    見つからなければ False。
    """
    # 候補の role=option を広めに取得
    listbox = _query_any(page, [
        "[role='listbox']",
        "[data-tid*='people-picker'] [role='listbox']",
        "[id*='Dropdown'] [role='listbox']",
    ])
    options = (listbox.locator("[role='option']") if listbox else _query_any(page, ["[role='option']"]))
    if not options:
        return False
    try:
        count = options.count()
        target_idx = None
        for i in range(count):
            txt = options.nth(i).inner_text(timeout=1000)
            if address.lower() in txt.lower():
                target_idx = i
                break
        if target_idx is None and count > 0:
            target_idx = 0  # 最初の候補にフォールバック
        if target_idx is not None:
            options.nth(target_idx).click()
            return True
    except Exception:
        pass
    return False

def _add_recipient(page: Page, to_box: Locator, address: str):
    """
    宛先にメールを入力し、候補を確定（選択）する。
    """
    if not address:
        return
    before = _recipient_chip_count(page)

    # フォーカス確保＆クリア
    to_box = _focus_to_field(page, to_box)
    if not to_box:
        raise RuntimeError("宛先欄へフォーカスできませんでした。")
    _clear_field(page, to_box)

    # 入力
    to_box.type(address, delay=20)
    page.wait_for_timeout(600)

    # 候補クリック優先（メール一致）、なければ Down→Enter
    picked = _click_matching_option(page, address)
    if not picked:
        try:
            page.keyboard.press("ArrowDown")
            page.wait_for_timeout(120)
            page.keyboard.press("Enter")
        except Exception:
            pass

    # 念のためEnter追撃
    page.wait_for_timeout(250)
    page.keyboard.press("Enter")

    # 反映確認＋招待ボタン対処
    for _ in range(6):
        page.wait_for_timeout(300)
        _ensure_invite_if_needed(page)
        after = _recipient_chip_count(page)
        if after > before:
            return  # 追加成功

    # 最後の手：optionをとにかくクリック
    option = _query_any(page, [
        "[role='option']",
        "[data-tid*='people-picker'] [role='option']",
        "[id*='Dropdown'] [role='option']",
    ])
    if option:
        option.first.click()
        page.wait_for_timeout(300)

def _type_multiline(page: Page, box: Locator, text: str):
    """
    Teamsのリッチエディタで改行を保持して入力。
    各行を type し、行間は Shift+Enter を挟む。送信は別途 Enter。
    """
    lines = text.splitlines()
    box.click()
    for i, line in enumerate(lines):
        if line:
            box.type(line, delay=10)
        if i < len(lines) - 1:
            box.press("Shift+Enter")

# ------- メイン操作 -------

def create_group_chat_and_send_message(*, admin_email: str, target_email: str,
                                       chat_name: str, message: str) -> None:
    """
    総務アカウントで、管理者＋対象者Aとのグループチャットを作成し、名称を設定して、messageを送信。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

        # 左ナビの「チャット」
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

        # 新しいチャット
        new_chat = _get_by_role_any(page, [
            ("button", r"新しいチャット"),
            ("button", r"New chat"),
            ("link",   r"新しいチャット"),
            ("link",   r"New chat"),
        ])
        if not new_chat:
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

        # 宛先欄 取得（出現が遅いことがあるためリトライ）
        to_box = None
        for _ in range(8):
            to_box = _find_to_field(page)
            if to_box:
                break
            page.wait_for_timeout(1000)
        if not to_box:
            raise RuntimeError("宛先入力欄が見つかりませんでした。")

        # 宛先：管理者
        _add_recipient(page, to_box, admin_email)

        # 宛先欄を再取得（1人目追加でDOMが変わることがある）
        to_box = _find_to_field(page) or _focus_to_field(page, to_box)
        if not to_box:
            to_box = _focus_to_field(page, None)
        if not to_box:
            raise RuntimeError("2人目追加のための宛先欄を再取得できませんでした。")

        # 宛先：対象者A
        _add_recipient(page, to_box, target_email)

        # チャット名（任意）
        _set_chat_name_if_available(page, chat_name)

        # メッセージ送信（改行保持）
        msg_box = _find_message_box(page)
        _type_multiline(page, msg_box, message)
        page.keyboard.press("Enter")  # 最後に送信
        page.wait_for_timeout(1200)
        print("送信完了：テンプレ①を投稿しました。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
