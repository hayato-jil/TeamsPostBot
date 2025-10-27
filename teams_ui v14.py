# -*- coding: utf-8 -*-
import os
import re
from typing import Optional, Sequence

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page, Locator

# ========= 可変チューニング値（.envで上書き可） =========
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"
_DEFAULT_TIMEOUT = 150_000  # 大域タイムアウト

# 候補ピルの最低待機/最大ポーリング時間（ミリ秒）
SUGGESTION_MIN_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MIN_WAIT_MS", "1400"))
SUGGESTION_MAX_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MAX_WAIT_MS", "6000"))

# 1人目ピル確定→2人目入力の“間”
BETWEEN_RECIPIENTS_PAUSE_MS = int(os.environ.get("TPL_BETWEEN_RECIPIENTS_PAUSE_MS", "250"))

# 2人目ピル確定→チャット名欄オープンの“間”
BEFORE_OPEN_CHATNAME_MS = int(os.environ.get("TPL_BEFORE_OPEN_CHATNAME_MS", "200"))

# 送信後の配信/既読アイコン（チェックマーク等）待機
DELIVERY_WAIT_MS = int(os.environ.get("TPL_DELIVERY_WAIT_MS", "20000"))  # 既定20秒

# =======================================================

def _launch():
    channel = os.environ.get("PLAYWRIGHT_CHANNEL", "msedge")  # "msedge"/"chrome" 推奨
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
        input("ログインが完了したら Enter を押してください...")
    finally:
        browser.close()
        pw.stop()

# ------- ユーティリティ -------

def _ensure_ready(page: Page):
    page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=_DEFAULT_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

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
    loc = _get_by_role_any(page, [("combobox", r"宛先"), ("combobox", r"To")])
    if loc: return loc
    loc = _get_by_role_any(page, [("textbox", r"宛先|To|検索|Search")])
    if loc: return loc
    loc = _query_any(page, [
        "input[placeholder*='宛先']", "input[placeholder*='To']",
        "input[aria-label*='宛先']", "input[aria-label*='To']",
        "div[role='combobox'][aria-label*='宛先']",
        "div[role='combobox'][aria-label*='To']",
        "div[contenteditable='true'][aria-label*='宛先']",
        "div[contenteditable='true'][aria-label*='To']",
        "input[placeholder*='名前']", "input[placeholder*='Name']",
        "[data-tid='people-picker-input']", "[data-tid='newChat-peoplePicker']",
        "div[role='combobox']",
    ])
    return loc

def _find_message_box(page: Page) -> Locator:
    loc = _query_any(page, [
        "div[contenteditable='true'][role='textbox']",
        "div[role='textbox'][aria-label*='メッセージ']",
        "div[role='textbox'][aria-label*='message']",
        "div[contenteditable='true']",
    ])
    return loc or page.locator("[contenteditable='true']").first

def _set_chat_name_if_available(page: Page, chat_name: str):
    if not chat_name:
        return
    name_btn = _get_by_role_any(page, [
        ("button", r"グループ名を追加"), ("button", r"Add group name"),
        ("button", r"チャット名の編集"), ("button", r"Edit chat name"),
        ("button", r"名前を追加"), ("button", r"Add name"),
    ])
    if not name_btn:
        return
    try:
        page.wait_for_timeout(BEFORE_OPEN_CHATNAME_MS)
        name_btn.click(timeout=10_000)
        name_box = _get_by_role_any(page, [("textbox", r"グループ名|チャット名|Name")])
        if name_box:
            name_box.fill(chat_name)
            save_btn = _get_by_role_any(page, [
                ("button", r"保存"), ("button", r"Save"),
                ("button", r"適用"), ("button", r"Apply"),
                ("button", r"完了"), ("button", r"Done"),
            ])
            if save_btn:
                save_btn.click()
            else:
                page.keyboard.press("Enter")
    except Exception:
        pass  # 任意UIなので失敗しても続行

def _recipient_chip_count(page: Page) -> int:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']", "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill", "[aria-label*='削除'] span.pill", "[aria-label*='Remove'] span.pill",
    ])
    if not loc:
        return 0
    try:
        return loc.count()
    except Exception:
        return 1

def _chip_exists(page: Page, address: str) -> bool:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']", "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill",
    ])
    if not loc:
        return False
    try:
        for i in range(loc.count()):
            txt = loc.nth(i).inner_text(timeout=800)
            if address.lower() in txt.lower():
                return True
    except Exception:
        pass
    return False

def _ensure_invite_if_needed(page: Page):
    btn = _get_by_role_any(page, [
        ("button", r"招待"), ("button", r"Invite"),
        ("button", r"追加"), ("button", r"Add"),
        ("button", r"参加"), ("button", r"Join"),
        ("button", r"送信"), ("button", r"Send invite"),
    ])
    if btn:
        try:
            btn.click(timeout=5_000)
            page.wait_for_timeout(300)
        except Exception:
            pass

def _focus_to_field(page: Page, to_box: Optional[Locator]) -> Optional[Locator]:
    if to_box:
        try:
            to_box.click(timeout=2_000)
            return to_box
        except Exception:
            pass
    container = _query_any(page, ["[data-tid='people-picker']", "[data-tid='newChat-peoplePicker']", "div[role='combobox']"])
    if container:
        try:
            container.click(timeout=2_000)
            page.keyboard.press("Tab")
            page.wait_for_timeout(100)
        except Exception:
            pass
    return _find_to_field(page)

# ---- 候補関連 ----
def _latest_listbox(page: Page) -> Optional[Locator]:
    boxes = page.locator("[role='listbox']")
    try:
        n = boxes.count()
        if n == 0:
            boxes = page.locator("[data-tid*='people-picker'] [role='listbox'], [id*='Dropdown'] [role='listbox']")
            n = boxes.count()
        if n == 0:
            return None
        for i in range(n - 1, -1, -1):
            lb = boxes.nth(i)
            try:
                if lb.is_visible():
                    return lb
            except Exception:
                continue
        return boxes.nth(n - 1)
    except Exception:
        return None

def _wait_for_suggestions(page: Page, *, min_wait_ms: int, max_wait_ms: int) -> bool:
    page.wait_for_timeout(min_wait_ms)
    elapsed, step = 0, 100
    while elapsed < max_wait_ms:
        lb = _latest_listbox(page)
        if lb:
            try:
                opts = lb.locator("[role='option']")
                if opts.count() > 0 and opts.first.is_visible():
                    return True
            except Exception:
                pass
        page.wait_for_timeout(step)
        elapsed += step
    return False

def _click_matching_option(page: Page, address: str) -> bool:
    lb = _latest_listbox(page)
    if not lb:
        return False
    options = lb.locator("[role='option']")
    try:
        count = options.count()
        if count == 0:
            return False
        target_idx = None
        for i in range(count):
            opt = options.nth(i)
            try:
                if not opt.is_visible():
                    continue
                txt = opt.inner_text(timeout=800)
            except Exception:
                continue
            if address.lower() in txt.lower():
                target_idx = i
                break
        if target_idx is None:
            for i in range(count):
                try:
                    if options.nth(i).is_visible():
                        target_idx = i
                        break
                except Exception:
                    continue
        if target_idx is None:
            target_idx = 0
        opt = options.nth(target_idx)
        try:
            opt.scroll_into_view_if_needed(timeout=2_000)
        except Exception:
            pass
        try:
            opt.hover(timeout=1_000)
        except Exception:
            pass
        try:
            opt.click(timeout=2_000)
        except Exception:
            opt.click(timeout=2_000, force=True)
        return True
    except Exception:
        return False

def _light_clear_typing_area(page: Page):
    try:
        page.keyboard.type(" ")
        page.wait_for_timeout(20)
        page.keyboard.press("Backspace")
    except Exception:
        pass

def _add_recipient(page: Page, to_box: Locator, address: str):
    if not address:
        return
    if _chip_exists(page, address):
        return

    before = _recipient_chip_count(page)
    to_box = _focus_to_field(page, to_box)
    if not to_box:
        raise RuntimeError("宛先欄へフォーカスできませんでした。")

    _light_clear_typing_area(page)
    to_box.type(address, delay=16)

    _wait_for_suggestions(page, min_wait_ms=SUGGESTION_MIN_WAIT_MS, max_wait_ms=SUGGESTION_MAX_WAIT_MS)
    clicked = _click_matching_option(page, address)

    for _ in range(10):
        page.wait_for_timeout(200)
        _ensure_invite_if_needed(page)
        after = _recipient_chip_count(page)
        if after > before or _chip_exists(page, address):
            return
    if not clicked:
        lb = _latest_listbox(page)
        if lb:
            opt = lb.locator("[role='option']").first
            try:
                opt.scroll_into_view_if_needed(timeout=2_000)
                opt.hover(timeout=1_000)
                opt.click(timeout=2_000, force=True)
            except Exception:
                pass

# ---- メッセージ入力 ----
def _type_multiline(page: Page, box: Locator, text: str):
    lines = text.splitlines()
    box.click()
    for i, line in enumerate(lines):
        if line:
            box.type(line, delay=10)
        if i < len(lines) - 1:
            box.press("Shift+Enter")

# ---- 送信ボタン＋投稿完了（チェックマーク）待ち ----
_STATUS_SELECTORS = [
    "[data-icon-name*='CheckMark']", "[data-icon-name*='Checkmark']",
    "svg[aria-label*='既読']", "svg[aria-label*='送信済み']", "svg[aria-label*='配信済み']",
    "svg[aria-label*='Read']", "svg[aria-label*='Delivered']", "svg[aria-label*='Seen']",
    "[aria-label*='既読']", "[aria-label*='送信済み']", "[aria-label*='配信済み']",
    "[aria-label*='Read']", "[aria-label*='Delivered']", "[aria-label*='Seen']",
]

def _delivery_icon_count(page: Page) -> int:
    c = 0
    for sel in _STATUS_SELECTORS:
        try:
            c += page.locator(sel).count()
        except Exception:
            pass
    return c

def _click_send_button(page: Page) -> bool:
    send_btn = _get_by_role_any(page, [("button", r"送信"), ("button", r"Send")]) or _query_any(page, [
        "[data-tid='send-message-button']", "[data-tid='send-button']",
        "button[aria-label*='送信']", "button[aria-label*='Send']",
    ])
    if not send_btn:
        return False
    try:
        send_btn.click(timeout=3_000)
        return True
    except Exception:
        try:
            send_btn.click(timeout=3_000, force=True)
            return True
        except Exception:
            return False

def _wait_delivery_increase(page: Page, base_count: int, timeout_ms: int) -> bool:
    waited, step = 0, 200
    while waited < timeout_ms:
        try:
            if _delivery_icon_count(page) > base_count:
                return True
        except Exception:
            pass
        page.wait_for_timeout(step)
        waited += step
    return False

# ------- メイン操作（テンプレ①用：既存） -------
def create_group_chat_and_send_message(*, admin_email: str, target_email: str,
                                       chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

        chat_nav = _get_by_role_any(page, [("link", r"チャット"), ("link", r"Chat"), ("button", r"チャット"), ("button", r"Chat")])
        if chat_nav:
            chat_nav.click()
            page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        new_chat = _get_by_role_any(page, [
            ("button", r"新しいチャット"), ("button", r"New chat"),
            ("link", r"新しいチャット"), ("link", r"New chat"),
        ]) or _query_any(page, [
            "[data-tid='new-chat-button']", "[aria-label*='新しいチャット']", "[aria-label*='New chat']",
        ])
        if not new_chat:
            raise RuntimeError("新しいチャットの入口が見つかりませんでした。")

        new_chat.click()

        # 宛先欄（リトライ：200ms刻み×最大2秒）
        to_box = None
        for _ in range(10):
            to_box = _find_to_field(page)
            if to_box:
                break
            page.wait_for_timeout(200)
        if not to_box:
            raise RuntimeError("宛先入力欄が見つかりませんでした。")

        # 宛先：管理者
        _add_recipient(page, to_box, os.environ.get("ADMIN_EMAIL",""))

        # 次の入力へ素早く移行（可変）
        page.wait_for_timeout(BETWEEN_RECIPIENTS_PAUSE_MS)

        # 宛先欄を素早く再取得（200ms刻み×最大2秒）
        to_box = None
        for _ in range(10):
            to_box = _find_to_field(page) or _focus_to_field(page, to_box)
            if to_box:
                break
            page.wait_for_timeout(200)
        if not to_box:
            raise RuntimeError("2人目追加のための宛先欄を再取得できませんでした。")

        # 宛先：対象者A（呼び出し元から渡すようにすべきだが既存互換のため引数で受け取り）
        # ここは既存 run_tpl1.py 側から target_email を渡しています
        _add_recipient(page, to_box, os.environ.get("TARGET_EMAIL",""))

        _set_chat_name_if_available(page, os.environ.get("CHAT_NAME",""))

        # メッセージ入力→送信
        msg_box = _find_message_box(page)
        _type_multiline(page, msg_box, os.environ.get("TPL1_MESSAGE",""))

        base_icons = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base_icons, timeout_ms=DELIVERY_WAIT_MS):
            print("送信完了：テンプレ①を投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()

# ------- 新規追加：既存チャットを開いて本文送信（テンプレ③の1stステップ） -------
def open_existing_chat_and_send_message(*, chat_name: str, message: str) -> None:
    """
    左上の検索ボックスから `chat_name` を検索 → 一致するチャットを開く → 本文 `message` を送信。
    改行保持 & 送信後は配信/既読アイコンが増えるまで待機。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

        # チャットナビへ
        chat_nav = _get_by_role_any(page, [("link", r"チャット"), ("link", r"Chat"), ("button", r"チャット"), ("button", r"Chat")])
        if chat_nav:
            chat_nav.click()
            page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        # 検索欄にフォーカス → チャット名で検索
        search_box = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
            "input[type='search']", "input[placeholder*='検索']", "input[placeholder*='Search']",
        ])
        if not search_box:
            raise RuntimeError("検索ボックスが見つかりませんでした。")
        search_box.click()
        search_box.fill(chat_name)
        search_box.press("Enter")

        # 検索結果から該当チャットをクリック（タイトルに完全一致優先→部分一致）
        result = page.locator(f"text={chat_name}").first
        try:
            result.wait_for(timeout=10_000)
        except Exception:
            # 部分一致フォールバック
            result = page.locator("div,span").filter(has_text=re.compile(re.escape(chat_name))).first
            result.wait_for(timeout=10_000)
        result.click()

        # メッセージ入力 → 送信
        msg_box = _find_message_box(page)
        _type_multiline(page, msg_box, message)

        base_icons = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base_icons, timeout_ms=DELIVERY_WAIT_MS):
            print("送信完了：本文のみ投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
