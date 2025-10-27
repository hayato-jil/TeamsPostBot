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
    loc = _get_by_role_any(page, [
        ("combobox", r"宛先"),
        ("combobox", r"To"),
    ])
    if loc: return loc
    loc = _get_by_role_any(page, [
        ("textbox", r"宛先|To|検索|Search"),
    ])
    if loc: return loc
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
        pass

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

def _chip_exists(page: Page, address: str) -> bool:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']",
        "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill",
    ])
    if not loc:
        return False
    try:
        count = loc.count()
        for i in range(count):
            txt = loc.nth(i).inner_text(timeout=800)
            if address.lower() in txt.lower():
                return True
    except Exception:
        pass
    return False

def _ensure_invite_if_needed(page: Page):
    btn = _get_by_role_any(page, [
        ("button", r"招待"),
        ("button", r"Invite"),
        ("button", r"追加"),
        ("button", r"Add"),
        ("button", r"参加"),
        ("button", r"Join"),
        ("button", r"送信"),
        ("button", r"Send invite"),
    ])
    if btn:
        try:
            btn.click(timeout=5_000)
            page.wait_for_timeout(800)
        except Exception:
            pass

def _focus_to_field(page: Page, to_box: Optional[Locator]) -> Optional[Locator]:
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

def _wait_for_suggestions(page: Page, *, min_wait_ms: int = 2200, max_wait_ms: int = 7000) -> bool:
    page.wait_for_timeout(min_wait_ms)
    elapsed = 0
    step = 100
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
        page.wait_for_timeout(30)
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

    to_box.type(address, delay=18)

    _wait_for_suggestions(page, min_wait_ms=2200, max_wait_ms=7000)
    clicked = _click_matching_option(page, address)

    for _ in range(10):
        page.wait_for_timeout(250)
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
    # 既読/送信済みなどのアイコン（環境により異なるためマルチパターン）
    "[data-icon-name*='CheckMark']",
    "[data-icon-name*='Checkmark']",
    "svg[aria-label*='既読']",
    "svg[aria-label*='送信済み']",
    "svg[aria-label*='配信済み']",
    "svg[aria-label*='Read']",
    "svg[aria-label*='Delivered']",
    "svg[aria-label*='Seen']",
    # 一部のUIではspanやiタグにaria-labelが付く
    "[aria-label*='既読']",
    "[aria-label*='送信済み']",
    "[aria-label*='配信済み']",
    "[aria-label*='Read']",
    "[aria-label*='Delivered']",
    "[aria-label*='Seen']",
]

def _delivery_icon_count(page: Page) -> int:
    count = 0
    for sel in _STATUS_SELECTORS:
        try:
            count += page.locator(sel).count()
        except Exception:
            pass
    return count

def _click_send_button(page: Page):
    send_btn = _get_by_role_any(page, [
        ("button", r"送信"),
        ("button", r"Send"),
    ]) or _query_any(page, [
        "[data-tid='send-message-button']",
        "[data-tid='send-button']",
        "button[aria-label*='送信']",
        "button[aria-label*='Send']",
    ])
    if send_btn:
        try:
            send_btn.click(timeout=3_000)
            return True
        except Exception:
            try:
                send_btn.click(timeout=3_000, force=True)
                return True
            except Exception:
                pass
    return False

def _wait_delivery_increase(page: Page, base_count: int, timeout_ms: int = 20_000) -> bool:
    """
    送信前に取得した配信/既読アイコン数(base_count)よりも数が増えるまで待つ。
    """
    waited = 0
    step = 200
    while waited < timeout_ms:
        try:
            if _delivery_icon_count(page) > base_count:
                return True
        except Exception:
            pass
        page.wait_for_timeout(step)
        waited += step
    return False

# ------- メイン操作 -------

def create_group_chat_and_send_message(*, admin_email: str, target_email: str,
                                       chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)

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

        new_chat = _get_by_role_any(page, [
            ("button", r"新しいチャット"),
            ("button", r"New chat"),
            ("link",   r"新しいチャット"),
            ("link",   r"New chat"),
        ]) or _query_any(page, [
            "[data-tid='new-chat-button']",
            "[aria-label*='新しいチャット']",
            "[aria-label*='New chat']",
        ])
        if not new_chat:
            raise RuntimeError("新しいチャットの入口が見つかりませんでした。")

        new_chat.click()
        page.wait_for_timeout(_SHORT_PAUSE)

        to_box = None
        for _ in range(8):
            to_box = _find_to_field(page)
            if to_box:
                break
            page.wait_for_timeout(1000)
        if not to_box:
            raise RuntimeError("宛先入力欄が見つかりませんでした。")

        _add_recipient(page, to_box, admin_email)

        to_box = _find_to_field(page) or _focus_to_field(page, to_box)
        if not to_box:
            to_box = _focus_to_field(page, None)
        if not to_box:
            raise RuntimeError("2人目追加のための宛先欄を再取得できませんでした。")

        _add_recipient(page, to_box, target_email)

        _set_chat_name_if_available(page, chat_name)

        msg_box = _find_message_box(page)
        _type_multiline(page, msg_box, message)

        # 送信前のアイコン数を取得
        base_icons = _delivery_icon_count(page)

        # 送信ボタンで送る（Enterは原則使わない）
        clicked = _click_send_button(page)
        if not clicked:
            try:
                page.keyboard.press("Control+Enter")
            except Exception:
                pass
            page.wait_for_timeout(300)
            try:
                page.keyboard.press("Enter")
            except Exception:
                pass

        # チェックマーク等が増えるまで待機（=配信/既読が付いたことを確認）
        if _wait_delivery_increase(page, base_icons, timeout_ms=20_000):
            print("送信完了：テンプレ①を投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。UI差異の可能性があります。")
        page.wait_for_timeout(800)

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
