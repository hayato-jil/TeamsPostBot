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

# 添付待ち・リトライ
ATTACH_UPLOAD_TIMEOUT_MS = int(os.environ.get("ATTACH_UPLOAD_TIMEOUT_MS", "45000"))  # 最大45秒待つ
ATTACH_RETRIES = int(os.environ.get("ATTACH_RETRIES", "1"))  # 失敗時の再試行回数
ATTACH_FAIL_BEHAVIOR = os.environ.get("ATTACH_FAIL_BEHAVIOR", "send_without_file")  # "abort" | "send_without_file"

# デバッグログ
DEBUG = os.environ.get("DEBUG_LOG", "0") == "1"
def debug(msg: str):
    if DEBUG:
        print(f"[DEBUG] {msg}")

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

# ---- 候補関連（宛先ピッカー）----
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

# ====== 検索：サジェストから既存チャットを開く（Enter禁止） ======
def _find_search_box(page: Page) -> Optional[Locator]:
    debug("search: try Ctrl+E")
    for _ in range(3):
        try:
            page.keyboard.press("Control+e")
            page.wait_for_timeout(200)
        except Exception:
            pass
        loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
            "input[type='search']",
            "input[placeholder*='検索']", "input[placeholder*='Search']",
            "input[role='combobox'][type='search']",
            "div[role='combobox'] input[type='search']",
            "input[aria-label*='検索']", "input[aria-label*='Search']",
        ])
        if loc:
            try:
                loc.first.wait_for(timeout=800)
                debug("search: got by selector after Ctrl+E")
                return loc
            except Exception:
                pass
    debug("search: click top bar area then re-detect")
    try:
        page.mouse.move(400, 60)
        page.mouse.click(400, 60)
        page.wait_for_timeout(150)
    except Exception:
        pass
    loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
        "input[type='search']",
        "input[placeholder*='検索']", "input[placeholder*='Search']",
        "input[role='combobox'][type='search']",
        "div[role='combobox'] input[type='search']",
        "input[aria-label*='検索']", "input[aria-label*='Search']",
    ])
    if loc:
        debug("search: got by selector after top click")
        return loc
    debug("search: fallback '/' key then re-detect")
    try:
        page.keyboard.press("/")
        page.wait_for_timeout(120)
    except Exception:
        pass
    loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
        "input[type='search']",
        "input[placeholder*='検索']", "input[placeholder*='Search']",
        "input[role='combobox'][type='search']",
        "div[role='combobox'] input[type='search']",
        "input[aria-label*='検索']", "input[aria-label*='Search']",
    ])
    if loc:
        debug("search: got by selector after '/'")
        return loc
    debug("search: NOT FOUND")
    return None

def _is_ng_suggestion_text(txt: str) -> bool:
    t = txt.replace("\n", " ")
    ng_patterns = [
        r"Enter.?キー.*結果.*表示",
        r"結果.*表示",
        r"ユーザー.*招待",
        r"Invite.*to.*Teams",
    ]
    for pat in ng_patterns:
        if re.search(pat, t):
            return True
    return False

def _get_group_chat_section(panel: Locator) -> Optional[Locator]:
    headers = panel.locator("text=/グループ.?チャット/i, text=/Group.?chat/i")
    try:
        if headers.count() > 0:
            for i in range(headers.count()):
                h = headers.nth(i)
                if h.is_visible():
                    return h
    except Exception:
        pass
    return None

def _click_center(page: Page, el: Locator) -> bool:
    try:
        box = el.bounding_box()
        if not box:
            return False
        page.mouse.move(box["x"] + box["width"]/2, box["y"] + box["height"]/2)
        page.wait_for_timeout(80)
        page.mouse.click(box["x"] + box["width"]/2, box["y"] + box["height"]/2)
        return True
    except Exception:
        return False

def _wait_search_suggestion_and_click(page: Page, chat_name: str) -> bool:
    page.wait_for_timeout(600)
    panel = page.locator("[role='listbox'], [data-tid*='search-suggestions']")
    elapsed = 0
    while elapsed < 7000:
        try:
            if panel.count() > 0 and panel.first.is_visible():
                p = panel.first
                header = _get_group_chat_section(p)
                if header:
                    candidates = p.locator("[role='option'], li, div[role='menuitem'], div[role='option']")
                    cnt = candidates.count()
                    for i in range(cnt):
                        it = candidates.nth(i)
                        try:
                            if not it.is_visible():
                                continue
                            txt = it.inner_text(timeout=800)
                        except Exception:
                            continue
                        if _is_ng_suggestion_text(txt):
                            continue
                        try:
                            hb = header.bounding_box()
                            ib = it.bounding_box()
                            if hb and ib and ib["y"] <= hb["y"]:
                                continue
                        except Exception:
                            pass
                        if chat_name in txt:
                            try: it.scroll_into_view_if_needed(timeout=1500)
                            except Exception: pass
                            try: it.hover(timeout=800)
                            except Exception: pass
                            try: it.click(timeout=2000)
                            except Exception:
                                if not _click_center(page, it):
                                    it.click(timeout=2000, force=True)
                            return True
                items = p.locator("[role='option'], li, div[role='menuitem'], div[role='option']")
                cnt = items.count()
                best_idx = None
                for i in range(cnt):
                    it = items.nth(i)
                    try:
                        if not it.is_visible():
                            continue
                        txt = it.inner_text(timeout=800)
                    except Exception:
                        continue
                    if _is_ng_suggestion_text(txt):
                        continue
                    if it.locator("[data-icon-name*='Search'], svg[aria-label*='検索']").count() > 0:
                        continue
                    has_avatar = it.locator("[data-tid*='avatar'], img, [class*='avatar']").count() > 0
                    if chat_name in txt and has_avatar:
                        best_idx = i
                        break
                if best_idx is not None:
                    it = items.nth(best_idx)
                    try: it.scroll_into_view_if_needed(timeout=1500)
                    except Exception: pass
                    try: it.hover(timeout=800)
                    except Exception: pass
                    try: it.click(timeout=2000)
                    except Exception:
                        if not _click_center(page, it):
                            it.click(timeout=2000, force=True)
                    return True
        except Exception:
            pass
        page.wait_for_timeout(200)
        elapsed += 200
    return False

def _open_chat_via_search_suggestion(page: Page, chat_name: str) -> bool:
    search_box = _find_search_box(page)
    if not search_box:
        return False
    try:
        search_box.click()
        try:
            search_box.fill("")
        except Exception:
            pass
        search_box.type(chat_name, delay=10)
        if _wait_search_suggestion_and_click(page, chat_name):
            for _ in range(20):
                try:
                    box = _find_message_box(page)
                    if box and box.is_visible():
                        return True
                except Exception:
                    pass
                page.wait_for_timeout(200)
    except Exception:
        return False
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

        _add_recipient(page, to_box, admin_email)
        page.wait_for_timeout(BETWEEN_RECIPIENTS_PAUSE_MS)

        to_box = None
        for _ in range(10):
            to_box = _find_to_field(page) or _focus_to_field(page, to_box)
            if to_box:
                break
            page.wait_for_timeout(200)
        if not to_box:
            raise RuntimeError("2人目追加のための宛先欄を再取得できませんでした。")

        _add_recipient(page, to_box, target_email)

        _set_chat_name_if_available(page, chat_name)

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
            print("送信完了：テンプレ①を投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")

    except PWTimeout as e:
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()

# ------- 既存チャットへ本文送信（テンプレ③ Step1） -------
def open_existing_chat_and_send_message(*, chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        chat_nav = _get_by_role_any(page, [("link", r"チャット"), ("link", r"Chat"), ("button", r"チャット"), ("button", r"Chat")])
        if chat_nav:
            chat_nav.click()
            page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        opened = _open_chat_via_search_suggestion(page, chat_name)
        if not opened:
            raise RuntimeError("検索サジェストから対象チャットを開けませんでした。")

        msg_box = None
        for _ in range(15):
            try:
                msg_box = _find_message_box(page)
                if msg_box and msg_box.is_visible():
                    break
            except Exception:
                pass
            page.wait_for_timeout(200)
        if not msg_box:
            raise RuntimeError("メッセージ入力欄が見つかりませんでした。")

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

# ================= 添付（堅牢版） =================

def _is_send_enabled(page: Page) -> bool:
    btn = _get_by_role_any(page, [("button", r"送信"), ("button", r"Send")]) or _query_any(page, [
        "[data-tid='send-message-button']", "[data-tid='send-button']",
        "button[aria-label*='送信']", "button[aria-label*='Send']",
    ])
    if not btn:
        return False
    try:
        # disabled/aria-disabled をチェック（なければクリック可能性で判断）
        state = btn.get_attribute("disabled") or btn.get_attribute("aria-disabled")
        if state in ("true", "disabled"):
            return False
        return True
    except Exception:
        return True  # 判定不能なら楽観視

def _wait_upload_ready(page: Page, file_name: str, timeout_ms: int) -> bool:
    """
    条件:
      1) ファイル名チップ（プレビュー/添付カード）が表示される
      2) 'アップロード中/Uploading' やスピナーが見えなくなる
      3) 送信ボタンが有効
    """
    waited = 0
    step = 300
    while waited < timeout_ms:
        ready_chip = False
        uploading = False

        try:
            # 1) ファイル名を含む要素
            chip = page.locator(f"text={file_name}")
            if chip.count() > 0 and chip.first.is_visible():
                ready_chip = True
        except Exception:
            pass

        try:
            # 2) 代表的なアップロード中インジケータ
            #   - 砂時計/スピナー/プログレス、"アップロード中"/"Uploading"文言など
            if page.locator("text=/アップロード中|Uploading/i").count() > 0:
                uploading = True
            elif page.locator("[role='progressbar'], progress").count() > 0:
                uploading = True
            elif page.locator("[data-icon-name*='Progress'], [data-icon-name*='Spinner']").count() > 0:
                uploading = True
        except Exception:
            pass

        send_ok = _is_send_enabled(page)

        debug(f"attach: chip={ready_chip} uploading={uploading} send_ok={send_ok}")

        if ready_chip and not uploading and send_ok:
            return True

        page.wait_for_timeout(step)
        waited += step
    return False

def _attach_one_file(page: Page, file_path: str) -> bool:
    file_name = os.path.basename(file_path)

    # 添付（クリップ）ボタン
    attach_btn = _get_by_role_any(page, [
        ("button", r"添付"), ("button", r"Attach"),
    ]) or _query_any(page, [
        "[data-tid='attach-button']",
        "button[aria-label*='添付']", "button[aria-label*='Attach']",
        "button[data-tid*='attachment']",
    ])
    if not attach_btn:
        raise RuntimeError("添付ボタンが見つかりませんでした。")

    try:
        attach_btn.click(timeout=5_000)
    except Exception:
        attach_btn.click(timeout=5_000, force=True)

    # 「このデバイスからアップロード」
    device_opt = _get_by_role_any(page, [
        ("menuitem", r"このデバイス|デバイスから"),
        ("menuitem", r"Upload from this device"),
    ]) or _query_any(page, [
        "div[role='menuitem']:has-text('このデバイス')",
        "div[role='menuitem']:has-text('Upload from this device')",
    ])
    if not device_opt:
        raise RuntimeError("『このデバイスからアップロード』が見つかりませんでした。")

    with page.expect_file_chooser(timeout=15_000) as fc_info:
        try:
            device_opt.click(timeout=5_000)
        except Exception:
            device_opt.click(timeout=5_000, force=True)
    file_chooser = fc_info.value
    file_chooser.set_files(file_path)

    # 完了待ち
    if _wait_upload_ready(page, file_name=file_name, timeout_ms=ATTACH_UPLOAD_TIMEOUT_MS):
        debug("attach: ready -> OK")
        return True

    debug("attach: not ready within timeout")
    return False

# ------- 既存チャットへ本文＋ファイル1つ添付（テンプレ③ Step2） -------
def open_existing_chat_and_send_message_with_file(*, chat_name: str, message: str, file_path: str) -> None:
    """
    既存チャットを検索サジェストから開き、本文を入力し、PDFを添付してから送信。
    アップロードが完了しない場合は自動リトライ。失敗時の振る舞いは .env で制御。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        chat_nav = _get_by_role_any(page, [("link", r"チャット"), ("link", r"Chat"), ("button", r"チャット"), ("button", r"Chat")])
        if chat_nav:
            chat_nav.click()
            page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        if not _open_chat_via_search_suggestion(page, chat_name):
            raise RuntimeError("検索サジェストから対象チャットを開けませんでした。")

        # 本文入力欄
        msg_box = None
        for _ in range(15):
            try:
                msg_box = _find_message_box(page)
                if msg_box and msg_box.is_visible():
                    break
            except Exception:
                pass
            page.wait_for_timeout(200)
        if not msg_box:
            raise RuntimeError("メッセージ入力欄が見つかりませんでした。")

        # 本文を先に入力（添付後に送信）
        _type_multiline(page, msg_box, message)

        # 添付（リトライあり）
        ok = _attach_one_file(page, file_path=file_path)
        retries = ATTACH_RETRIES
        while not ok and retries > 0:
            debug("attach: retry")
            ok = _attach_one_file(page, file_path=file_path)
            retries -= 1

        if not ok and ATTACH_FAIL_BEHAVIOR == "abort":
            print("添付アップロードが完了しなかったため、送信を中止しました。")
            return
        elif not ok:
            print("警告: 添付が完了しませんでした。本文のみ送信します。")

        # 送信 → チェックマーク待機
        base_icons = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base_icons, timeout_ms=DELIVERY_WAIT_MS):
            if ok:
                print("送信完了：本文＋添付を投稿しました。（配信確認OK）")
            else:
                print("送信完了：本文のみ投稿しました。（配信確認OK）")
        else:
            if ok:
                print("送信は試行しましたが、配信確認が取れませんでした。（本文＋添付）")
            else:
                print("送信は試行しましたが、配信確認が取れませんでした。（本文のみ）")

    except PWTimeout as e:
        try:
            debug(f"timeout: {e}")
            page.wait_for_timeout(2000)  # 状況確認の猶予
        except Exception:
            pass
        raise RuntimeError(f"Playwright timeout: {e}") from e
    finally:
        browser.close()
        pw.stop()
