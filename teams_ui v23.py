# -*- coding: utf-8 -*-
import os
import re
from typing import Optional, Sequence
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page, Locator

# ========= 可変チューニング値（.envで上書き可） =========
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"
_DEFAULT_TIMEOUT = 150_000

SUGGESTION_MIN_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MIN_WAIT_MS", "1400"))
SUGGESTION_MAX_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MAX_WAIT_MS", "6000"))
BETWEEN_RECIPIENTS_PAUSE_MS = int(os.environ.get("TPL_BETWEEN_RECIPIENTS_PAUSE_MS", "250"))
BEFORE_OPEN_CHATNAME_MS = int(os.environ.get("TPL_BEFORE_OPEN_CHATNAME_MS", "200"))
DELIVERY_WAIT_MS = int(os.environ.get("TPL_DELIVERY_WAIT_MS", "20000"))

ATTACH_UPLOAD_TIMEOUT_MS = int(os.environ.get("ATTACH_UPLOAD_TIMEOUT_MS", "60000"))
ATTACH_RETRIES = int(os.environ.get("ATTACH_RETRIES", "2"))
ATTACH_FAIL_BEHAVIOR = os.environ.get("ATTACH_FAIL_BEHAVIOR", "send_without_file")  # "abort" | "send_without_file"

DEBUG = os.environ.get("DEBUG_LOG", "0") == "1"
def debug(m:str):
    if DEBUG: print(f"[DEBUG] {m}")

# =======================================================
def _launch():
    channel = os.environ.get("PLAYWRIGHT_CHANNEL", "msedge")
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
        browser.close(); pw.stop()

# ------- ユーティリティ -------
def _ensure_ready(page: Page):
    page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=_DEFAULT_TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

def _get_by_role_any(page: Page, candidates: Sequence[tuple[str, str]]) -> Optional[Locator]:
    for role, name_pat in candidates:
        try:
            loc = page.get_by_role(role, name=re.compile(name_pat))
            loc.first.wait_for(timeout=2000)
            return loc
        except Exception:
            continue
    return None

def _query_any(page: Page, selectors: Sequence[str]) -> Optional[Locator]:
    for sel in selectors:
        loc = page.locator(sel)
        try:
            loc.first.wait_for(timeout=2000)
            return loc
        except Exception:
            continue
    return None

def _find_to_field(page: Page) -> Optional[Locator]:
    loc = _get_by_role_any(page, [("combobox", r"宛先"), ("combobox", r"To")])
    if loc: return loc
    loc = _get_by_role_any(page, [("textbox", r"宛先|To|検索|Search")])
    if loc: return loc
    return _query_any(page, [
        "input[placeholder*='宛先']", "input[placeholder*='To']",
        "input[aria-label*='宛先']", "input[aria-label*='To']",
        "div[role='combobox'][aria-label*='宛先']", "div[role='combobox'][aria-label*='To']",
        "div[contenteditable='true'][aria-label*='宛先']", "div[contenteditable='true'][aria-label*='To']",
        "input[placeholder*='名前']", "input[placeholder*='Name']",
        "[data-tid='people-picker-input']", "[data-tid='newChat-peoplePicker']",
        "div[role='combobox']",
    ])

def _find_message_box(page: Page) -> Locator:
    loc = _query_any(page, [
        "div[contenteditable='true'][role='textbox']",
        "div[role='textbox'][aria-label*='メッセージ']",
        "div[role='textbox'][aria-label*='message']",
        "div[contenteditable='true']",
    ])
    return loc or page.locator("[contenteditable='true']").first

def _set_chat_name_if_available(page: Page, chat_name: str):
    if not chat_name: return
    name_btn = _get_by_role_any(page, [
        ("button", r"グループ名を追加|チャット名の編集|名前を追加"),
        ("button", r"Add group name|Edit chat name|Add name"),
    ])
    if not name_btn: return
    try:
        page.wait_for_timeout(BEFORE_OPEN_CHATNAME_MS)
        name_btn.click(timeout=10000)
        name_box = _get_by_role_any(page, [("textbox", r"グループ名|チャット名|Name")])
        if name_box:
            name_box.fill(chat_name)
            save_btn = _get_by_role_any(page, [("button", r"保存|適用|完了|Save|Apply|Done")])
            (save_btn or name_box).click()
    except Exception:
        pass

def _recipient_chip_count(page: Page) -> int:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']", "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill", "[aria-label*='削除'] span.pill", "[aria-label*='Remove'] span.pill",
    ])
    if not loc: return 0
    try: return loc.count()
    except Exception: return 1

def _chip_exists(page: Page, address: str) -> bool:
    loc = _query_any(page, [
        "[data-tid='people-picker-selected']", "[data-tid='people-picker-selectedItem']",
        ".people-picker .pill",
    ])
    if not loc: return False
    try:
        for i in range(loc.count()):
            txt = loc.nth(i).inner_text(timeout=800)
            if address.lower() in txt.lower(): return True
    except Exception:
        pass
    return False

def _ensure_invite_if_needed(page: Page):
    btn = _get_by_role_any(page, [("button", r"招待|Invite|追加|Add|参加|Join|送信|Send")])
    if btn:
        try: btn.click(timeout=5000); page.wait_for_timeout(300)
        except Exception: pass

def _focus_to_field(page: Page, to_box: Optional[Locator]) -> Optional[Locator]:
    if to_box:
        try: to_box.click(timeout=2000); return to_box
        except Exception: pass
    container = _query_any(page, ["[data-tid='people-picker']", "[data-tid='newChat-peoplePicker']", "div[role='combobox']"])
    if container:
        try:
            container.click(timeout=2000); page.keyboard.press("Tab"); page.wait_for_timeout(100)
        except Exception: pass
    return _find_to_field(page)

# ---- 候補関連（宛先ピッカー）----
def _latest_listbox(page: Page) -> Optional[Locator]:
    boxes = page.locator("[role='listbox']")
    try:
        n = boxes.count()
        if n == 0:
            boxes = page.locator("[data-tid*='people-picker'] [role='listbox'], [id*='Dropdown'] [role='listbox']")
            n = boxes.count()
        if n == 0: return None
        for i in range(n - 1, -1, -1):
            lb = boxes.nth(i)
            try:
                if lb.is_visible(): return lb
            except Exception: continue
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
                if opts.count() > 0 and opts.first.is_visible(): return True
            except Exception: pass
        page.wait_for_timeout(step); elapsed += step
    return False

def _click_matching_option(page: Page, address: str) -> bool:
    lb = _latest_listbox(page)
    if not lb: return False
    options = lb.locator("[role='option']")
    try:
        count = options.count()
        if count == 0: return False
        target_idx = None
        for i in range(count):
            opt = options.nth(i)
            try:
                if not opt.is_visible(): continue
                txt = opt.inner_text(timeout=800)
            except Exception: continue
            if address.lower() in txt.lower():
                target_idx = i; break
        if target_idx is None:
            for i in range(count):
                try:
                    if options.nth(i).is_visible(): target_idx = i; break
                except Exception: continue
        if target_idx is None: target_idx = 0
        opt = options.nth(target_idx)
        try: opt.scroll_into_view_if_needed(timeout=2000)
        except Exception: pass
        try: opt.hover(timeout=1000)
        except Exception: pass
        try: opt.click(timeout=2000)
        except Exception: opt.click(timeout=2000, force=True)
        return True
    except Exception:
        return False

def _light_clear_typing_area(page: Page):
    try: page.keyboard.type(" "); page.wait_for_timeout(20); page.keyboard.press("Backspace")
    except Exception: pass

def _add_recipient(page: Page, to_box: Locator, address: str):
    if not address: return
    if _chip_exists(page, address): return
    before = _recipient_chip_count(page)
    to_box = _focus_to_field(page, to_box)
    if not to_box: raise RuntimeError("宛先欄へフォーカスできませんでした。")
    _light_clear_typing_area(page)
    to_box.type(address, delay=16)
    _wait_for_suggestions(page, min_wait_ms=SUGGESTION_MIN_WAIT_MS, max_wait_ms=SUGGESTION_MAX_WAIT_MS)
    clicked = _click_matching_option(page, address)
    for _ in range(10):
        page.wait_for_timeout(200); _ensure_invite_if_needed(page)
        after = _recipient_chip_count(page)
        if after > before or _chip_exists(page, address): return
    if not clicked:
        lb = _latest_listbox(page)
        if lb:
            opt = lb.locator("[role='option']").first
            try:
                opt.scroll_into_view_if_needed(timeout=2000); opt.hover(timeout=1000); opt.click(timeout=2000, force=True)
            except Exception: pass

# ---- メッセージ入力・送信 ----
def _type_multiline(page: Page, box: Locator, text: str):
    lines = text.splitlines()
    box.click()
    for i, line in enumerate(lines):
        if line: box.type(line, delay=10)
        if i < len(lines) - 1: box.press("Shift+Enter")

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
        try: c += page.locator(sel).count()
        except Exception: pass
    return c

def _click_send_button(page: Page) -> bool:
    send_btn = _get_by_role_any(page, [("button", r"送信|Send")]) or _query_any(page, [
        "[data-tid='send-message-button']", "[data-tid='send-button']",
        "button[aria-label*='送信']", "button[aria-label*='Send']",
    ])
    if not send_btn: return False
    try: send_btn.click(timeout=3000); return True
    except Exception:
        try: send_btn.click(timeout=3000, force=True); return True
        except Exception: return False

def _wait_delivery_increase(page: Page, base: int, timeout_ms: int) -> bool:
    waited, step = 0, 200
    while waited < timeout_ms:
        try:
            if _delivery_icon_count(page) > base: return True
        except Exception: pass
        page.wait_for_timeout(step); waited += step
    return False

# ====== 検索（Enter禁止・サジェストクリック） ======
def _find_search_box(page: Page) -> Optional[Locator]:
    debug("search: try Ctrl+E")
    for _ in range(3):
        try: page.keyboard.press("Control+e"); page.wait_for_timeout(200)
        except Exception: pass
        loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
            "input[type='search']", "input[placeholder*='検索']", "input[placeholder*='Search']",
            "input[role='combobox'][type='search']", "div[role='combobox'] input[type='search']",
            "input[aria-label*='検索']", "input[aria-label*='Search']",
        ])
        if loc:
            try: loc.first.wait_for(timeout=800); debug("search: got by selector after Ctrl+E"); return loc
            except Exception: pass
    debug("search: click top bar area then re-detect")
    try: page.mouse.move(400, 60); page.mouse.click(400, 60); page.wait_for_timeout(150)
    except Exception: pass
    loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
        "input[type='search']", "input[placeholder*='検索']", "input[placeholder*='Search']",
        "input[role='combobox'][type='search']", "div[role='combobox'] input[type='search']",
        "input[aria-label*='検索']", "input[aria-label*='Search']",
    ])
    if loc: debug("search: got by selector after top click"); return loc
    debug("search: fallback '/' key then re-detect")
    try: page.keyboard.press("/"); page.wait_for_timeout(120)
    except Exception: pass
    loc = _get_by_role_any(page, [("textbox", r"検索|Search")]) or _query_any(page, [
        "input[type='search']", "input[placeholder*='検索']", "input[placeholder*='Search']",
        "input[role='combobox'][type='search']", "div[role='combobox'] input[type='search']",
        "input[aria-label*='検索']", "input[aria-label*='Search']",
    ])
    if loc: debug("search: got by selector after '/'"); return loc
    debug("search: NOT FOUND"); return None

def _is_ng_suggestion_text(txt: str) -> bool:
    t = txt.replace("\n", " ")
    for pat in [r"Enter.?キー.*結果.*表示", r"結果.*表示", r"ユーザー.*招待", r"Invite.*to.*Teams"]:
        if re.search(pat, t): return True
    return False

def _click_center(page: Page, el: Locator) -> bool:
    try:
        box = el.bounding_box()
        if not box: return False
        x = box["x"] + box["width"]/2; y = box["y"] + box["height"]/2
        page.mouse.move(x, y); page.wait_for_timeout(80); page.mouse.click(x, y)
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
                items = p.locator("[role='option'], li, div[role='menuitem'], div[role='option']")
                cnt = items.count()
                best = None
                for i in range(cnt):
                    it = items.nth(i)
                    try:
                        if not it.is_visible(): continue
                        txt = it.inner_text(timeout=800)
                    except Exception: continue
                    if _is_ng_suggestion_text(txt): continue
                    if it.locator("[data-icon-name*='Search'], svg[aria-label*='検索']").count() > 0:
                        continue
                    has_avatar = it.locator("[data-tid*='avatar'], img, [class*='avatar']").count() > 0
                    if chat_name in txt and has_avatar:
                        best = i; break
                if best is None and cnt > 0: best = 0
                if best is not None:
                    it = items.nth(best)
                    try: it.scroll_into_view_if_needed(timeout=1500)
                    except Exception: pass
                    try: it.hover(timeout=800)
                    except Exception: pass
                    try: it.click(timeout=2000)
                    except Exception:
                        if not _click_center(page, it): it.click(timeout=2000, force=True)
                    return True
        except Exception: pass
        page.wait_for_timeout(200); elapsed += 200
    return False

def _open_chat_via_search_suggestion(page: Page, chat_name: str) -> bool:
    sb = _find_search_box(page)
    if not sb: return False
    try:
        sb.click()
        try: sb.fill("")
        except Exception: pass
        sb.type(chat_name, delay=10)
        if _wait_search_suggestion_and_click(page, chat_name):
            for _ in range(20):
                try:
                    box = _find_message_box(page)
                    if box and box.is_visible(): return True
                except Exception: pass
                page.wait_for_timeout(200)
    except Exception:
        return False
    return False

# ------- テンプレ① -------
def create_group_chat_and_send_message(*, admin_email: str, target_email: str, chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        chat_nav = _get_by_role_any(page, [("link", r"チャット|Chat"), ("button", r"チャット|Chat")])
        if chat_nav: chat_nav.click(); page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        new_chat = _get_by_role_any(page, [("button", r"新しいチャット|New chat"), ("link", r"新しいチャット|New chat")]) \
                   or _query_any(page, ["[data-tid='new-chat-button']", "[aria-label*='新しいチャット']", "[aria-label*='New chat']"])
        if not new_chat: raise RuntimeError("新しいチャットの入口が見つかりませんでした。")
        new_chat.click()

        to_box = None
        for _ in range(10):
            to_box = _find_to_field(page)
            if to_box: break
            page.wait_for_timeout(200)
        if not to_box: raise RuntimeError("宛先入力欄が見つかりませんでした。")

        _add_recipient(page, to_box, admin_email)
        page.wait_for_timeout(BETWEEN_RECIPIENTS_PAUSE_MS)
        to_box = _find_to_field(page) or _focus_to_field(page, to_box)
        if not to_box: raise RuntimeError("2人目追加のための宛先欄を再取得できませんでした。")
        _add_recipient(page, to_box, target_email)
        _set_chat_name_if_available(page, chat_name)

        msg_box = _find_message_box(page)
        _type_multiline(page, msg_box, message)

        base = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base, timeout_ms=DELIVERY_WAIT_MS):
            print("送信完了：テンプレ①を投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")
    finally:
        browser.close(); pw.stop()

# ------- テンプレ③ Step1：本文のみ -------
def open_existing_chat_and_send_message(*, chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        chat_nav = _get_by_role_any(page, [("link", r"チャット|Chat"), ("button", r"チャット|Chat")])
        if chat_nav: chat_nav.click(); page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        if not _open_chat_via_search_suggestion(page, chat_name):
            raise RuntimeError("検索サジェストから対象チャットを開けませんでした。")

        msg_box = None
        for _ in range(15):
            try:
                msg_box = _find_message_box(page)
                if msg_box and msg_box.is_visible(): break
            except Exception: pass
            page.wait_for_timeout(200)
        if not msg_box: raise RuntimeError("メッセージ入力欄が見つかりませんでした。")

        _type_multiline(page, msg_box, message)

        base = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base, timeout_ms=DELIVERY_WAIT_MS):
            print("送信完了：本文のみ投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")
    finally:
        browser.close(); pw.stop()

# ================= 添付（強化版） =================
def _is_send_enabled(page: Page) -> bool:
    btn = _get_by_role_any(page, [("button", r"送信|Send")]) or _query_any(page, [
        "[data-tid='send-message-button']", "[data-tid='send-button']",
        "button[aria-label*='送信']", "button[aria-label*='Send']",
    ])
    if not btn: return False
    try:
        state = btn.get_attribute("disabled") or btn.get_attribute("aria-disabled")
        return not (state in ("true", "disabled"))
    except Exception:
        return True

def _composer_root(page: Page) -> Optional[Locator]:
    for sel in [
        "[data-tid='messagePane']", "[data-tid='message-composer']",
        "div[role='textbox'] >> xpath=ancestor::div[contains(@class,'composer')][1]",
        "div[role='textbox'] >> xpath=ancestor::div[1]",
    ]:
        loc = page.locator(sel)
        try:
            if loc.count() > 0 and loc.first.is_visible(): return loc.first
        except Exception: pass
    return None

def _try_direct_file_input(page: Page, file_path: str) -> bool:
    root = _composer_root(page) or page
    fis = root.locator("input[type='file']")
    try: n = fis.count()
    except Exception: n = 0
    if n == 0: return False
    for i in range(min(n, 5)):
        try:
            debug("attach: try direct input[type=file]")
            fis.nth(i).set_input_files(file_path)
            return True
        except Exception: continue
    return False

def _wait_upload_ready(page: Page, file_name: str, timeout_ms: int) -> bool:
    waited = 0; step = 300; short = file_name[:8]
    while waited < timeout_ms:
        ready = False; uploading = False
        try:
            chip = page.locator(f"text={file_name}")
            if chip.count() == 0: chip = page.locator(f"text={short}")
            if chip.count() > 0 and chip.first.is_visible(): ready = True
        except Exception: pass
        try:
            if page.locator("text=/アップロード中|Uploading/i").count() > 0: uploading = True
            elif page.locator("[role='progressbar'], progress").count() > 0: uploading = True
            elif page.locator("[data-icon-name*='Progress'], [data-icon-name*='Spinner']").count() > 0: uploading = True
        except Exception: pass
        send_ok = _is_send_enabled(page)
        debug(f"attach: chip={ready} uploading={uploading} send_ok={send_ok}")
        if ready and not uploading and send_ok: return True
        page.wait_for_timeout(step); waited += step
    return False

def _wait_menu_open(page: Page, timeout_ms: int = 5000) -> Optional[Locator]:
    waited = 0; step = 120
    while waited < timeout_ms:
        panel = page.locator("[role='menu'], [data-tid*='menu']")
        try:
            if panel.count() > 0 and panel.first.is_visible(): return panel.first
        except Exception: pass
        page.wait_for_timeout(step); waited += step
    return None

def _open_attach_menu_and_pick_device(page: Page):
    """ クリップ or ＋ のメニューから『このデバイスからアップロード』を開く。成功時 FileChooser を返す。 """
    # 1) まずクリップ
    attach_btn = (
        _get_by_role_any(page, [("button", r"添付|ファイル|Attach|File")]) or
        _query_any(page, [
            "[data-tid='attach-button']", "[data-tid*='attachment']",
            "button[title*='添付']", "button[title*='Attach']", "button[title*='ファイル']",
            "button[aria-label*='添付']", "button[aria-label*='Attach']", "button[aria-label*='ファイル']",
            "[data-icon-name='Attach']", "[data-icon-name*='Paperclip']", "[data-icon-name*='Clip']",
        ])
    )
    used = None
    if attach_btn:
        try:
            debug("attach: click clip button")
            attach_btn.click(timeout=5000)
            used = "clip"
        except Exception:
            pass

    # 2) クリップで開かなければ ＋（さらに）
    if not used:
        plus_btn = (
            _get_by_role_any(page, [("button", r"さらに|More|アプリ|Apps|\+"),]) or
            _query_any(page, [
                "[data-icon-name*='Add']", "[data-icon-name*='Plus']", "button[aria-label*='さらに']",
                "button[title*='さらに']", "button[aria-label*='More']", "button[title*='More']",
            ])
        )
        if plus_btn:
            try:
                debug("attach: click plus button")
                plus_btn.click(timeout=5000)
                used = "plus"
            except Exception:
                pass

    # 3) メニュー待ち
    menu = _wait_menu_open(page, 5000)
    if not menu:
        debug("attach: menu not opened")
        return None

    # 4) 『このデバイスからアップロード』
    device_item = (
        page.get_by_role("menuitem", name=re.compile(r"(このデバイス|デバイスから|アップロード|Upload from this device)", re.I))
    )
    try:
        if device_item.count() == 0:
            device_item = menu.locator(":is(div,button)[role='menuitem']:has-text('このデバイス'), :is(div,button):has-text('このデバイスからアップロード')")
        device_item.first.wait_for(timeout=3000)
    except Exception:
        debug("attach: device menuitem not found")
        return None

    # 5) FileChooser
    try:
        with page.expect_file_chooser(timeout=15000) as fc_info:
            try: device_item.first.click(timeout=3000)
            except Exception: device_item.first.click(timeout=3000, force=True)
        return fc_info.value
    except Exception:
        debug("attach: expect_file_chooser failed")
        return None

def _handle_replace_modal_if_any(page: Page):
    try:
        modal = page.locator("text=/このファイルは既に存在します/i")
        if modal.count() > 0 and modal.first.is_visible():
            rep = page.get_by_role("button", name=re.compile(r"(置換|Replace)", re.I))
            if rep.count() > 0:
                rep.first.click(timeout=5000)
                page.wait_for_timeout(200)
    except Exception:
        pass

def _attach_one_file(page: Page, file_path: str) -> bool:
    file_name = os.path.basename(file_path)

    if _try_direct_file_input(page, file_path):
        debug("attach: direct input[type=file] -> set_files OK")
        _handle_replace_modal_if_any(page)
        return _wait_upload_ready(page, file_name, ATTACH_UPLOAD_TIMEOUT_MS)

    fc = _open_attach_menu_and_pick_device(page)
    if not fc:
        debug("attach: could not open device upload")
        return False

    try:
        fc.set_files(file_path)
    except Exception:
        debug("attach: file chooser set_files failed")
        return False

    _handle_replace_modal_if_any(page)
    return _wait_upload_ready(page, file_name, ATTACH_UPLOAD_TIMEOUT_MS)

# ------- テンプレ③ Step2：本文＋1ファイル -------
def open_existing_chat_and_send_message_with_file(*, chat_name: str, message: str, file_path: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        chat_nav = _get_by_role_any(page, [("link", r"チャット|Chat"), ("button", r"チャット|Chat")])
        if chat_nav: chat_nav.click(); page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)

        if not _open_chat_via_search_suggestion(page, chat_name):
            raise RuntimeError("検索サジェストから対象チャットを開けませんでした。")

        msg_box = None
        for _ in range(15):
            try:
                msg_box = _find_message_box(page)
                if msg_box and msg_box.is_visible(): break
            except Exception: pass
            page.wait_for_timeout(200)
        if not msg_box: raise RuntimeError("メッセージ入力欄が見つかりませんでした。")

        _type_multiline(page, msg_box, message)

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

        base = _delivery_icon_count(page)
        if not _click_send_button(page):
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(200)
            try: page.keyboard.press("Enter")
            except Exception: pass

        if _wait_delivery_increase(page, base, timeout_ms=DELIVERY_WAIT_MS):
            print("送信完了：{}（配信確認OK）".format("本文＋添付を投稿しました。" if ok else "本文のみ投稿しました。"))
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。{}".format("（本文＋添付）" if ok else "（本文のみ）"))
    finally:
        browser.close(); pw.stop()
