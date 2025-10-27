# -*- coding: utf-8 -*-
import os, re
from typing import Optional, Sequence, List, Tuple
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page, Locator

# ========= 可変チューニング値（.envで上書き可） =========
PERSISTENT_DIR = os.environ.get("PW_PROFILE_DIR", ".\\pw-profile")
TEAMS_URL = "https://teams.microsoft.com/"
_DEFAULT_TIMEOUT = 150_000

SUGGESTION_MIN_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MIN_WAIT_MS", "1400"))
SUGGESTION_MAX_WAIT_MS = int(os.environ.get("TPL_SUGGESTION_MAX_WAIT_MS", "6000"))
BETWEEN_RECIPIENTS_PAUSE_MS = int(os.environ.get("TPL_BETWEEN_RECIPIENTS_PAUSE_MS", "250"))
BEFORE_OPEN_CHATNAME_MS = int(os.environ.get("TPL_BEFORE_OPEN_CHATNAME_MS", "200"))

DELIVERY_WAIT_MS = int(os.environ.get("DELIVERY_WAIT_MS", os.environ.get("TPL_DELIVERY_WAIT_MS", "30000")))
SEND_RETRIES = int(os.environ.get("SEND_RETRIES", "2"))

ATTACH_UPLOAD_TIMEOUT_MS = int(os.environ.get("ATTACH_UPLOAD_TIMEOUT_MS", "60000"))
ATTACH_RETRIES = int(os.environ.get("ATTACH_RETRIES", "2"))
ATTACH_FAIL_BEHAVIOR = os.environ.get("ATTACH_FAIL_BEHAVIOR", "send_without_file")  # "abort" | "send_without_file"

ATTACH_TOOLTIP_HINT = os.environ.get("ATTACH_TOOLTIP_HINT", "").strip()
ATTACH_FALLBACK_HINTS = ["ファイルを添付", "添付", "ファイル", "Attach", "Add file", "Attach a file"]

DEVICE_ITEM_HINTS = [s.strip() for s in os.environ.get(
    "DEVICE_ITEM_HINTS",
    "このデバイスからアップロード,このデバイスから,デバイスから,コンピューターから,Upload from this device,From this device,From computer"
).split(",") if s.strip()]

CLOUD_NG_HINTS = [s.strip() for s in os.environ.get(
    "CLOUD_NG_HINTS",
    "クラウド,OneDrive,SharePoint,Cloud,チーム サイト,Teams サイト"
).split(",") if s.strip()]

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
        print("Teams を開きました。必要ならサインインしてください。このウィンドウが閉じれば保存されます。")
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
    try: send_btn.click(timeout=2000); return True
    except Exception:
        try: send_btn.click(timeout=2000, force=True); return True
        except Exception: return False

def _msgbox_is_empty(page: Page) -> bool:
    try:
        mb = _find_message_box(page)
        if not mb: return False
        txt = (mb.inner_text(timeout=800) or "").strip()
        html = (mb.inner_html(timeout=800) or "").strip()
        return (txt == "") or (re.sub(r"<[^>]+>", "", html).strip() == "")
    except Exception:
        return False

def _latest_message_region(page: Page) -> Optional[Locator]:
    for sel in ["[data-tid='messageBodyContent']","[data-tid*='message']","div[class*='message']"]:
        loc = page.locator(sel)
        try:
            if loc.count() > 0: return loc.last
        except Exception: pass
    return None

def _wait_message_posted(page: Page, *, text_hint: Optional[str], file_hint: Optional[str], timeout_ms: int) -> bool:
    waited, step = 0, 300
    while waited < timeout_ms:
        if _msgbox_is_empty(page): return True
        reg = _latest_message_region(page)
        if reg:
            try:
                inner = (reg.inner_text(timeout=800) or "")
                if text_hint and (text_hint[:8] in inner or text_hint.strip() in inner): return True
                if file_hint and (file_hint[:6] in inner): return True
            except Exception: pass
        page.wait_for_timeout(step); waited += step
    return False

def _wait_delivery_increase(page: Page, base: int, timeout_ms: int) -> bool:
    waited, step = 0, 200
    while waited < timeout_ms:
        try:
            if _delivery_icon_count(page) > base: return True
        except Exception: pass
        page.wait_for_timeout(step); waited += step
    return False

def _send_and_confirm(page: Page, *, text_for_hint: Optional[str], file_for_hint: Optional[str]) -> bool:
    base = _delivery_icon_count(page)
    for attempt in range(SEND_RETRIES + 1):
        ok = _click_send_button(page)
        if not ok:
            try: page.keyboard.press("Control+Enter")
            except Exception: pass
            page.wait_for_timeout(150)
            try: page.keyboard.press("Enter")
            except Exception: pass
        if _wait_delivery_increase(page, base, timeout_ms=DELIVERY_WAIT_MS): return True
        if _wait_message_posted(page, text_hint=text_for_hint, file_hint=file_for_hint, timeout_ms=DELIVERY_WAIT_MS): return True
        page.wait_for_timeout(300)
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

# ------- テンプレ①：新規GC作成 -------
def _go_to_chat_tab(page: Page):
    chat_nav = (
        _get_by_role_any(page, [("link", r"チャット|Chat"), ("button", r"チャット|Chat")]) or
        _query_any(page, [
            "a[aria-label*='チャット']", "button[aria-label*='チャット']",
            "a[aria-label*='Chat']",   "button[aria-label*='Chat']",
            "[data-tid='app-bar-2']",
        ])
    )
    if chat_nav:
        try: chat_nav.click(timeout=8000); page.wait_for_load_state("networkidle", timeout=_DEFAULT_TIMEOUT)
        except Exception: pass

def create_group_chat_and_send_message(*, admin_email: str, target_email: str, chat_name: str, message: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page); _go_to_chat_tab(page)

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

        sent = _send_and_confirm(page, text_for_hint=message, file_for_hint=None)
        if sent:
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
        sent = _send_and_confirm(page, text_for_hint=message, file_for_hint=None)
        if sent:
            print("送信完了：本文のみ投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。")
    finally:
        browser.close(); pw.stop()

# ================= 添付 =================
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

def _activate_composer(page: Page) -> Optional[Locator]:
    msg = _find_message_box(page)
    try: msg.scroll_into_view_if_needed(timeout=2000)
    except Exception: pass
    try:
        msg.click(timeout=3000)
        page.wait_for_timeout(80)
        page.keyboard.press("Shift+Tab"); page.wait_for_timeout(50)
        page.keyboard.press("Tab"); page.wait_for_timeout(120)
    except Exception:
        return None
    return _composer_root(page)

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

def _try_direct_file_input(page: Page, file_path: str) -> bool:
    root = _composer_root(page) or page
    fis = root.locator("input[type='file']")
    try: n = fis.count()
    except Exception: n = 0
    if n == 0: return False
    debug(f"attach: found file input(s)= {n}")
    for i in range(min(n, 5)):
        try:
            if file_path != "__DONT_SET__":
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

def _collect_toolbar_buttons(page: Page, root: Optional[Locator]) -> List[Tuple[str, Locator]]:
    labels: List[Tuple[str, Locator]] = []
    search_areas = [root] if root else []
    search_areas.append(page)

    hints = [ATTACH_TOOLTIP_HINT] + ATTACH_FALLBACK_HINTS if ATTACH_TOOLTIP_HINT else ATTACH_FALLBACK_HINTS

    for area in search_areas:
        if area is None: continue
        for h in hints:
            if not h: continue
            loc = area.locator(f"button[aria-label*='{h}'], button[title*='{h}']")
            if loc.count() > 0:
                for i in range(min(loc.count(), 6)):
                    labels.append((h, loc.nth(i)))
        loc2 = area.locator("button [data-icon-name*='Attach'], button [data-icon-name*='Paperclip']").locator("xpath=ancestor::button[1]")
        try:
            c2 = loc2.count()
            for i in range(min(c2, 6)):
                labels.append(("icon:Attach", loc2.nth(i)))
        except Exception: pass
        loc3 = area.locator("[data-tid*='attach']")
        try:
            c3 = loc3.count()
            for i in range(min(c3, 6)):
                labels.append(("data-tid:attach", loc3.nth(i)))
        except Exception: pass
        loc4 = area.locator("button[aria-label*='さらに'], button[title*='さらに'], [data-icon-name*='Add'], [data-icon-name*='Plus']").locator("xpath=ancestor::button[1]")
        try:
            c4 = loc4.count()
            for i in range(min(c4, 4)):
                labels.append(("plus", loc4.nth(i)))
        except Exception: pass

    # 近接重複の排除
    uniq: List[Tuple[str, Locator]] = []
    seen = set()
    for lbl, el in labels:
        try:
            box = el.bounding_box(); key = (round(box["x"] if box else 0, 0), round(box["y"] if box else 0, 0))
        except Exception:
            key = (0,0)
        if key in seen: continue
        seen.add(key); uniq.append((lbl, el))
    debug(f"attach: toolbar candidates= {len(uniq)}")
    return uniq

def _click_button_robust(page: Page, el: Locator, tag: str) -> bool:
    try: el.scroll_into_view_if_needed(timeout=2000)
    except Exception: pass
    try: el.hover(timeout=800)
    except Exception: pass
    try: el.focus()
    except Exception: pass
    try:
        el.click(timeout=1500); debug(f"attach: click {tag} (normal)"); return True
    except Exception:
        try:
            el.click(timeout=1500, force=True); debug(f"attach: click {tag} (force)"); return True
        except Exception:
            try:
                box = el.bounding_box()
                if box:
                    x = box["x"] + box["width"]/2; y = box["y"] + box["height"]/2
                    page.mouse.move(x, y); page.wait_for_timeout(60); page.mouse.click(x, y)
                    debug(f"attach: click {tag} (center)"); return True
            except Exception: pass
    return False

def _wait_menu_open(page: Page, timeout_ms: int = 7000) -> Optional[Locator]:
    waited = 0; step = 140
    page.wait_for_timeout(300)
    while waited < timeout_ms:
        panel = page.locator("[role='menu'], [data-tid*='menu']")
        try:
            if panel.count() > 0 and panel.first.is_visible():
                debug("attach: menu opened"); return panel.first
        except Exception: pass
        # role=menu が無くても「このデバイス…」が見えていれば良し
        for hint in DEVICE_ITEM_HINTS:
            dev = page.get_by_text(re.compile(re.escape(hint), re.I))
            try:
                if dev.count() > 0 and dev.first.is_visible():
                    debug("attach: device item visible without menu role")
                    return panel.first if (panel and panel.count()>0) else None
            except Exception: pass
        # input[type=file] だけ出たケース
        root = _composer_root(page) or page
        try:
            if root.locator("input[type='file']").count() > 0:
                debug("attach: file input appeared (no menu)"); return None
        except Exception: pass
        page.wait_for_timeout(step); waited += step
    debug("attach: menu not opened"); return None

def _menu_pick_device_item(page: Page, menu: Optional[Locator]) -> Optional[Locator]:
    candidates: List[Locator] = []
    for hint in DEVICE_ITEM_HINTS:
        loc = page.get_by_text(re.compile(re.escape(hint), re.I))
        try:
            c = loc.count()
            for i in range(min(c, 8)):
                el = loc.nth(i)
                if el.is_visible(): candidates.append(el)
        except Exception: pass

    if menu is not None:
        items = menu.locator("[role='menuitem'], :is(div,button)[role='menuitem'], li, button, div[aria-role='menuitem']")
        try:
            c = items.count()
            for i in range(min(c, 12)):
                el = items.nth(i)
                try:
                    if el.is_visible():
                        txt = (el.inner_text(timeout=600) or "").strip()
                        if any(re.search(re.escape(h), txt, re.I) for h in DEVICE_ITEM_HINTS):
                            candidates.append(el)
                except Exception: pass
        except Exception: pass

    # NG(クラウド) 除外 & 重複排除
    filtered: List[Locator] = []
    seen = set()
    for el in candidates:
        try:
            txt = (el.inner_text(timeout=600) or "").strip()
            if any(re.search(re.escape(ng), txt, re.I) for ng in CLOUD_NG_HINTS):
                continue
            box = el.bounding_box()
            key = (round(box["x"] if box else 0, 1), round(box["y"] if box else 0, 1))
            if key in seen: continue
            seen.add(key); filtered.append(el)
        except Exception: continue

    labels = []
    for el in filtered:
        try: labels.append((el.inner_text(timeout=400) or "").strip())
        except Exception: labels.append("?")
    if labels: debug(f"attach: menu items= {labels}")

    if not filtered: return None
    return filtered[0]

def _click_inside_box(page: Page, el: Locator, *, y_offset: float = -6.0) -> bool:
    try:
        box = el.bounding_box()
        if not box: return False
        x = box["x"] + box["width"]/2
        y = box["y"] + box["height"]/2 + y_offset
        page.mouse.move(x, y); page.wait_for_timeout(60); page.mouse.click(x, y)
        return True
    except Exception:
        return False

def _open_attach_menu_and_pick_device(page: Page):
    """
    クリップ/＋ を叩き、メニューから「このデバイス系」を厳選。
    ★変更点：一度デバイス項目を検出したら【他のツールバー候補を叩かない】で確定する。
    """
    root = _activate_composer(page) or _composer_root(page) or page

    # 直で input[type=file] が作られていないか確認（UI差異対策）
    _try_direct_file_input(page, "__DONT_SET__")

    candidates = _collect_toolbar_buttons(page, root)
    found_device_once = False

    for idx, (lbl, el) in enumerate(candidates, start=1):
        # すでにデバイス項目を視認済みなら、他候補はもうクリックしない
        if found_device_once:
            break

        debug(f"attach: try {lbl}#{idx}")
        if not _click_button_robust(page, el, f"{lbl}#{idx}"):
            continue

        menu = _wait_menu_open(page, 6000)
        # どちらにせよ “このデバイス” を画面全体から拾う
        item = _menu_pick_device_item(page, menu)
        if not item:
            continue

        # 以後は他ボタンを触らない
        found_device_once = True

        # 先に file chooser を待ち受け → 1クリックで確実に拾う
        try:
            with page.expect_file_chooser(timeout=15000) as fc_info:
                # 枠内クリック → 通常 → force の順に最大1回ずつ
                if not _click_inside_box(page, item):
                    if not _click_button_robust(page, item, "device-item"):
                        pass
            return fc_info.value
        except Exception:
            # 失敗したらメニュー残存に備えて軽く待って次へ
            page.wait_for_timeout(200)
            break  # ここで打ち切り（多重クリック防止）

    # メニュー経由が無理でも、input[type=file] が生成されていればそれを使う
    root2 = _composer_root(page) or page
    try:
        fin = root2.locator("input[type='file']")
        if fin.count() > 0:
            debug(f"attach: found file input(s)= {fin.count()}")
            return None
    except Exception: pass

    return None

def _handle_replace_modal_if_any(page: Page):
    try:
        modal = page.locator("text=/このファイルは既に存在します/i")
        if modal.count() > 0 and modal.first.is_visible():
            rep = page.get_by_role("button", name=re.compile(r"(置換|Replace)", re.I))
            if rep.count() > 0:
                rep.first.click(timeout=5000); page.wait_for_timeout(200)
    except Exception: pass

def _attach_one_file(page: Page, file_path: str) -> bool:
    file_name = os.path.basename(file_path)

    if _try_direct_file_input(page, file_path):
        debug("attach: direct input[type=file] -> set_files OK")
        _handle_replace_modal_if_any(page)
        return _wait_upload_ready(page, file_name, ATTACH_UPLOAD_TIMEOUT_MS)

    fc = _open_attach_menu_and_pick_device(page)
    if fc:
        try:
            fc.set_files(file_path)
        except Exception:
            debug("attach: file chooser set_files failed"); return False
        _handle_replace_modal_if_any(page)
        return _wait_upload_ready(page, file_name, ATTACH_UPLOAD_TIMEOUT_MS)

    if _try_direct_file_input(page, file_path):
        debug("attach: fallback direct set_files OK")
        _handle_replace_modal_if_any(page)
        return _wait_upload_ready(page, file_name, ATTACH_UPLOAD_TIMEOUT_MS)

    debug("attach: could not open device upload")
    return False

# ------- テンプレ③ Step2：本文＋1ファイル -------
def open_existing_chat_and_send_message_with_file(*, chat_name: str, message: str, file_path: str) -> None:
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
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
            _activate_composer(page)
            ok = _attach_one_file(page, file_path=file_path)
            retries -= 1

        if not ok and ATTACH_FAIL_BEHAVIOR == "abort":
            print("添付アップロードが完了しなかったため、送信を中止しました。"); return
        elif not ok:
            print("警告: 添付が完了しませんでした。本文のみ送信します。")

        sent = _send_and_confirm(page, text_for_hint=message, file_for_hint=os.path.basename(file_path) if ok else None)
        if sent:
            print("送信完了：{}（配信確認OK）".format("本文＋添付を投稿しました。" if ok else "本文のみ投稿しました。"))
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。{}".format("（本文＋添付）" if ok else "（本文のみ）"))
    finally:
        browser.close(); pw.stop()

def open_existing_chat_and_send_message_with_files(*, chat_name: str, message: str, file_paths: list[str]) -> None:
    """
    既存チャットを検索→本文を入力→file_paths の順で複数添付→一度だけ送信。
    既存の ATTACH_RETRIES / ATTACH_FAIL_BEHAVIOR / DELIVERY_WAIT_MS をそのまま利用。
    """
    pw, browser, page = _launch()
    try:
        _ensure_ready(page)
        if not _open_chat_via_search_suggestion(page, chat_name):
            raise RuntimeError("検索サジェストから対象チャットを開けませんでした。")

        # メッセージ欄を待機
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

        # 本文入力
        _type_multiline(page, msg_box, message)

        # 添付（順番に）
        attached_any = False
        for fp in file_paths:
            ok = _attach_one_file(page, file_path=fp)
            retries = ATTACH_RETRIES
            while not ok and retries > 0:
                debug("attach: retry (multi)")
                _activate_composer(page)
                ok = _attach_one_file(page, file_path=fp)
                retries -= 1
            attached_any = attached_any or ok
            if (not ok) and ATTACH_FAIL_BEHAVIOR == "abort":
                print(f"添付アップロードが完了しなかったため、送信を中止しました。（{os.path.basename(fp)}）")
                return
            if not ok:
                print(f"警告: 添付に失敗しました（{os.path.basename(fp)}）。続行します。")

        # 送信（本文のみ or 本文＋添付）
        file_hint = os.path.basename(file_paths[0]) if attached_any else None
        sent = _send_and_confirm(page, text_for_hint=message, file_for_hint=file_hint)
        if sent:
            if attached_any:
                print("送信完了：本文＋添付を投稿しました。（配信確認OK）")
            else:
                print("送信完了：本文のみ投稿しました。（配信確認OK）")
        else:
            print("送信は試行しましたが、配信確認が取れませんでした。{}".format("（本文＋添付）" if attached_any else "（本文のみ）"))
    finally:
        browser.close()
        pw.stop()
