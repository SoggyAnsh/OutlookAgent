import logging
import os
import re
import time
from pathlib import Path
from typing import List, Tuple

from PIL import Image, ImageDraw
from patchright.sync_api import (
    expect,
    Playwright,
    sync_playwright,
    TimeoutError as PWTimeoutError,
    Error as PWError,
)

BASE_URL = os.getenv("BASE_URL", "https://signup.live.com")
OUT = Path("chartifacts")
OUT.mkdir(parents=True, exist_ok=True)
ENABLE_TRACE = False

FIRST_NAME = "Abuel"
LAST_NAME = "Nexus"
EMAIL = f"{FIRST_NAME.lower()}_{LAST_NAME.lower()}9839@outlook.com"
PASSWORD = "AbleNexus@65"
COUNTRY = "Mexico"
DOB_DAY = "19"
DOB_MONTH = "January"
DOB_YEAR = "1992"

FULLSCREEN = True
PYAUTO_FOR_HUMAN_CHECK = True

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler(OUT / "run.log", mode="w", encoding="utf-8")],
)
log = logging.getLogger("signup")

try:
    import pyautogui

    pyautogui.PAUSE = 0.0
    _HAS_PYAUTO = True
except Exception as _e:
    log.warning("pyautogui unavailable: %s", _e)
    _HAS_PYAUTO = False

# --- calibrated CSS->screen mapping ---
_CAL_DPR: float = 1.0
_CAL_XOFF_CSS: float = 0.0
_CAL_YOFF_CSS: float = 0.0


def _update_css_screen_calibration(page) -> None:
    global _CAL_DPR, _CAL_XOFF_CSS, _CAL_YOFF_CSS
    vals = page.evaluate("""() => ({
        dpr: window.devicePixelRatio || 1,
        sx: window.screenX, sy: window.screenY,
        ow: window.outerWidth, oh: window.outerHeight,
        iw: window.innerWidth, ih: window.innerHeight
    })""")
    _CAL_DPR = float(vals["dpr"]) or 1.0
    left_border_css = max(0.0, (vals["ow"] - vals["iw"]) / 2.0)
    top_chrome_css = max(0.0, (vals["oh"] - vals["ih"]) - left_border_css)
    _CAL_XOFF_CSS = float(vals["sx"]) + left_border_css
    _CAL_YOFF_CSS = float(vals["sy"]) + top_chrome_css
    log.info(
        "MAP | dpr=%.3f offs_css=(%.1f,%.1f) outer=%dx%d inner=%dx%d screenXY=(%d,%d)",
        _CAL_DPR, _CAL_XOFF_CSS, _CAL_YOFF_CSS,
        vals["ow"], vals["oh"], vals["iw"], vals["ih"], vals["sx"], vals["sy"]
    )


def _css_to_screen(page, x_css: float, y_css: float) -> Tuple[int, int]:
    sx = int(round((_CAL_XOFF_CSS + x_css) * _CAL_DPR))
    sy = int(round((_CAL_YOFF_CSS + y_css) * _CAL_DPR))
    return sx, sy


def _pyauto_click_css(page, x_css: float, y_css: float, label: str):
    if not _HAS_PYAUTO:
        return False
    sx, sy = _css_to_screen(page, x_css, y_css)
    log.info("%s | pyautogui.click screen=(%d,%d) from css=(%.1f,%.1f)", label, sx, sy, x_css, y_css)
    try:
        pyautogui.moveTo(sx, sy, duration=0.15)
        pyautogui.click()
        return True
    except Exception as e:
        log.warning("%s | pyautogui failed: %s", label, e)
        return False


def _pyauto_move_css(page, x_css: float, y_css: float, label: str) -> bool:
    if not _HAS_PYAUTO:
        return False
    sx, sy = _css_to_screen(page, x_css, y_css)
    log.info("%s | pyautogui.moveTo screen=(%d,%d) from css=(%.1f,%.1f)", label, sx, sy, x_css, y_css)
    try:
        pyautogui.moveTo(sx, sy, duration=0.20)
        return True
    except Exception as e:
        log.warning("%s | pyautogui move failed: %s", label, e)
        return False


def _annotate(path_in: Path, rects: List[Tuple[float, float, float, float]], path_out: Path) -> None:
    img = Image.open(path_in).convert("RGBA")
    draw = ImageDraw.Draw(img, "RGBA")
    for (x, y, w, h) in rects:
        draw.rectangle([x, y, x + w, y + h], outline=(0, 200, 255, 180), width=4)
    img.save(path_out)


def snap_and_annotate(page, rects: List[Tuple[float, float, float, float]], name: str) -> None:
    tmp = OUT / f"{name}_raw.png"
    page.screenshot(path=str(tmp), full_page=True)
    outp = OUT / f"{name}.png"
    _annotate(tmp, rects, outp)
    tmp.unlink(missing_ok=True)


def bbox(locator) -> Tuple[float, float, float, float]:
    box = locator.bounding_box()
    if not box:
        time.sleep(0.2)
        box = locator.bounding_box()
    if not box:
        return (0, 0, 0, 0)
    return (box["x"], box["y"], box["width"], box["height"])


def safe_click(locator, step_name: str, page, also_snap=True):
    locator.wait_for(state="visible", timeout=10_000)
    box = bbox(locator)
    log.info("%s | click @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step_name, *box)
    cx = box[0] + box[2] / 2.0
    cy = box[1] + box[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, cx, cy, f"{step_name}_move")
        _pyauto_click_css(page, cx, cy, f"{step_name}_click")
    else:
        locator.click()
    if also_snap:
        snap_and_annotate(page, [box], step_name)


def safe_type(locator, text: str, step_name: str, page, also_snap=True):
    locator.wait_for(state="visible", timeout=10_000)
    locator.fill("")
    box = bbox(locator)
    log.info("%s | type '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step_name, text, *box)
    locator.type(text, delay=20)
    if also_snap:
        snap_and_annotate(page, [box], step_name)


# --- ARIA combobox utilities ---
COMBO_ID_BY_NAME = {
    "Country/Region": "#countryDropdownId",
    "Month": "#BirthMonthDropdown",
    "Day": "#BirthDayDropdown",
}


def _find_combo(page, names):
    for n in names:
        sel = COMBO_ID_BY_NAME.get(n)
        if sel:
            loc = page.locator(sel)
            try:
                loc.wait_for(state="visible", timeout=3_000)
                return loc
            except PWTimeoutError:
                pass
    for n in names:
        try:
            cb = page.get_by_role("combobox", name=n, exact=True)
            cb.wait_for(state="visible", timeout=3_000)
            return cb
        except PWTimeoutError:
            pass
    for n in names:
        if "/" in n:
            continue
        try:
            cb = page.get_by_role("combobox", name=re.compile(rf"^{re.escape(n)}$", re.I))
            cb.wait_for(state="visible", timeout=3_000)
            return cb
        except PWTimeoutError:
            pass
    return page.get_by_role("combobox").first


def _open_combo_with_fallback(combo, page):
    try:
        box = bbox(combo)
        cx = box[0] + box[2] / 2.0
        cy = box[1] + box[3] / 2.0
        if _HAS_PYAUTO:
            _pyauto_move_css(page, cx, cy, "combo_open_move")
            _pyauto_click_css(page, cx, cy, "combo_open_click")
        else:
            combo.click()
    except (PWTimeoutError, PWError):
        pass

    combo.focus()
    for key in ["Enter", "Alt+ArrowDown"]:
        combo.press(key)
        try:
            page.get_by_role("listbox").first.wait_for(state="visible", timeout=800)
            return
        except PWTimeoutError:
            continue

    box = bbox(combo)
    px = box[0] + max(1, box[2] - 10)
    py = box[1] + box[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, px, py, "combo_open_edge_move")
        _pyauto_click_css(page, px, py, "combo_open_edge_click")
    else:
        combo.click(position={"x": max(1, box[2] - 10), "y": max(1, box[3] / 2)}, force=True)


def ensure_listbox_closed(page):
    try:
        lb = page.get_by_role("listbox").first
        if lb.is_visible():
            page.keyboard.press("Escape")
            try:
                expect(lb).to_be_hidden(timeout=2000)
            except AssertionError:
                pass
    except PWTimeoutError:
        pass


def choose_from_combobox(page, name: str, value: str, step: str):
    names = [name]
    if name.lower() == "month":
        names = ["Month", "Birth month"]
    if name.lower() == "day":
        names = ["Day", "Birth day"]
    combo = _find_combo(page, names)
    cbox = bbox(combo)
    log.info("%s | open '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step, name, *cbox)
    _open_combo_with_fallback(combo, page)
    listbox = page.get_by_role("listbox").first
    option = listbox.get_by_role("option", name=value, exact=True)
    option.scroll_into_view_if_needed()
    obox = bbox(option)
    log.info("%s | choose '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step, value, *obox)
    ocx = obox[0] + obox[2] / 2.0
    ocy = obox[1] + obox[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, ocx, ocy, f"{step}_option_move")
        _pyauto_click_css(page, ocx, ocy, f"{step}_option_click")
    else:
        option.click()
    try:
        expect(combo).to_have_text(re.compile(re.escape(value), re.I), timeout=3_000)
    except AssertionError:
        log.warning("%s | value not reflected proceeding", step)
    snap_and_annotate(page, [cbox, obox], f"{step}_selected")


def find_year_input(page):
    candidates = [
        page.get_by_role("spinbutton", name="Birth year", exact=True),
        page.get_by_role("spinbutton", name=re.compile(r"^Birth year$|^Year$", re.I)),
        page.locator("input[name='BirthYear']"),
        page.locator("input[aria-label='Birth year']"),
        page.get_by_role("textbox", name="Year", exact=True),
    ]
    for loc in candidates:
        try:
            loc.wait_for(state="visible", timeout=1500)
            return loc
        except PWTimeoutError:
            continue
    raise PWTimeoutError("Year input not found/visible")


# --- frames + human check ---
def _all_contexts(page):
    return [page, *page.frames]


def _first_visible(ctx, css: str, timeout_ms: int = 4000):
    loc = ctx.locator(css).filter(has_not=ctx.locator("[aria-disabled='true']")).first
    loc = ctx.locator(css).first
    loc.wait_for(state="visible", timeout=timeout_ms)
    return loc


def _scan_all(page, css_list, timeout_ms=4000):
    deadline = time.monotonic() + (timeout_ms / 1000.0)
    last_err = None
    while time.monotonic() < deadline:
        for ctx in [page, *page.frames]:
            for css in css_list:
                try:
                    return _first_visible(ctx, css, timeout_ms=250)
                except PWTimeoutError as e:
                    last_err = e
        time.sleep(0.05)
    raise PWTimeoutError("element not found in any context") from last_err


def _dump_visible_buttons(page, prefix="debug"):
    try:
        for i, ctx in enumerate(_all_contexts(page)):
            btns = ctx.get_by_role("button")
            n = min(btns.count(), 60)
            for j in range(n):
                b = btns.nth(j)
                try:
                    if b.is_visible():
                        name = b.get_attribute("aria-label") or (b.inner_text() or "").strip()
                        log.info("%s | frame[%d] button[%d] name=%r", prefix, i, j, name)
                except Exception:
                    pass
    except Exception:
        pass


def _press_hold_button(page):
    for ctx in _all_contexts(page):
        try:
            btn = ctx.get_by_role("button", name=re.compile(r"^Press and hold$", re.I)).first
            if btn.count():
                btn.wait_for(state="visible", timeout=3000)
                return btn
        except PWTimeoutError:
            continue
    return None


def solve_human_check(page):
    acc_css = [
        "[role='button'][aria-label='Accessible challenge']:visible",
        "[role='button'][aria-label*='Accessible' i]:visible",
        "a[role='button'][aria-label*='Accessible' i]:visible",
        "[role='button'][aria-label*='Accessibility' i]:visible",
        "[role='button'][aria-label='Accessible challenge']:not([aria-disabled='true'])",
        "[role='button'][aria-label*='Accessible' i]:not([aria-disabled='true'])",
        "a[role='button'][aria-label*='Accessible' i]:not([aria-disabled='true'])",
    ]
    try:
        acc_btn = _scan_all(page, acc_css, timeout_ms=14000)
        acc_btn.scroll_into_view_if_needed()

        abox = bbox(acc_btn)
        ax = abox[0] + abox[2] / 2.0
        ay = abox[1] + abox[3] / 2.0
        if _HAS_PYAUTO:
            _pyauto_move_css(page, ax, ay, "step_13_accessible_challenge_move")
            time.sleep(0.5)
            _pyauto_click_css(page, ax, ay, "step_13_accessible_challenge_click")
        else:
            safe_click(acc_btn, "step_13_accessible_challenge_click_fallback_patchright", page)

        snap_and_annotate(page, [abox], "step_13_accessible_challenge_target")

    except PWTimeoutError:
        _dump_visible_buttons(page, "step_13_buttons_list")
        ph = _press_hold_button(page)
        if not ph:
            raise
        box = bbox(ph)
        x_css = max(1, box[0] - 60)
        y_css = box[1] + box[3] / 2
        if _HAS_PYAUTO:
            _pyauto_move_css(page, x_css, y_css, "step_13_accessible_challenge_os_move")
            time.sleep(0.5)
            _pyauto_click_css(page, x_css, y_css, "step_13_accessible_challenge_os_click")
        else:
            _pyauto_click_css(page, x_css, y_css, "step_13_accessible_challenge_os_click_no_pyauto_fallback")
        snap_and_annotate(page, [(x_css - 10, y_css - 10, 20, 20)], "step_13_accessible_challenge_fallback")

    page.wait_for_timeout(9_000)
    press_again_css = [
        "[role='button'][aria-label='Press again']",
        "[role='button'][aria-label*='Press again' i]",
    ]
    try:
        pa_btn = _scan_all(page, press_again_css, timeout_ms=7000)
    except PWTimeoutError:
        pa_btn = None
        for ctx in _all_contexts(page):
            cand = ctx.get_by_role("button", name=re.compile(r"^Press again$", re.I)).first
            if cand.count():
                cand.wait_for(state="visible", timeout=4000)
                pa_btn = cand
                break
        if pa_btn is None:
            _dump_visible_buttons(page, "step_14_buttons_list")
            raise

    pbox = bbox(pa_btn)
    px = pbox[0] + pbox[2] / 2.0
    py = pbox[1] + pbox[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, px, py, "step_14_press_again_move")
        time.sleep(0.5)
        _pyauto_click_css(page, px, py, "step_14_press_again_click")
    else:
        safe_click(pa_btn, "step_14_press_again_click_fallback_patchright", page)

    snap_and_annotate(page, [pbox], "step_14_press_again_target")


def _find_button_any(page, name_regex, timeout_ms=4000):
    """Return a *fresh* visible button locator across dynamic frames, or None."""
    deadline = time.monotonic() + (timeout_ms / 1000.0)
    while time.monotonic() < deadline:
        for ctx in [page, *page.frames]:
            try:
                btn = ctx.get_by_role("button", name=name_regex).first
                btn.wait_for(state="visible", timeout=250)
                return btn
            except (PWTimeoutError, PWError):
                continue
        time.sleep(0.05)
    return None


def handle_quick_note(page) -> bool:
    """Click 'OK' on 'A quick note about your account' if present (buttons only)."""
    hdr_rx = re.compile(r"quick note", re.I)
    ok_rx = re.compile(r"^OK$", re.I)

    saw_hdr = False
    for ctx in [page, *page.frames]:
        try:
            hdr = ctx.get_by_role("heading", name=hdr_rx).first
            hdr.wait_for(state="visible", timeout=250)
            saw_hdr = True
            break
        except (PWTimeoutError, PWError):
            continue

    ok = _find_button_any(page, ok_rx, timeout_ms=1500)
    if not ok:
        if saw_hdr:
            log.info("Quick note header present but no OK found (might have auto-dismissed).")
        else:
            log.info("No Quick note page detected.")
        return False

    box = bbox(ok)
    cx = box[0] + box[2] / 2.0
    cy = box[1] + box[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, cx, cy, "step_15_quick_note_ok_move")
        time.sleep(0.35)
        _pyauto_click_css(page, cx, cy, "step_15_quick_note_ok_click")
        time.sleep(0.35)
        _pyauto_click_css(page, cx, cy, "step_15_quick_note_ok_click")
    else:
        safe_click(ok, "step_15_quick_note_ok_click_fallback_patchright", page)
    snap_and_annotate(page, [box], "step_15_quick_note_ok_target")
    return True


def handle_skip_for_now(page) -> bool:
    """Dismiss passkey prompt by clicking 'Skip for now' (buttons only)."""
    rx = re.compile(r"^Skip for now$", re.I)
    loc = _find_button_any(page, rx, timeout_ms=2500)
    if not loc:
        log.info("No 'Skip for now' found.")
        return False

    box = bbox(loc)
    cx = box[0] + box[2] / 2.0
    cy = box[1] + box[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, cx, cy, "step_16_passkey_skip_move")
        time.sleep(0.35)
        _pyauto_click_css(page, cx, cy, "step_16_passkey_skip_click")
    else:
        safe_click(loc, "step_16_passkey_skip_click_fallback_patchright", page)
    snap_and_annotate(page, [box], "step_16_passkey_skip_target")
    return True


def handle_stay_signed_in(page) -> bool:
    """Dismiss 'Stay signed in?' by clicking 'No' (buttons only)."""
    no_rx = re.compile(r"^No(,?\s*thanks)?$", re.I)
    loc = _find_button_any(page, no_rx, timeout_ms=3000)
    if not loc:
        log.info("No 'Stay signed in' screen detected.")
        return False

    box = bbox(loc)
    cx = box[0] + box[2] / 2.0
    cy = box[1] + box[3] / 2.0
    if _HAS_PYAUTO:
        _pyauto_move_css(page, cx, cy, "step_17_stay_signed_in_no_move")
        time.sleep(0.35)
        _pyauto_click_css(page, cx, cy, "step_17_stay_signed_in_no_click")
    else:
        safe_click(loc, "step_17_stay_signed_in_no_click_fallback_patchright", page)
    snap_and_annotate(page, [box], "step_17_stay_signed_in_no_target")
    return True


def run(playwright: Playwright):
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(OUT / "udata"),
        channel="chrome",
        headless=False,
        no_viewport=True,
    )
    if ENABLE_TRACE:
        context.tracing.start(screenshots=True, snapshots=True, sources=True)

    page = context.new_page()
    page.set_default_timeout(15_000)

    try:
        log.info("OPEN | %s", BASE_URL)
        page.goto(BASE_URL)
        page.bring_to_front()

        _update_css_screen_calibration(page)

        snap_and_annotate(page, [], "step_00_open")
        email_box = page.get_by_role("textbox", name="Email")
        safe_type(email_box, EMAIL, "step_01_email_input", page)
        next_btn = page.get_by_role("button", name="Next")
        next_btn.wait_for(state="visible", timeout=10_000)
        safe_click(next_btn, "step_02_next_after_email", page)

        pwd_box = page.get_by_role("textbox", name="Password")
        safe_type(pwd_box, PASSWORD, "step_03_password_input", page)
        safe_click(page.get_by_role("button", name="Next"), "step_04_next_after_password", page)

        choose_from_combobox(page, "Country/Region", COUNTRY, "step_05_country")
        choose_from_combobox(page, "Month", DOB_MONTH, "step_06_month")
        choose_from_combobox(page, "Day", DOB_DAY, "step_07_day")
        ensure_listbox_closed(page)

        year_box = find_year_input(page)
        safe_type(year_box, DOB_YEAR, "step_08_year_input", page)
        safe_click(page.get_by_role("button", name="Next"), "step_09_next_after_dob", page)

        safe_type(page.get_by_label("First name"), FIRST_NAME, "step_10_first_name", page)
        safe_type(page.get_by_label("Last name"), LAST_NAME, "step_11_last_name", page)
        names_next = page.get_by_role("button", name="Next")
        names_next.wait_for(state="visible", timeout=10_000)

        nb_names = bbox(names_next)
        nx_css = nb_names[0] + nb_names[2] / 2.0
        ny_css = nb_names[1] + nb_names[3] / 2.0

        _pyauto_move_css(page, nx_css, ny_css, "step_12a_mouse_to_next_names")
        safe_click(names_next, "step_12_next_after_names", page)
        time.sleep(10)

        try:
            solve_human_check(page)
        except PWTimeoutError:
            log.warning("Accessible challenge not found skipping as site may differ.")
            log.info(PWTimeoutError.message)
        time.sleep(20)

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except PWTimeoutError:
            pass
        time.sleep(0.3)

        try:
            handled_qn = handle_quick_note(page)
            if handled_qn:
                try:
                    page.wait_for_load_state("networkidle", timeout=3000)
                except PWTimeoutError:
                    pass
                time.sleep(0.3)

            handled_passkey = handle_skip_for_now(page)
            if handled_passkey:
                try:
                    page.wait_for_load_state("networkidle", timeout=3000)
                except PWTimeoutError:
                    pass
                time.sleep(0.3)

            handled_stay = handle_stay_signed_in(page)
            if handled_stay:
                try:
                    page.wait_for_load_state("networkidle", timeout=3000)
                except PWTimeoutError:
                    pass
                time.sleep(0.3)
        except Exception as e:
            log.warning("Optional-screen pipeline error: %s", e)

        log.info("ACCOUNT | email=%s | password=%s", EMAIL, PASSWORD)

        time.sleep(10)
        snap_and_annotate(page, [], "step_16_done")
        log.info("FLOW COMPLETE")

    finally:
        context.close()


if __name__ == "__main__":
    with sync_playwright() as p:
        run(p)
