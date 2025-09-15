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

FIRST_NAME = "Jonathan"
LAST_NAME = "Crackerson"
EMAIL = "jonathan_crawson33@outlook.com"
PASSWORD = "JohnTrumbo@65"
COUNTRY = "Mexico"
DOB_DAY = "17"
DOB_MONTH = "September"
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

# --- calibrated CSS->screen mapping (added) ---
_CAL_DPR: float = 1.0
_CAL_XOFF_CSS: float = 0.0  # viewport-left in screen CSS px
_CAL_YOFF_CSS: float = 0.0  # viewport-top  in screen CSS px


def _update_css_screen_calibration(page) -> None:
    """Compute CSS->screen mapping once per window."""
    global _CAL_DPR, _CAL_XOFF_CSS, _CAL_YOFF_CSS
    vals = page.evaluate("""() => ({
        dpr: window.devicePixelRatio || 1,
        sx: window.screenX, sy: window.screenY,
        ow: window.outerWidth, oh: window.outerHeight,
        iw: window.innerWidth, ih: window.innerHeight
    })""")
    _CAL_DPR = float(vals["dpr"]) or 1.0

    # Approximate left/right borders as equal compute top chrome as the rest.
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
    # Convert a point in the viewport (CSS px) to physical screen px using the calibration above.
    sx = int(round((_CAL_XOFF_CSS + x_css) * _CAL_DPR))
    sy = int(round((_CAL_YOFF_CSS + y_css) * _CAL_DPR))
    return sx, sy


def _pyauto_click_css(page, x_css: float, y_css: float, label: str):
    if not _HAS_PYAUTO: return False
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
    if not box: return (0, 0, 0, 0)
    return (box["x"], box["y"], box["width"], box["height"])


def safe_click(locator, step_name: str, page, also_snap=True):
    locator.wait_for(state="visible", timeout=10_000)
    box = bbox(locator)
    log.info("%s | click @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step_name, *box)
    locator.click()
    if also_snap: snap_and_annotate(page, [box], step_name)


def safe_type(locator, text: str, step_name: str, page, also_snap=True):
    locator.wait_for(state="visible", timeout=10_000)
    locator.fill("")
    box = bbox(locator)
    log.info("%s | type '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step_name, text, *box)
    locator.type(text, delay=20)
    if also_snap: snap_and_annotate(page, [box], step_name)


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
                loc.wait_for(state="visible", timeout=3_000);
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
        if "/" in n: continue
        try:
            cb = page.get_by_role("combobox", name=re.compile(rf"^{re.escape(n)}$", re.I))
            cb.wait_for(state="visible", timeout=3_000)
            return cb
        except PWTimeoutError:
            pass
    return page.get_by_role("combobox").first


def _open_combo_with_fallback(combo, page):
    try:
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
    if name.lower() == "month": names = ["Month", "Birth month"]
    if name.lower() == "day": names = ["Day", "Birth day"]
    combo = _find_combo(page, names)
    cbox = bbox(combo)
    log.info("%s | open '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step, name, *cbox)
    _open_combo_with_fallback(combo, page)
    listbox = page.get_by_role("listbox").first
    option = listbox.get_by_role("option", name=value, exact=True)
    option.scroll_into_view_if_needed()
    obox = bbox(option)
    log.info("%s | choose '%s' @ (x=%.1f,y=%.1f,w=%.1f,h=%.1f)", step, value, *obox)
    option.click()
    try:
        expect(combo).to_have_text(re.compile(re.escape(value), re.I), timeout=3_000)
    except AssertionError:
        log.warning("%s | value not reflected proceeding", step)
    snap_and_annotate(page, [cbox, obox], f"{step}_selected")


# --- Year ---
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
            loc.wait_for(state="visible", timeout=1500);
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
                    return _first_visible(ctx, css, timeout_ms=250)  # small per-attempt wait
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
    # Prefer the non-disabled Accessible button (sometimes there is a disabled twin).
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
        # 1) Find & scroll the Accessible button into view
        acc_btn = _scan_all(page, acc_css, timeout_ms=14000)
        acc_btn.scroll_into_view_if_needed()

        # Move first, then click (pyautogui). Fallback to Patchright if pyautogui unavailable.
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
        # If the explicit button isn’t found, use your geometric fallback near the big "Press and hold" widget
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
            log.info("step_13_accessible_challenge_fallback | page.mouse click")
            page.mouse.click(x_css, y_css)
        snap_and_annotate(page, [(x_css - 10, y_css - 10, 20, 20)], "step_13_accessible_challenge_fallback")

    page.wait_for_timeout(9_000)
    press_again_css = [
        "[role='button'][aria-label='Press again']",
        "[role='button'][aria-label*='Press again' i]",
    ]
    try:
        pa_btn = _scan_all(page, press_again_css, timeout_ms=7000)
    except PWTimeoutError:
        # Fallback: role+text across frames if aria-label isn’t present
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


def handle_quick_note(page) -> bool:
    """Optionally dismiss 'A quick note about your account page' by clicking OK."""
    # Prefer detecting the heading (if present), but fall back to 'OK' button alone.
    try:
        # Try to spot the heading in any frame
        found_heading = False
        for ctx in [page, *page.frames]:
            hdr = ctx.get_by_role("heading", name=re.compile(r"quick note about your account page", re.I)).first
            if hdr.count():
                hdr.wait_for(state="visible", timeout=2000)
                found_heading = True
                break

        # Find an OK button (in any context)
        ok_btn = None
        for ctx in [page, *page.frames]:
            cand = ctx.get_by_role("button", name=re.compile(r"^OK$", re.I)).first
            if cand.count():
                cand.wait_for(state="visible", timeout=2000)
                ok_btn = cand
                break

        if ok_btn is None:
            if found_heading:
                log.info("Quick note detected but no OK button visible yet.")
            else:
                log.info("No 'Quick note' screen detected.")
            return False

        abox = bbox(ok_btn)
        ax = abox[0] + abox[2] / 2.0
        ay = abox[1] + abox[3] / 2.0
        if _HAS_PYAUTO:
            _pyauto_move_css(page, ax, ay, "step_15_quick_note_ok_move")
            time.sleep(0.4)
            _pyauto_click_css(page, ax, ay, "step_15_quick_note_ok_click")
        else:
            safe_click(ok_btn, "step_15_quick_note_ok_click_fallback_patchright", page)
        snap_and_annotate(page, [abox], "step_15_quick_note_ok_target")
        return True

    except PWTimeoutError:
        log.info("No 'Quick note' screen detected (timeout).")
        return False


def handle_stay_signed_in(page) -> bool:
    """Optionally dismiss 'Stay signed in?' by clicking No."""
    try:
        no_btn = None
        for ctx in [page, *page.frames]:
            cand = ctx.get_by_role("button", name=re.compile(r"^No$", re.I)).first
            if cand.count():
                cand.wait_for(state="visible", timeout=2000)
                no_btn = cand
                break
        if no_btn is None:
            log.info("No 'Stay signed in?' dialog detected.")
            return False

        nbox = bbox(no_btn)
        nx = nbox[0] + nbox[2] / 2.0
        ny = nbox[1] + nbox[3] / 2.0
        if _HAS_PYAUTO:
            _pyauto_move_css(page, nx, ny, "step_16_stay_signed_in_no_move")
            time.sleep(0.4)
            _pyauto_click_css(page, nx, ny, "step_16_stay_signed_in_no_click")
        else:
            safe_click(no_btn, "step_16_stay_signed_in_no_click_fallback_patchright", page)
        snap_and_annotate(page, [nbox], "step_16_stay_signed_in_no_target")
        return True

    except PWTimeoutError:
        log.info("No 'Stay signed in?' dialog detected (timeout).")
        return False


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

        # Calibrate CSS->screen mapping (NEW)
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
        time.sleep(5)

        try:
            solve_human_check(page)
        except PWTimeoutError:
            log.warning("Accessible challenge not found skipping as site may differ.")
            log.info(PWTimeoutError.message)
        time.sleep(10)

        handled_quick_note = handle_quick_note(page)
        handled_stay_signed_in = handle_stay_signed_in(page)
        log.info("OPTIONAL | quick_note=%s stay_signed_in=%s",
                 handled_quick_note, handled_stay_signed_in)

        log.info("ACCOUNT | email=%s | password=%s", EMAIL, PASSWORD)

        time.sleep(10)
        snap_and_annotate(page, [], "step_16_done")
        log.info("FLOW COMPLETE")

    finally:
        context.close()


if __name__ == "__main__":
    with sync_playwright() as p:
        run(p)
