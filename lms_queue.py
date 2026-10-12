"""
LMS training queue - SumTotal (PPD) assignments.

Opens each assignment from the exported .docx in one persistent, logged-in
Chrome session, in due-date order. You do the reading and the e-signature;
the script handles navigation, ordering and progress tracking (resumable).

Usage:
    pip install selenium python-docx
    python lms_queue.py            # pick the .docx in a file dialog
    python lms_queue.py POL.docx   # or pass it directly
"""
from __future__ import annotations

import base64
import csv
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    NoSuchFrameException,
    SessionNotCreatedException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

PROFILE_DIR = Path.home() / ".lms_selenium_profile"  # keeps SSO session between runs
LOG_FILE = Path("lms_progress.csv")
DATE_FMT = "%d %b %y"
DEBUG_DIR = Path("lms_debug")  # page dumps when Start isn't found; local only (gitignored)

# Exact labels only (case-insensitive), so "Restart", "Start date" - and any Yes /
# completion / attestation control - can never match.
START_LABELS = ["start", "start activity", "start course", "launch", "launch activity", "launch course"]


def _button_xpath(labels: list[str]) -> str:
    """Button-like element whose text/value/aria-label/title equals one of labels (case-insensitive)."""
    lower = "translate(normalize-space({}), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
    attrs = [".", "@value", "@aria-label", "@title"]
    conds = " or ".join(f"{lower.format(a)} = '{lbl}'" for a in attrs for lbl in labels)
    return f"//*[self::button or self::a or self::input or @role='button'][{conds}]"


START_XPATHS = [  # priority order; searched in the page and in every (nested) iframe
    _button_xpath(START_LABELS),
    "/html/body/div[2]/div/div/div[5]/div/div[2]/div/div/div[3]/div/div[2]/div[2]/div[2]/div/div[2]/button",
]
MAX_FRAME_DEPTH = 3
# Start on an already-completed item opens a RESTART / REVIEW / CANCEL dialog. Only CANCEL is
# ever clicked (no credit, no signature); RESTART and REVIEW are never touched.
RESTART_XPATH = _button_xpath(["restart"])
CANCEL_XPATH = _button_xpath(["cancel"])
ATTEND_TEXT_XPATH = ("//*[not(self::script) and not(self::style)][text()[contains(translate(., "
                     "'ATEND', 'atend'), 'attend')]]")
# Visible status label under the activity title (e.g. green "ATTENDED"). Exact text only, so
# "Unattended" / "Not attended" don't count. Read-only: nothing is ever clicked on a match.
DONE_LABELS = ["attended", "completed"]
STATUS_XPATH = (
    "//*[not(self::script) and not(self::style)][text()["
    + " or ".join("translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                  f"'abcdefghijklmnopqrstuvwxyz') = '{lbl}'" for lbl in DONE_LABELS)
    + "]]"
)
STATUS_GRACE = 1.0  # s to keep looking for the status label after Start has rendered;
# raise it if an attended item ever gets its Start clicked (label rendering later than that)
POLL = 0.2  # s between page checks
DONE_STATUSES = {"done", "done-unverified", "already-attended"}


@dataclass
class Item:
    title: str
    due: datetime
    url: str
    activity_id: str


def unwrap_urldefense(url: str) -> str:
    """Decode a Proofpoint urldefense v3 link back to the original URL."""
    m = re.match(r"https://urldefense\.com/v3/__(.+?)__;(.*?)!!", url)
    if not m:
        return url
    wrapped, enc = m.groups()
    pad = "=" * (-len(enc) % 4)
    repl = iter(base64.urlsafe_b64decode(enc + pad).decode("utf-8"))
    decoded = re.sub(r"\*", lambda _: next(repl, "*"), wrapped)
    return decoded.replace("https:/", "https://", 1)


def load_items(docx_path: Path) -> list[Item]:
    doc = Document(docx_path)
    rels = doc.part.rels
    items: list[Item] = []
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) < 2:
                continue
            link = cells[0]._tc.find(".//" + qn("w:hyperlink"))
            rid = link.get(qn("r:id")) if link is not None else None
            if not rid or rid not in rels:
                continue
            title = cells[0].text.strip()
            try:
                due = datetime.strptime(cells[1].text.strip(), DATE_FMT)
            except ValueError:
                print(f"WARNING: unparseable due date {cells[1].text.strip()!r} - skipping {title}")
                continue
            url = unwrap_urldefense(rels[rid].target_ref)
            aid = re.search(r"ActivityId%3D(\d+)", url)
            items.append(Item(title=title, due=due, url=url,
                              activity_id=aid.group(1) if aid else ""))
    return sorted(items, key=lambda i: i.due)


def load_done() -> set[str]:
    if not LOG_FILE.exists():
        return set()
    with LOG_FILE.open(encoding="utf-8") as f:
        return {r["activity_id"] for r in csv.DictReader(f) if r["status"] in DONE_STATUSES}


def log(item: Item, status: str) -> None:
    new = not LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["timestamp", "activity_id", "title", "due", "status"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), item.activity_id,
                    item.title, item.due.date(), status])


def _search_frames(driver, xpaths: list[str], depth: int = 0):
    """Depth-first search of the current frame and its child frames for a visible match.

    On success the driver is left switched into the frame that holds the element.
    """
    for xp in xpaths:
        for el in driver.find_elements(By.XPATH, xp):
            if el.is_displayed() and el.is_enabled():
                return el
    if depth < MAX_FRAME_DEPTH:
        for frame in driver.find_elements(By.CSS_SELECTOR, "iframe, frame"):
            driver.switch_to.frame(frame)
            el = _search_frames(driver, xpaths, depth + 1)
            if el:
                return el
            driver.switch_to.parent_frame()
    return None


def _find(driver, xpaths: list[str]):
    driver.switch_to.default_content()
    try:
        return _search_frames(driver, xpaths)
    except (StaleElementReferenceException, NoSuchFrameException):
        return None  # page still re-rendering; poll again


def check_page(driver, timeout: int = 20):
    """Wait for the activity page to render and classify it. Read-only.

    Returns ("done", label_text), ("start", button) with the driver in the button's frame,
    or (None, None). Attended items still show Start (relaunch), and the status label can
    render after it - so once Start appears, keep looking for the label for STATUS_GRACE s.
    """
    deadline = time.monotonic() + timeout
    start_seen_at = None
    while time.monotonic() < deadline:
        label = _find(driver, [STATUS_XPATH])
        if label:
            text = label.text.strip() or "attended"
            driver.switch_to.default_content()
            if start_seen_at is not None:  # tells you how low STATUS_GRACE can safely go
                text += f" (label {time.monotonic() - start_seen_at:.1f}s after Start)"
            return "done", text
        if start_seen_at is None and _find(driver, START_XPATHS):
            start_seen_at = time.monotonic()
        if start_seen_at is not None and time.monotonic() - start_seen_at >= STATUS_GRACE:
            btn = _find(driver, START_XPATHS)  # re-find: the page may have re-rendered
            if btn:
                return "start", btn
            start_seen_at = None
        time.sleep(POLL)
    driver.switch_to.default_content()
    return None, None


def click(driver, el) -> None:
    """Normal click; JS click if an overlay/spinner intercepts it."""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
    try:
        el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException):
        driver.execute_script("arguments[0].click();", el)


def dump_debug(driver, item: Item) -> None:
    """Save a screenshot and every frame's HTML, and list clickable-looking elements."""
    DEBUG_DIR.mkdir(exist_ok=True)
    stem = DEBUG_DIR / (item.activity_id or "unknown")
    driver.switch_to.default_content()
    driver.save_screenshot(f"{stem}.png")
    print(f"  Debug: screenshot + page HTML saved under {DEBUG_DIR}/ - send me the button list below.")

    def walk(path: str, depth: int) -> None:
        Path(f"{stem}_frame{path or '-top'}.html").write_text(driver.page_source, encoding="utf-8")
        els = driver.find_elements(By.CSS_SELECTOR, "button, a, input[type=button], input[type=submit], [role=button]")
        shown = [e for e in els if e.is_displayed()]
        print(f"  frame {path or 'top'}: {len(shown)} visible clickable elements")
        for e in shown[:25]:
            label = (e.text or e.get_attribute("value") or e.get_attribute("aria-label")
                     or e.get_attribute("title") or "").strip().replace("\n", " ")[:50]
            print(f"    <{e.tag_name}> {label!r} id={e.get_attribute('id')!r} class={(e.get_attribute('class') or '')[:60]!r}")
        if depth < MAX_FRAME_DEPTH:
            for i, frame in enumerate(driver.find_elements(By.CSS_SELECTOR, "iframe, frame")):
                driver.switch_to.frame(frame)
                walk(f"{path}-{i}", depth + 1)
                driver.switch_to.parent_frame()

    try:
        walk("", 0)
    except Exception as e:  # diagnostics must never kill the run
        print(f"  Debug dump incomplete: {e!r}")
    driver.switch_to.default_content()


def is_completed(driver, url: str, timeout: int = 15) -> bool:
    """Reload the activity page and look for the Attended/Completed label. Read-only."""
    driver.get(url)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if _find(driver, [STATUS_XPATH]):
            driver.switch_to.default_content()
            return True
        time.sleep(POLL)
    driver.switch_to.default_content()
    return False


def _restart_dialog_cancel(driver):
    """CANCEL button of a visible RESTART/REVIEW/CANCEL dialog (driver in its frame), or None."""
    if not _find(driver, [RESTART_XPATH]):
        driver.switch_to.default_content()
        return None
    for el in driver.find_elements(By.XPATH, CANCEL_XPATH):  # same frame as RESTART
        if el.is_displayed() and el.is_enabled():
            return el
    driver.switch_to.default_content()
    return None


def _report_missed_label(driver) -> None:
    """The label check missed an attended item: print what the status element looks like."""
    el = _find(driver, [ATTEND_TEXT_XPATH])
    html = el.get_attribute("outerHTML")[:300] if el else "(no element with 'attend' text found)"
    driver.switch_to.default_content()
    print(f"  Note: status label not recognised. Send this to fix it: {html}")


def open_document(driver, btn) -> str:
    """Click Start, then either cancel the RESTART dialog ("attended") or focus the document
    popup and wait for you to close it ("opened"). "none" if neither appears."""
    main_handle = driver.current_window_handle  # frame switches don't change the window handle
    before = set(driver.window_handles)
    click(driver, btn)
    driver.switch_to.default_content()

    deadline = time.monotonic() + 20
    while time.monotonic() < deadline and len(driver.window_handles) == len(before):
        cancel = _restart_dialog_cancel(driver)
        if cancel:
            click(driver, cancel)
            driver.switch_to.default_content()
            print("  RESTART/REVIEW dialog - already completed. Clicked CANCEL, moving on.")
            _report_missed_label(driver)
            return "attended"
        time.sleep(POLL)
    if len(driver.window_handles) == len(before):
        print("  No popup detected - continuing on the main page.")
        return "none"

    try:
        driver.switch_to.window((set(driver.window_handles) - before).pop())
        print("  Document open - read it, then close the popup.")
        WebDriverWait(driver, 4 * 3600, poll_frequency=1).until(
            lambda d: main_handle in d.window_handles and len(d.window_handles) == len(before))
    except TimeoutException:
        print("  Popup still open after 4 h - continuing on the main page.")
    driver.switch_to.window(main_handle)
    return "opened"


def pick_docx() -> Path | None:
    """Ask for the assignments .docx in a native file dialog. None if cancelled."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # otherwise the dialog can open behind the console
    path = filedialog.askopenfilename(
        parent=root,
        title="Select the LMS assignments .docx (table with links and due dates)",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
    )
    root.destroy()
    return Path(path) if path else None


def main(docx_path: Path) -> None:
    if not docx_path.is_file():
        sys.exit(f"File not found: {docx_path.resolve()}")
    items = load_items(docx_path)
    if not items:
        sys.exit(f"No linked assignments found in {docx_path.name} - is it the exported table?")
    done = load_done()
    queue = [i for i in items if i.activity_id not in done]
    print(f"{len(items)} assignments, {len(done)} done, {len(queue)} remaining.\n")

    if not queue:
        return

    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    try:
        driver = webdriver.Chrome(options=opts)
    except SessionNotCreatedException as e:
        sys.exit(f"Could not start Chrome - is another Chrome window using {PROFILE_DIR}? "
                 f"Close it and retry.\n{e.msg}")

    unverified: list[Item] = []
    try:
        driver.get(queue[0].url)
        input("Log in in the browser window if prompted, then press Enter...")

        for n, item in enumerate(queue, 1):
            driver.get(item.url)
            print(f"[{n}/{len(queue)}] due {item.due:%d %b} | {item.title}")
            state, found = check_page(driver)
            if state == "done":
                print(f"  Status '{found}' - already done, moving on.")
                log(item, "already-attended")
                continue
            if state == "start":
                if open_document(driver, found) == "attended":
                    log(item, "already-attended")
                    continue
            else:
                dump_debug(driver, item)
                print("  Neither a status label nor a Start button found - handle manually.")
            ans = input("  Answer the read confirmation in the browser, then Enter = done | s = skip | q = quit: ").strip().lower()
            if ans == "q":
                break
            if ans == "s":
                log(item, "skipped")
                continue
            if is_completed(driver, item.url):
                log(item, "done")
                continue
            ans = input("  'Completed' not found on the activity page. Mark done anyway? [y/N]: ").strip().lower()
            if ans == "y":
                log(item, "done-unverified")
                unverified.append(item)
            else:
                log(item, "not-completed")  # stays in the queue for the next run
    finally:
        driver.quit()

    if unverified:
        print("\nMarked done WITHOUT on-page confirmation - check these in SumTotal:")
        for item in unverified:
            print(f"  {item.activity_id}  {item.title}")


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else pick_docx()
    if path is None:
        sys.exit("No file selected.")
    main(path)
