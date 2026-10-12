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
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

PROFILE_DIR = Path.home() / ".lms_selenium_profile"  # keeps SSO session between runs
LOG_FILE = Path("lms_progress.csv")
DATE_FMT = "%d %b %y"
DEBUG_DIR = Path("lms_debug")  # page dumps when Start isn't found; local only (gitignored)

# Exact labels only (case-insensitive), so "Restart", "Start date" - and any Yes /
# completion / attestation control - can never match.
START_LABELS = ["start", "start activity", "start course", "launch", "launch activity", "launch course"]


def _start_xpath() -> str:
    lower = "translate(normalize-space({}), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
    attrs = [".", "@value", "@aria-label", "@title"]
    conds = " or ".join(f"{lower.format(a)} = '{lbl}'" for a in attrs for lbl in START_LABELS)
    return f"//*[self::button or self::a or self::input or @role='button'][{conds}]"


START_XPATHS = [  # priority order; searched in the page and in every (nested) iframe
    _start_xpath(),
    "/html/body/div[2]/div/div/div[5]/div/div[2]/div/div/div[3]/div/div[2]/div[2]/div[2]/div/div[2]/button",
]
MAX_FRAME_DEPTH = 3
# Read-only check of the activity page after you confirm. Adjust to SumTotal's actual status text.
COMPLETED_XPATH = "//*[not(self::script)][translate(normalize-space(text()), 'COMPLETED', 'completed') = 'completed']"
DONE_STATUSES = {"done", "done-unverified"}


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


def _search_frames(driver, depth: int = 0):
    """Depth-first search of the current frame and its child frames.

    On success the driver is left switched into the frame that holds the button.
    """
    for xp in START_XPATHS:
        for el in driver.find_elements(By.XPATH, xp):
            if el.is_displayed() and el.is_enabled():
                return el
    if depth < MAX_FRAME_DEPTH:
        for frame in driver.find_elements(By.CSS_SELECTOR, "iframe, frame"):
            driver.switch_to.frame(frame)
            el = _search_frames(driver, depth + 1)
            if el:
                return el
            driver.switch_to.parent_frame()
    return None


def _first_clickable(driver):
    driver.switch_to.default_content()
    try:
        return _search_frames(driver) or False
    except (StaleElementReferenceException, NoSuchFrameException):
        return False  # page still re-rendering; poll again


def find_start(driver, timeout: int = 20):
    """Return the Start button (driver switched into its frame), or None. One shared wait."""
    try:
        return WebDriverWait(driver, timeout).until(_first_clickable)
    except TimeoutException:
        driver.switch_to.default_content()
        return None


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


def is_completed(driver, url: str, timeout: int = 10) -> bool:
    """Reload the activity page and look for a Completed status. Read-only."""
    driver.get(url)
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, COMPLETED_XPATH)))
        return True
    except TimeoutException:
        return False


def open_document(driver, item: Item) -> bool:
    """Click Start, focus the document popup, wait for you to close it."""
    main_handle = driver.current_window_handle
    before = set(driver.window_handles)
    btn = find_start(driver)
    if btn is None:
        dump_debug(driver, item)
        return False
    click(driver, btn)
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, 20).until(lambda d: len(d.window_handles) > len(before))
        driver.switch_to.window((set(driver.window_handles) - before).pop())
        print("  Document open - read it, then close the popup.")
        WebDriverWait(driver, 4 * 3600, poll_frequency=1).until(
            lambda d: main_handle in d.window_handles and len(d.window_handles) == len(before))
    except TimeoutException:
        print("  No popup detected (or timed out) - continuing on the main page.")
    driver.switch_to.window(main_handle)
    return True


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
            if not open_document(driver, item):
                print("  Start button not found (course/ILT or already started?) - handle manually.")
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
