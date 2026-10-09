"""
LMS training queue - SumTotal (PPD) assignments.

Opens each assignment from the exported .docx in one persistent, logged-in
Chrome session, in due-date order. You do the reading and the e-signature;
the script handles navigation, ordering and progress tracking (resumable).

Usage:
    pip install selenium python-docx
    python lms_queue.py POL.docx
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
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

PROFILE_DIR = Path.home() / ".lms_selenium_profile"  # keeps SSO session between runs
LOG_FILE = Path("lms_progress.csv")
DATE_FMT = "%d %b %y"
START_XPATHS = [
    "//button[contains(translate(normalize-space(.), 'START', 'start'), 'start')]",
    "/html/body/div[2]/div/div/div[5]/div/div[2]/div/div/div[3]/div/div[2]/div[2]/div[2]/div/div[2]/button",
]


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
            url = unwrap_urldefense(rels[rid].target_ref)
            aid = re.search(r"ActivityId%3D(\d+)", url)
            items.append(Item(
                title=cells[0].text.strip(),
                due=datetime.strptime(cells[1].text.strip(), DATE_FMT),
                url=url,
                activity_id=aid.group(1) if aid else "",
            ))
    return sorted(items, key=lambda i: i.due)


def load_done() -> set[str]:
    if not LOG_FILE.exists():
        return set()
    with LOG_FILE.open(encoding="utf-8") as f:
        return {r["activity_id"] for r in csv.DictReader(f) if r["status"] == "done"}


def log(item: Item, status: str) -> None:
    new = not LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["timestamp", "activity_id", "title", "due", "status"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), item.activity_id,
                    item.title, item.due.date(), status])


def find_start(driver, timeout: int = 20):
    """Return the Start button, trying the text-based XPath first, absolute one second."""
    for xp in START_XPATHS:
        try:
            return WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xp)))
        except TimeoutException:
            timeout = 3  # page is loaded by now; fallback only needs a short wait
    return None


def open_document(driver) -> bool:
    """Click Start, focus the document popup, wait for you to close it."""
    main_handle = driver.current_window_handle
    before = set(driver.window_handles)
    btn = find_start(driver)
    if btn is None:
        return False
    btn.click()
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


def main(docx_path: str) -> None:
    items = load_items(Path(docx_path))
    done = load_done()
    queue = [i for i in items if i.activity_id not in done]
    print(f"{len(items)} assignments, {len(done)} done, {len(queue)} remaining.\n")

    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    driver = webdriver.Chrome(options=opts)

    if queue:
        driver.get(queue[0].url)
        input("Log in in the browser window if prompted, then press Enter...")

    for n, item in enumerate(queue, 1):
        driver.get(item.url)
        print(f"[{n}/{len(queue)}] due {item.due:%d %b} | {item.title}")
        if not open_document(driver):
            print("  Start button not found (course/ILT or already started?) - handle manually.")
        ans = input("  Answer the read confirmation in the browser, then Enter = done | s = skip | q = quit: ").strip().lower()
        if ans == "q":
            break
        log(item, "skipped" if ans == "s" else "done")

    driver.quit()


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "POL.docx")
