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
from selenium.common.exceptions import SessionNotCreatedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

PROFILE_DIR = Path.home() / ".lms_selenium_profile"  # keeps SSO session between runs
LOG_FILE = Path("lms_progress.csv")
DATE_FMT = "%d %b %y"
START_XPATHS = [  # priority order; exact text so "Restart" / "Start date" don't match
    "//button[translate(normalize-space(.), 'START', 'start') = 'start']",
    "/html/body/div[2]/div/div/div[5]/div/div[2]/div/div/div[3]/div/div[2]/div[2]/div[2]/div/div[2]/button",
]
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
