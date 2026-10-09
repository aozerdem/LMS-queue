# Handover: `lms_queue.py` (PPD SumTotal training queue)

**Owner:** Ahmet, Localization Engineer, Thermo Fisher / PPD CRG. Started 01.10.2026.
**Date of handover:** 09.10.2026

## Goal

Work through 48 assigned SumTotal LMS items (POL / SOP / WI documents plus a few COURSE / ILT / OC items) exported to `POL.docx`. Due dates run from 29 Oct 26 to 30 Dec 26.

## Current state

The script is working: login, SSO persistence and Start-click handling all run.

Setup on the Windows machine:

```
pip uninstall -y docx          # wrong package, Py2 code, causes "No module named 'exceptions'"
pip install python-docx selenium
python lms_queue.py POL.docx
```

## What the script does

1. Parses the hyperlink table in `POL.docx` (column 1 is the title with its link, column 2 is the due date in `%d %b %y` format).
2. Decodes the Proofpoint urldefense v3 wrappers back to the real `https://PPD.sumtotal.host/...ActivityId%3D<id>` URLs. All 48 items parse.
3. Sorts the items by due date.
4. Runs Chrome with a persistent profile at `~/.lms_selenium_profile`, so the SSO login survives between runs.
5. For each item:
   - Opens the URL.
   - Clicks **Start**. The primary locator is a text-based XPath; the user-supplied absolute XPath is the fallback.
   - Switches to the document popup and waits up to 4 hours for the user to close it.
   - Returns to the main window.
6. **The user clicks Yes on the "have you read this document?" confirmation manually**, then presses Enter in the terminal (`s` skips, `q` quits).
7. Logs each item to `lms_progress.csv` with a timestamp, activity_id, title, due date and status. Items marked done are skipped on the next run.

Absolute Start XPath supplied by the user:

`/html/body/div[2]/div/div/div[5]/div/div[2]/div/div/div[3]/div/div[2]/div[2]/div[2]/div/div[2]/button`

## Deliberate boundary: do not automate the "Yes" attestation

Ahmet asked twice for the script to click Yes automatically. The second time he said he had already read everything. I declined both times and kept that click manual. Reasons:

- The Yes click is a Part 11 e-signature on a GxP training record at a CRO, and it must be executed by the signer. A bot clicking Yes within a second of the popup closing, 48 times in a row, reads as falsification in the SumTotal audit trail, whether or not the reading happened.
- The queue itself includes POL-00392 (e-records / e-signatures), POL-00389 (Data Integrity), SOP-16948 (GDP), SOP-16957 (Acceptable Use of Technologies) and POL-00551 (AI acceptable use). A Selenium-driven SSO session that signs attestations conflicts with these directly.
- He is on probation in a quality-facing role.
- The time saved would be about 48 clicks, roughly two minutes.

**Fine to help with:** navigation, locators, popup and window handling, resume logic, CSV reporting, robustness fixes, and Start-button variants for COURSE-type items.

**Not fine:** clicking Yes or any other completion or attestation control, auto-closing an unread document, or faking dwell time.

## Known open issues and next steps

- **COURSE, ILT and OC items** (GCP Academy, Veeva QMS, Safety Awareness, Customer Delight, PLP checklist, etc.) probably use a different button or flow. The script currently reports "Start button not found" and leaves those to the user. Locators for them still need to be added if Ahmet supplies them.
- **Duplicate items to clarify with the manager or LMS admin:** SOP-16959 appears at v21 and v22, and SOP-16957 appears as both DOC and COURSE. Completing one may supersede the other.
- **Role-critical documents worth a real read:** SOP-18808, SOP-18810, SOP-16950, WI-06667 and SOP-16959.
