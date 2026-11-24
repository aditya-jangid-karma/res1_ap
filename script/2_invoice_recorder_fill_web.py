"""TASK 2 -- fill the Nexus (Bottomline) portal from invoices.xlsx.

Fills ONLY three values, taken from the Excel row whose "Image Name" matches
the image open in the portal:
    1. Property        (column "Property")
    2. Vendor          (column "Vendor")
    3. Invoice number  (column "Invoice Number")

How it works
  1. Logs in (credentials from .env).
  2. Same steps as open_index_selected.py:
       Image Management -> Images to be Indexed -> Records per page 200
       -> last page -> select all -> Index Selected.
     The first selected image opens in the indexing form.
  3. Reads the Image Name shown in the form, finds the Excel row with that
     image name, and PASTES its property / vendor / invoice number. The portal
     shows a dropdown; the match is clicked. If the full name finds nothing,
     trailing words are dropped until the dropdown shows a match
     ("Seneca Creek Senior Living" -> "Seneca Creek").
  4. Reads the boxes back to prove the values stuck.
  By default only the image that opens is filled. With --walk it presses
  "Next" and repeats for the other selected images.

Rows processed (column "Status"):
  APPROVED  -> always.
  PENDING   -> only if "Qwen Status" and "Property Status" are not NEEDS_REVIEW.
After a row the Status becomes FILLED or SKIPPED (reason in "Recorder Note").

SAFETY: it NEVER presses Convert to Invoice, Save, Index, Save as Exception or
Delete. A human reviews each filled form. It never clicks a dropdown entry
unless exactly one entry matches.

Run (use the local venv, not the one on the network drive):
    C:\\res1_venv\\Scripts\\python.exe script\\2_invoice_recorder_fill_web.py --dry-run
    C:\\res1_venv\\Scripts\\python.exe script\\2_invoice_recorder_fill_web.py

.env (project root): BOTTOMLINE_URL, BOTTOMLINE_USERNAME, BOTTOMLINE_PASSWORD
The login session is saved to .bottomline_session.json (treat it as a password).
"""
import argparse

import json
import os
import re
import shutil
import sys
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

from fill_plan import package_candidates, package_verdict
from queue_navigator import (build_queue_index, names_match, next_page_available,
                             normalise_name, open_invoice, return_to_queue,
                             wait_for_queue_page, wait_for_queue_rows)

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
DEFAULT_XLSX = PROJECT / "invoices.xlsx"
SELECTORS_FILE = HERE / "selectors.json"
SESSION_FILE = PROJECT / ".bottomline_session.json"

PROCESS_STATUSES = ("PENDING", "APPROVED")
DROPDOWN_WAIT_MS = 2500
PDF_DIR = PROJECT / "invoice"                 # where Task 1 read the PDFs from
ERROR_DIR = PROJECT / "error_invoices"        # invoices that failed to fill
DONE_DIR = PROJECT / "filled_invoices"        # invoices filled without error
PNF_DIR = PROJECT / "property not found"      # property not in the portal
# fill_property() ends its reason with this text when no package has the property
PNF_MARK = "was not found under any integration package"
VNF_DIR = PROJECT / "vendor not found"        # vendor not in the portal
VNF_MARK = "was not found in the portal vendor list"   # see fill_vendor()
NOT_OPENED_NOTES = ("not found in the portal queue",
                    "could not open/verify the invoice")   # set in main()
DUP_DIR = PROJECT / "duplicate invoices"      # portal says it is a duplicate
DUP_MARK = "looks like a duplicate in the portal"      # see duplicate_warning()


def move_invoice(rec, dest_dir, note=None, log=print):
    """MOVE the invoice PDF from invoice/ into dest_dir. For errors a .txt with
    the reason is written beside it. Returns True only if the PDF is safely in
    dest_dir (so the caller may then delete the portal copy). Never raises."""
    try:
        dest_dir.mkdir(exist_ok=True)
        name = rec["image_name"]
        src = next((p for p in PDF_DIR.glob("*")
                    if p.suffix.lower() == ".pdf" and same_image(p.stem, name)),
                   None)
        if src is None:
            log(f"  could not move invoice: no PDF named '{name}' in {PDF_DIR}")
            return False
        target = dest_dir / src.name
        shutil.copy2(src, target)                  # copy, verify, then remove
        if not target.exists() or target.stat().st_size != src.stat().st_size:
            log("  copy check failed; original left in place.")
            return False
        src.unlink()
        if note:
            (dest_dir / f"{src.stem}.txt").write_text(
                f"{src.name}\nvendor: {rec['vendor']}\n"
                f"property: {rec['property']}\n"
                f"invoice number: {rec['invoice_number']}\nerror: {note}\n",
                encoding="utf-8")
        log(f"  invoice moved -> {target}")
        return True
    except Exception as error:
        log(f"  could not move invoice ({error})")
        return False


def delete_current_invoice(page, selectors, rec, log):
    """Delete the invoice open in the indexing form from the portal. Only runs
    for an ERROR invoice whose PDF is already saved in ERROR_DIR, and only if
    the form on screen really is that invoice. The 'Delete' control and its
    confirmation dialog are matched by visible text and have NOT been recorded
    in selectors.json -- test on one invoice first. Returns True if the form
    closed."""
    shown = shown_image_name(page, selectors)
    if not shown or not same_image(shown, rec["image_name"]):
        log(f"  NOT deleting: form shows '{shown}', expected '{rec['image_name']}'.")
        return False
    try:
        button = page.get_by_text("Delete", exact=True).locator("visible=true").first
        button.click(timeout=5000)
        confirm = page.locator(".x-message-box, .x-window").get_by_text(
            re.compile(r"^(Yes|OK)$")).locator("visible=true").first
        confirm.click(timeout=5000)
        page.locator(selectors["verify"]["image_name"]).first.wait_for(
            state="detached", timeout=15000)
        log("  deleted from the portal.")
        return True
    except Exception as error:
        log(f"  portal delete did not complete ({error}); check the portal.")
        return False


def finish_row(page, selectors, rec, note, delete_errors, log):
    """After a row is filled (note None) or failed (note = reason): file the
    PDF. Success -> DONE_DIR, portal copy kept. Error -> ERROR_DIR, then (only
    with --delete-errors, and only once the PDF is safe) delete from portal."""
    if not note:
        move_invoice(rec, DONE_DIR, None, log)
        return
    if note in NOT_OPENED_NOTES:         # never reached the form: nothing is wrong
        log("  PDF left in invoice/ (the image was not opened in the portal).")
        return                           # with the invoice, so do not move it
    if PNF_MARK in note:                 # property missing: park the PDF, keep
        move_invoice(rec, PNF_DIR, note, log)   # the portal copy, go to the next
        return
    if VNF_MARK in note:                 # vendor missing: park the PDF the same way
        move_invoice(rec, VNF_DIR, note, log)
        return
    if DUP_MARK in note:                 # duplicate: park it, keep portal copy
        move_invoice(rec, DUP_DIR, note, log)
        return
    if move_invoice(rec, ERROR_DIR, note, log) and delete_errors:
        delete_current_invoice(page, selectors, rec, log)


# --------------------------------------------------------------- spreadsheet
def same_image(a, b):
    """Image names match exactly, or differ only by a trailing '_<digits>'
    stamp (the sheet has '..._1793783728', the portal shows no stamp)."""
    if names_match(a, b):
        return True
    strip = lambda t: re.sub(r"_\d{8,}$", "", str(t).strip())
    return names_match(strip(a), strip(b))


def _key(header):
    """'Invoice Due Date' -> 'invoice_due_date'."""
    cleaned = "".join(c if (c.isalnum() or c.isspace()) else " "
                      for c in str(header))
    return "_".join(cleaned.lower().split())


def column_map(sheet):
    return {_key(c.value): c.column for c in sheet[1] if c.value}


def cell(sheet, columns, row, key):
    col = columns.get(key)
    if not col:
        return ""
    value = sheet.cell(row=row, column=col).value
    return "" if value is None else str(value).strip()


def set_cell(sheet, columns, row, key, header, value):
    if key not in columns:                      # add the column if missing
        columns[key] = sheet.max_column + 1
        sheet.cell(row=1, column=columns[key], value=header)
    sheet.cell(row=row, column=columns[key], value=value)


def save_workbook(workbook, path):
    try:
        workbook.save(path)
    except PermissionError:
        print(f"  !! Could not save {path} -- is it open in Excel? Close it.")


def rows_to_process(sheet, columns, only_image=None):
    """[(row_number, record)] for the rows this script should fill. With
    only_image, just the row with that Image Name (any Status)."""
    jobs = []
    for r in range(2, sheet.max_row + 1):
        name = cell(sheet, columns, r, "image_name")
        if not name:
            continue
        status = cell(sheet, columns, r, "status").upper()
        if only_image:
            if not same_image(name, only_image):
                continue
        elif status not in PROCESS_STATUSES:
            continue
        elif status == "PENDING" and "NEEDS_REVIEW" in (
                cell(sheet, columns, r, "qwen_status").upper(),
                cell(sheet, columns, r, "property_status").upper()):
            continue
        jobs.append((r, {k: cell(sheet, columns, r, k) for k in (
            "image_name", "vendor", "property", "invoice_number",
            "integration_package")}))
    return jobs


def split_sheet_duplicates(sheet, columns, jobs):
    """Pre-portal duplicate check. A job is a duplicate when another row with the
    same vendor + invoice number is already FILLED, or comes earlier in this run.
    Returns (jobs_to_fill, [(row, rec, note)]). Blank invoice numbers are never
    treated as duplicates."""
    key = lambda rec: (normalise_name(rec["vendor"]),
                       normalise_name(rec["invoice_number"]))
    seen = set()
    for r in range(2, sheet.max_row + 1):
        if cell(sheet, columns, r, "status").upper() == "FILLED":
            k = (normalise_name(cell(sheet, columns, r, "vendor")),
                 normalise_name(cell(sheet, columns, r, "invoice_number")))
            if k[0] and k[1]:
                seen.add(k)
    keep, dups = [], []
    for r, rec in jobs:
        k = key(rec)
        if k[0] and k[1] and k in seen:
            dups.append((r, rec, f"invoice number {DUP_MARK}: same vendor and "
                         f"invoice number '{rec['invoice_number']}' already in "
                         "the spreadsheet"))
            continue
        if k[0] and k[1]:
            seen.add(k)
        keep.append((r, rec))
    return keep, dups


def missing_fields(rec):
    return [k for k in ("vendor", "property") if not rec[k]]


# ------------------------------------------------------------------- portal
def visible_items(page, selectors):
    return page.locator(selectors["dropdown_item_selector"])


def paste_into(page, selector, text):
    """Paste `text` into a box, then nudge it (Space+Backspace) so the
    portal's autocomplete notices and opens its dropdown."""
    box = page.locator(selector).first
    box.click()
    box.fill("")
    box.fill(str(text))
    page.keyboard.press("Space")
    page.keyboard.press("Backspace")


def word_prefixes(text):
    """'Seneca Creek Senior Living' -> that, 'Seneca Creek Senior',
    'Seneca Creek', 'Seneca'. The portal's names are often shorter than the
    sheet's, and its dropdown matches from the start of the name."""
    words = str(text).replace(",", " ").split()
    return [" ".join(words[:n]) for n in range(len(words), 0, -1)]


def dropdown_options(page, selectors):
    """Text of every VISIBLE dropdown item after a short wait."""
    items = visible_items(page, selectors)
    try:
        items.first.wait_for(state="visible", timeout=DROPDOWN_WAIT_MS)
    except Exception:
        return []
    page.wait_for_timeout(250)                  # let the list finish filtering
    return [items.nth(i).inner_text().strip() for i in range(items.count())]


def fill_autocomplete(page, selectors, field, text, log):
    """Paste `text` in an autocomplete box and click the match.

    Returns (verdict, options): 'confirmed' (clicked the single match),
    'ambiguous' (several, none chosen) or 'not_here' (nothing)."""
    selector = selectors["fill"][field]
    paste_into(page, selector, text)
    options = dropdown_options(page, selectors)
    verdict, index = package_verdict(len(options)), 0
    if verdict == "ambiguous":                  # take an exact name if unique
        exact = [i for i, o in enumerate(options)
                 if normalise_name(o) == normalise_name(text)]
        if len(exact) == 1:
            verdict, index = "confirmed", exact[0]
    if verdict == "confirmed":
        visible_items(page, selectors).nth(index).click()
        try:                                    # continue the moment it is set
            page.wait_for_function(
                "s => { const e = document.querySelector(s); return e && e.value; }",
                arg=selector, timeout=2000)
        except Exception:
            pass
        if not page.locator(selector).first.input_value():
            log(f"  {field}: clicked '{text}' but the box is still empty.")
            return "not_here", options
    return verdict, options


def clear_box(page, selectors, field):
    try:
        page.locator(selectors["fill"][field]).first.fill("")
    except Exception:
        pass


def fill_property(page, selectors, rec, log):
    """Select the property. The package already on the form is tried first and
    left alone; other packages only if the property is not found under it.
    Returns None on success or a reason string."""
    try:
        current = page.locator(selectors["fill"]["integration_package"]
                               ).first.input_value(timeout=3000).strip()
    except Exception:
        current = ""
    packages = ([current] if current else []) + [
        p for p in package_candidates(rec.get("integration_package"))
        if p.upper() != current.upper()]
    for package in packages:
        if package.upper() != current.upper():
            verdict, _ = fill_autocomplete(page, selectors, "integration_package",
                                           package, log)
            if verdict != "confirmed":
                log(f"  package {package}: could not select it ({verdict}).")
                continue
            current = package
        for text in word_prefixes(rec["property"]):
            verdict, options = fill_autocomplete(
                page, selectors, "property", text, log)
            if verdict == "confirmed":
                log(f"  package {package}; property '{text}' set "
                    f"(sheet says '{rec['property']}').")
                return None
            if verdict == "ambiguous":
                return (f"property '{text}' matches several entries under "
                        f"{package}: {options[:5]} -- a human must pick")
        clear_box(page, selectors, "property")
        log(f"  property '{rec['property']}' not found under {package}.")
    return (f"property '{rec['property']}' was not found under any "
            "integration package")


def fill_vendor(page, selectors, rec, log):
    for text in word_prefixes(rec["vendor"]):
        verdict, options = fill_autocomplete(page, selectors, "vendor", text, log)
        if verdict == "confirmed":
            log(f"  vendor set via '{text}'.")
            return None
        if verdict == "ambiguous":
            return f"vendor '{text}' matches several entries {options[:5]}"
        clear_box(page, selectors, "vendor")
    return f"vendor '{rec['vendor']}' {VNF_MARK}"


def fill_invoice_number(page, selectors, rec, log):
    selector = selectors["fill"]["invoice_number"]
    paste_into(page, selector, rec["invoice_number"])
    page.locator(selector).first.press("Tab")
    log(f"  invoice_number = {rec['invoice_number']}")
    return duplicate_warning(page, selector, log)


_DUP_RE = re.compile(r"duplicat|already (exist|been|entered|used|recorded)", re.I)
_POPUP_JS = """() => [...document.querySelectorAll(
    '.x-message-box, .x-window, .x-tip, .x-form-error-msg, .x-form-invalid-tip')]
    .filter(e => e.offsetWidth || e.offsetHeight).map(e => e.innerText.trim())"""


def duplicate_warning(page, invoice_selector, log):
    """After the invoice number is entered, look for the portal telling us it
    is a duplicate (a dialog, tooltip or error text mentioning 'duplicate' /
    'already exists'). Returns a reason containing DUP_MARK, else None. The
    portal's real wording is NOT yet recorded -- check the first live run's
    log line 'portal popup:' and adjust _DUP_RE if needed. A dialog is closed
    with No/Cancel/OK/Close (never Yes), so nothing is forced through."""
    page.wait_for_timeout(1200)                  # let the check/dialog appear
    try:
        texts = [t for t in page.evaluate(_POPUP_JS) if t]
    except Exception:
        return None
    if texts:
        log(f"  portal popup: {texts}")
    hit = next((t for t in texts if _DUP_RE.search(t)), None)
    if hit is None:
        return None
    for label in ("No", "Cancel", "OK", "Close"):
        try:
            btn = page.locator(".x-message-box, .x-window").get_by_text(
                label, exact=True).locator("visible=true").first
            btn.click(timeout=1500)
            break
        except Exception:
            continue
    return f"invoice number {DUP_MARK}: {hit[:120]}"


def verify_filled(page, selectors, rec, log):
    """Read the boxes back. Returns a reason string, or None if all is well."""
    got = {}
    for field in ("integration_package", "property", "vendor", "invoice_number"):
        try:
            got[field] = page.locator(selectors["fill"][field]).first.input_value(
                timeout=3000).strip()
        except Exception:
            got[field] = ""
    log(f"  read back from portal: {got}")
    problems = []
    if not got["property"]:
        problems.append("property box is empty")
    if not got["vendor"]:
        problems.append("vendor box is empty")
    if rec["invoice_number"] and (normalise_name(got["invoice_number"])
                                  != normalise_name(rec["invoice_number"])):
        problems.append(f"invoice number is '{got['invoice_number']}', "
                        f"expected '{rec['invoice_number']}'")
    return "; ".join(problems) if problems else None


def fill_one(page, selectors, rec, log):
    """Fill the open invoice. Returns None on success or a reason string."""
    reason = fill_property(page, selectors, rec, log)
    if reason:
        return reason
    reason = fill_vendor(page, selectors, rec, log)
    if reason:
        return reason
    if rec["invoice_number"]:                   # blank = vendor+property only
        reason = fill_invoice_number(page, selectors, rec, log)
        if reason:
            return reason
    return verify_filled(page, selectors, rec, log)


def set_page_size(page, selectors, log, size_text="200"):
    """Records per page = 200 on the 'Images to be Indexed' screen."""
    try:
        wait_for_queue_rows(page, selectors, log)
        boxes = page.locator('input[id^="combo-"][id$="-inputEl"]')
        for i in range(boxes.count()):
            box = boxes.nth(i)
            if box.is_visible() and box.input_value().strip().isdigit():
                if box.input_value().strip() == size_text:
                    return True
                box.click()
                page.locator(selectors["dropdown_item_selector"],
                             has_text=size_text).first.click()
                try:                            # until the grid has reloaded
                    page.wait_for_function(
                        "n => document.querySelectorAll('tr.x-grid-row').length > n",
                        arg=25, timeout=15000)
                except Exception:
                    page.wait_for_timeout(2000)
                log(f"Records per page = {size_text}.")
                return True
    except Exception as error:
        log(f"Could not set records per page ({error}).")
    return False


_ROW_NAMES_JS = """sel => [...document.querySelectorAll(sel)].map(r => {
    for (const td of r.querySelectorAll('td')) {
        const t = td.innerText.trim(); if (t) return t; }
    return ''; })"""


def read_queue_fast(page, selectors, log):
    """Image name -> queue page number, reading each page in ONE browser call
    (the old reader made hundreds of calls per page). Returns the index."""
    entries, number = [], 1
    while True:
        names = page.evaluate(_ROW_NAMES_JS, selectors["queue"]["grid_row"])
        entries += [(n, number) for n in names]
        if not next_page_available(page, selectors):
            break
        first = names[0] if names else ""
        page.locator(selectors["queue"]["next_page"]).first.click()
        if wait_for_queue_page(page, selectors, number + 1) != number + 1:
            raise RuntimeError(f"Queue paging stuck after page {number}.")
        number += 1
        try:                                    # wait for the new rows to draw
            page.wait_for_function(
                "a => { const r = document.querySelector(a[0]); "
                "return r && r.innerText.trim() !== a[1]; }",
                arg=[selectors["queue"]["grid_row"] + " td", first], timeout=8000)
        except Exception:
            page.wait_for_timeout(800)
    index, _ = build_queue_index(entries)
    log(f"Queue read: {len(index)} images on {number} page(s).")
    return index


def run_index_steps(page, selectors, log):
    """The steps from open_index_selected.py: 200 per page -> last page ->
    select all -> Index Selected. Ends with the first image open in the form."""
    if not set_page_size(page, selectors, log):
        return False
    box = selectors["queue"]["page_number_box"]
    before = page.locator(box).first.input_value()
    page.locator(".x-tbar-page-last").first.click()
    try:                                        # until the page number changes
        page.wait_for_function(
            "a => document.querySelector(a[0]).value !== a[1]",
            arg=[box, before], timeout=10000)
    except Exception:
        pass                                    # already on the last page
    page.wait_for_timeout(500)
    log("Last page: " + page.locator(box).first.input_value())
    page.locator(".x-column-header-checkbox, .x-column-header-first"
                 ).first.click()
    try:
        page.locator(".x-grid-item-selected").first.wait_for(
            state="visible", timeout=5000)
    except Exception:
        pass
    count = page.locator(".x-grid-item-selected").count()
    log(f"Selected rows: {count}")
    if count == 0:
        return False
    page.locator(selectors["queue"]["index_selected"]).first.click()
    try:
        page.locator(selectors["verify"]["image_name"]).first.wait_for(
            state="visible", timeout=30000)
    except Exception:
        return False
    try:                                        # form fully loaded (values set)
        page.wait_for_function(
            "a => a.every(s => { const e = document.querySelector(s); "
            "return e && e.value; })",
            arg=[selectors["verify"]["image_name"],
                 selectors["fill"]["integration_package"]], timeout=15000)
    except Exception:
        page.wait_for_timeout(2500)
    return True


def shown_image_name(page, selectors):
    """The Image Name currently shown in the indexing form."""
    try:
        return page.locator(selectors["verify"]["image_name"]).first.input_value(
            timeout=5000).strip()
    except Exception:
        return ""


def click_next(page):
    """'Next' in the left bar of the indexing form (only used with --walk)."""
    try:
        link = page.get_by_text("Next", exact=True).first
        if link.is_visible():
            link.click()
            return True
    except Exception:
        pass
    return False


def on_login_page(page):
    return page.locator("input[type='password']").count() > 0


def login(page, url, log, auto=False):
    """Open the portal and sign in with the .env credentials. Normal mode: a
    human finishes any 2FA and presses ENTER. --auto: wait up to ~40 s for the
    login form to disappear, else stop."""
    page.goto(url)
    try:                                        # login form (or saved session)
        page.locator("input[type='password'], .x-btn-inner").first.wait_for(
            state="visible", timeout=8000)
    except Exception:
        pass
    user, pwd = os.getenv("BOTTOMLINE_USERNAME"), os.getenv("BOTTOMLINE_PASSWORD")
    if user and pwd and on_login_page(page):
        try:
            page.locator("input[type='text'], input[type='email']").first.fill(user)
            page.locator("input[type='password']").first.fill(pwd)
            page.keyboard.press("Enter")
        except Exception as error:
            log(f"Auto-fill of login failed ({error}); log in by hand.")
    # Wait for the login form to go away. If it does, the login worked (no 2FA)
    # and there is nothing to press ENTER for.
    for _ in range(20 if auto else 10):
        if not on_login_page(page):
            log("Logged in.")
            return True
        page.wait_for_timeout(2000)
    if auto:
        log("STOPPED: still on the login page -- wrong login, or a 2FA code "
            "is needed. Run once WITHOUT --auto and finish login by hand.")
        return False
    print("\n>>> Finish login / 2FA in the browser if asked.")
    input("Press ENTER once you can see the portal... ")
    return True


# --------------------------------------------------------------------- main
def main():
    load_dotenv(PROJECT / ".env")
    ap = argparse.ArgumentParser(
        description="TASK 2: fill invoices.xlsx rows into the Nexus portal "
                    "(never converts/saves).")
    ap.add_argument("--xlsx", type=Path,
                    default=Path(os.getenv("XLSX_PATH") or DEFAULT_XLSX))
    ap.add_argument("--limit", type=int, default=None,
                    help="Only use this many Excel rows (1 for a first test).")
    ap.add_argument("--auto", action="store_true",
                    help="No prompts: do not wait for ENTER.")
    ap.add_argument("--headless", action="store_true",
                    help="Run the browser without a window (use with --auto).")
    ap.add_argument("--walk", action="store_true",
                    help="After filling, press Next and continue with the other "
                         "selected images (default: only the image that opens).")
    ap.add_argument("--steps", action="store_true",
                    help="First run the open_index_selected.py steps (200 per "
                         "page, last page, select all, Index Selected) and fill "
                         "the image that opens. Slower; off by default.")
    ap.add_argument("--only-image", default=None, metavar="NAME",
                    help="Fill ONLY the Excel row with this Image Name (any "
                         "Status); every other image is left alone.")
    ap.add_argument("--only-vendor-property", action="store_true",
                    help="Do not enter the invoice number.")
    ap.add_argument("--delete-errors", action="store_true",
                    help="After an invoice FAILS and its PDF is safely moved to "
                         "error_invoices/, delete it from the portal. Off by "
                         "default. Test on one invoice with --limit 1 first.")
    ap.add_argument("--dry-run", action="store_true",
                    help="List the rows that would be filled. No browser.")
    args = ap.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"STOPPED: no spreadsheet at {args.xlsx}")
    workbook = load_workbook(args.xlsx)
    if "Invoices" not in workbook.sheetnames:
        sys.exit("STOPPED: the workbook has no 'Invoices' sheet. Run "
                 "1_invoice_reader_pdf_to_excle.py (task 1) first.")
    sheet = workbook["Invoices"]
    columns = column_map(sheet)
    jobs = rows_to_process(sheet, columns, args.only_image)
    if args.only_vendor_property:
        for _, rec in jobs:
            rec["invoice_number"] = ""
    if not args.only_image:
        jobs, sheet_dups = split_sheet_duplicates(sheet, columns, jobs)
        for r, rec, note in sheet_dups:
            print(f"  row {r}: {rec['image_name']} -> DUPLICATE ({note})")
            if not args.dry_run:
                move_invoice(rec, DUP_DIR, note, print)
                set_cell(sheet, columns, r, "status", "Status", "SKIPPED")
                set_cell(sheet, columns, r, "recorder_note", "Recorder Note", note)
        if sheet_dups and not args.dry_run:
            save_workbook(workbook, args.xlsx)
    if args.limit is not None:
        jobs = jobs[:args.limit]

    print(f"{len(jobs)} row(s) ready to fill (Status PENDING/APPROVED).")
    if args.dry_run or not jobs:
        if not jobs and not args.dry_run:
            print("No browser opened because no row is ready. Rows marked "
                  "NEEDS_REVIEW are skipped while Status is PENDING. Check the "
                  "vendor and property in Excel, then set Status to APPROVED.")
        for r, rec in jobs:
            gap = missing_fields(rec)
            print(f"  row {r}: {rec['image_name']} | {rec['vendor']} | "
                  f"{rec['property']} | #{rec['invoice_number']}"
                  + (f"  MISSING: {gap}" if gap else ""))
        return

    url = os.getenv("BOTTOMLINE_URL")
    if not url:
        sys.exit("STOPPED: set BOTTOMLINE_URL=<portal address> in .env")

    from playwright.sync_api import sync_playwright
    selectors = json.loads(SELECTORS_FILE.read_text(encoding="utf-8"))

    def log(msg):
        print(f"  [portal] {msg}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=args.headless)
        context = (browser.new_context(storage_state=str(SESSION_FILE))
                   if SESSION_FILE.exists() else browser.new_context())
        page = context.new_page()
        if not login(page, url, log, auto=args.auto):
            browser.close()
            sys.exit(1)
        try:
            context.storage_state(path=str(SESSION_FILE))
        except Exception:
            pass

        # The Home page draws its menu a few seconds after login; wait for it.
        try:
            page.locator(selectors["queue"]["menu_opener"] +
                         f':has-text("{selectors["queue"]["menu_opener_text"]}")'
                         ).first.wait_for(state="visible", timeout=45000)
        except Exception:
            browser.close()
            sys.exit("STOPPED: the portal menu never appeared after login.")
        if not return_to_queue(page, selectors, log):
            browser.close()
            sys.exit("STOPPED: could not open 'Images to be Indexed'. "
                     "Nothing was typed.")

        remaining = dict(jobs)                   # row -> Excel record
        seen = set()
        # Optional (--steps): the open_index_selected.py steps, then fill the
        # image that opens. Skipped by default -- it only wastes time when the
        # Excel images are looked up by name below.
        if args.steps and not run_index_steps(page, selectors, log):
            browser.close()
            sys.exit("STOPPED: could not open the 'Index Selected' form. "
                     "Nothing was typed.")
        while args.steps and remaining:
            shown = shown_image_name(page, selectors)
            if not shown or shown in seen:
                log("No new image on screen; stopping.")
                break
            seen.add(shown)
            row = next((r for r, rec in remaining.items()
                        if same_image(rec["image_name"], shown)), None)
            print(f"\n=== portal image: {shown} ===")
            if row is None:
                print("  no ready Excel row has this image name -- not filled.")
            else:
                rec = remaining.pop(row)
                gap = missing_fields(rec)
                note = (f"missing in sheet: {gap}" if gap
                        else fill_one(page, selectors, rec, log))
                if note:
                    print(f"  SKIPPED: {note}")
                    set_cell(sheet, columns, row, "status", "Status", "SKIPPED")
                    set_cell(sheet, columns, row, "recorder_note",
                             "Recorder Note", note)
                else:
                    print("  FILLED (not saved). Review it in the browser.")
                    set_cell(sheet, columns, row, "status", "Status", "FILLED")
                    set_cell(sheet, columns, row, "recorder_note",
                             "Recorder Note", "")
                finish_row(page, selectors, rec, note, args.delete_errors, log)
                save_workbook(workbook, args.xlsx)
                if not args.auto and not note:
                    input("  Check the form. ENTER to continue... ")
            if not args.walk or not remaining or not click_next(page):
                break
            page.wait_for_timeout(2500)

        # Any Excel row not yet handled: find that image BY NAME and open it.
        # Other images are never opened or changed.
        if remaining:
            if args.steps:
                return_to_queue(page, selectors, log)
            set_page_size(page, selectors, log)
            queue_url = page.url
            queue_index = read_queue_fast(page, selectors, log)   # read ONCE
            for row, rec in list(remaining.items()):
                name = rec["image_name"]
                print(f"\n=== row {row}: {name} ===")
                found = next(((k, n) for k, n in queue_index.items()
                              if same_image(k, name)), None)
                if found is None:
                    note = "not found in the portal queue"
                elif not open_invoice(page, selectors, found[0], found[1], log,
                                      queue_url=queue_url):
                    note = "could not open/verify the invoice"
                else:
                    note = fill_one(page, selectors, rec, log)
                if note:
                    print(f"  SKIPPED: {note}")
                    set_cell(sheet, columns, row, "status", "Status", "SKIPPED")
                    set_cell(sheet, columns, row, "recorder_note",
                             "Recorder Note", note)
                else:
                    print("  FILLED (not saved). Review it in the browser.")
                    set_cell(sheet, columns, row, "status", "Status", "FILLED")
                    set_cell(sheet, columns, row, "recorder_note",
                             "Recorder Note", "")
                finish_row(page, selectors, rec, note, args.delete_errors, log)
                save_workbook(workbook, args.xlsx)
                if not args.auto and not note:
                    input("  Check the form. ENTER to continue... ")
                if row != list(remaining)[-1] and not (note and "queue" in note):
                    return_to_queue(page, selectors, log, queue_url)
                    set_page_size(page, selectors, log)

        if not args.auto:
            input("Browser stays open so you can check. ENTER to close... ")
        browser.close()
    print("\nDone.")


if __name__ == "__main__":
    main()
