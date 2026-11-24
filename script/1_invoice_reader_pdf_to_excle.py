"""TASK 1 of 2 -- folder of PDFs -> invoices.xlsx. Never touches the portal;
TASK 2 (2_invoice_recorder_fill_web.py) types the spreadsheet into Nexus.

Reads invoice PDFs with LOCAL Python only (script/reader_local.py):
PyMuPDF text layer + layout-aware label/value rules, and RapidOCR (offline,
ONNX models bundled in the pip package) for scanned pages / logo-only vendor
names. No Ollama, no Qwen, no network, no other machine.

Writes:
  * "Invoices" sheet  -- one row per NEW PDF, same columns as before
                         (plus an "Extraction Note" column). Rows that are
                         already there are never changed.
  * "LineItems" sheet -- every line item (and statement lines / extra invoices
                         found in the same PDF), keyed by Image Name + Invoice #.
  * "InvoiceDetails"  -- PO, subtotal, tax, freight, discount, bill-to,
                         ship-to, notes and doubts per PDF.
  * conversion_report_<yyyymmdd>.xlsx in the project root (one row per PDF).

Non-invoices (statements, past-due letters, notices, check requests,
receipt bundles) are written with Status = NOT_INVOICE so TASK 2 skips them.

Run:
    C:\\res1_venv\\Scripts\\python.exe script\\1_invoice_reader_pdf_to_excle.py

Preview only (read every PDF, print results, write nothing):  add --dry-run
"""
import argparse
import base64
import hashlib
import sys
import os
import json
import time
import re
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

from reader_local import extract_invoice_local, VendorIndex
from utility_lookup import UtilityLookup
from reader_directory import load_known_vendors, match_vendor, _normalise
from vendor_lookup import load_master, load_aliases, load_alias_examples, Lookup
from master_index import build_index, extra_for_confident_match
from invoice_schema import NexusExtractedInvoice


HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent

# Spreadsheet to append to (lives in the project root: X:\res1_automation)
DEFAULT_XLSX = PROJECT / "invoices.xlsx"

# Invoice / bill PDFs folder
DEFAULT_PDF_DIR = PROJECT / "invoice"

MASTERS_DIR = PROJECT / "masters"
VENDORS_CSV = MASTERS_DIR / "vendors_data.csv"
ALIASES_CSV = MASTERS_DIR / "aliases.csv"                  # optional
PROPERTIES_CSV = MASTERS_DIR / "properties_data.csv"
PROPERTY_ALIASES_CSV = MASTERS_DIR / "property_aliases.csv"  # optional


def _key(header):
    """'Invoice Due Date' -> 'invoice_due_date'. Mirrors invoice_recorder's
    header_to_key so the two agree on column names."""
    cleaned = "".join(c if (c.isalnum() or c.isspace()) else " "
                      for c in str(header))
    return "_".join(cleaned.lower().split())


def to_base64(png_bytes_list):
    return [base64.b64encode(b).decode("ascii") for b in png_bytes_list]


# Date formats we accept from an invoice, tried in order. US month-first is
# tried before day-first so an ambiguous "07/08/2026" reads as July 8 (the US
# convention Bottomline expects), not August 7.
_DATE_FORMATS = ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%Y-%m-%d", "%Y/%m/%d",
                 "%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y", "%d/%m/%Y")


def _to_mmddyyyy(value):
    """Return the date as MM/DD/YYYY. Blank stays blank. A value we cannot
    parse is returned unchanged -- we never invent or reshuffle a date we do
    not understand; it stays visible for a human to fix."""
    if value in (None, ""):
        return value
    text = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    # Try month-year formats e.g. "09/2026", "September 2026" -> 09/01/2026
    for fmt in ("%m/%Y", "%m-%Y", "%B %Y", "%b %Y", "%Y-%m", "%Y/%m"):
        try:
            return datetime.strptime(text, fmt).strftime("%m/01/%Y")
        except ValueError:
            pass
    return text


def parse_date_to_yyyymmdd(dt_val):
    """Robustly parse date strings like '09/2026', 'September 2026', '09/09/2026', etc. to YYYYMMDD."""
    if not dt_val:
        return ""
    text = str(dt_val).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d")
        except ValueError:
            pass
    # Try month-year formats e.g. "09/2026", "September 2026"
    for fmt in ("%m/%Y", "%m-%Y", "%B %Y", "%b %Y", "%Y-%m", "%Y/%m"):
        try:
            d = datetime.strptime(text, fmt)
            # Default to end of month or 1st of month
            return d.strftime("%Y%m01")
        except ValueError:
            pass
    # Regex fallback: 8 continuous digits (YYYYMMDD or MMDDYYYY)
    digits = re.sub(r"\D", "", text)
    if len(digits) == 8:
        if int(digits[:4]) > 1900:  # YYYYMMDD
            return digits
        else:  # MMDDYYYY -> YYYYMMDD
            return f"{digits[4:]}{digits[:4]}"
    return ""


def last_data_row(sheet, columns):
    """The highest row that actually holds invoice data.

    openpyxl's `sheet.max_row` counts empty-but-formatted cells, so on a sheet
    with styling/validation reaching far down it reports a row hundreds below
    the real data -- which made the first live run append into that gap. This
    scans the key columns instead and returns the last row with a real value
    (or the header row, 1, when the sheet is empty)."""
    key_cols = [columns[k] for k in ("image_name", "vendor", "invoice_number")
                if k in columns]
    if not key_cols:                       # no recognised headers: use them all
        key_cols = [c.column for c in sheet[1] if c.value]
    last = 1                               # header row
    for row in sheet.iter_rows(min_row=2):
        if any(row[c - 1].value not in (None, "") for c in key_cols):
            last = row[0].row
    return last


def existing_image_names(sheet):
    """Normalised image names already in the sheet (the duplicate guard)."""
    col = None
    for cell in sheet[1]:
        if cell.value and _key(cell.value) == "image_name":
            col = cell.column
            break
    if col is None:
        return set()
    names = set()
    for row in sheet.iter_rows(min_row=2):
        value = row[col - 1].value
        if value not in (None, ""):
            names.add(_normalise(value))
    return names


def append_invoice(sheet, columns, fields, image_name, row,
                   raw_vendor=None, qwen_status=None,
                   raw_property=None, property_status=None,
                   integration_package=None, status="PENDING",
                   extraction_note=None):
    """Write ONE row at `row` (the caller finds it with last_data_row()+1, NOT
    sheet.max_row, which is unreliable). invoice_number is stored as TEXT so
    leading zeros survive; dates are forced to MM/DD/YYYY.

    raw_vendor / qwen_status are the vendor_lookup helper columns
    ("Qwen Vendor (raw)" / "Qwen Status"). raw_property / property_status are the
    matching property helper columns ("Qwen Property (raw)" / "Property Status"),
    and integration_package is the property's package (YARDI/INTACCT) when a
    confident property match gave us one -- a HINT for the recorder, blank on
    review. Every one of these is written ONLY when the sheet actually has that
    column (the skip rule in the write loop)."""
    r = row
    to_write = {
        "image_name": image_name,
        "vendor": fields.get("vendor"),
        "vendor_code": fields.get("vendor_code"),
        "property": fields.get("property"),
        "property_code": fields.get("property_code"),
        "invoice_number": fields.get("invoice_number"),
        "invoice_date": _to_mmddyyyy(fields.get("invoice_date")),
        "invoice_due_date": _to_mmddyyyy(fields.get("due_date")),
        "amount": fields.get("amount"),
        "cycle_from": _to_mmddyyyy(fields.get("service_period_start")),
        "cycle_to": _to_mmddyyyy(fields.get("service_period_end")),
        "late_fee": fields.get("late_fee"),
        "gl_code": fields.get("gl_code"),
        "account_number": fields.get("account_number"),
        "service_address": fields.get("service_address"),
        "vendor_address": fields.get("vendor_address"),
        "bill_type": fields.get("bill_type"),
        "status": status,
        "integration_package": integration_package,
        "qwen_vendor_raw": raw_vendor,
        "qwen_status": qwen_status,
        "qwen_property_raw": raw_property,
        "property_status": property_status,
        "extraction_note": extraction_note,
    }
    for key, value in to_write.items():
        if value in (None, ""):
            continue
        if key not in columns:
            # Add missing column header dynamically so data is never discarded
            new_col = sheet.max_column + 1
            header_name = " ".join(word.capitalize() for word in key.split("_"))
            sheet.cell(row=1, column=new_col, value=header_name)
            columns[key] = new_col
        cell = sheet.cell(row=r, column=columns[key], value=str(value))
        if key in ("invoice_number", "account_number"):
            cell.number_format = "@"   # keep leading zeros as text


# Whole-word utility keywords. Word boundaries stop "Powers Plumbing" or
# "Gas Station Supply" style vendors from being treated as utility bills.
_UTILITY_RE = re.compile(
    r"\b(pepco|bge|wssc|dominion|comcast|verizon|at&t|"
    r"gas|water|sewer|electric|electricity|power|telecom|internet)\b",
    re.IGNORECASE)


_UTILITY_BILL_TYPES = {
    "electric": "Utility (Electric)",
    "gas": "Utility (Gas)",
    "water": "Utility (Water/Sewer)",
    "sewer": "Utility (Water/Sewer)",
    "trash": "Utility (Trash/Waste)",
    "phone/internet": "Telecom / Internet",
}


def bill_type_for(fields, is_util_bill):
    """The 'Bill Type' shown in the Excel. Utility bills take it from the
    utility type Qwen read; everything else from Qwen's doc_type."""
    if is_util_bill:
        u = str(fields.get("utility_type") or "").lower()
        for word, label in _UTILITY_BILL_TYPES.items():
            if word in u:
                return label
        return "Utility (Other)"
    return fields.get("doc_type") or "Standard Vendor Invoice"


def _looks_like_utility(vendor_name):
    return bool(vendor_name and _UTILITY_RE.search(str(vendor_name)))


REQUIRED_HEADERS = ["Image Name", "Vendor", "Vendor Code", "Property",
                    "Property Code", "Invoice Number", "Invoice Date",
                    "Invoice Due Date", "Amount", "Cycle From", "Cycle To",
                    "Status"]


def get_invoice_sheet(workbook):
    """The 'Invoices' sheet. If the workbook has none (e.g. it only holds
    other sheets like 'Invoice Queue'), create a NEW 'Invoices' sheet with the
    expected headers instead of writing into some unrelated first sheet."""
    if "Invoices" in workbook.sheetnames:
        return workbook["Invoices"]
    sheet = workbook.create_sheet("Invoices")
    for col, header in enumerate(REQUIRED_HEADERS, start=1):
        sheet.cell(row=1, column=col, value=header)
    print("  [No 'Invoices' sheet found -- created one with the standard "
          "headers; your other sheets are untouched.]")
    return sheet


def file_hash(path):
    digest = hashlib.md5()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def save_workbook(workbook, path):
    """Save, but never crash the whole run if the file is open in Excel."""
    try:
        workbook.save(path)
        return True
    except PermissionError:
        print(f"  !! Could not save {path} -- is it open in Excel? Close it. "
              "Will retry on the next row.")
        return False


def _columns(sheet):
    return {_key(c.value): c.column for c in sheet[1] if c.value}


def main():
    load_dotenv(PROJECT / ".env")
    load_dotenv(HERE / ".env")
    parser = argparse.ArgumentParser(
        description="Read invoice PDFs locally (PyMuPDF + RapidOCR); fill invoices.xlsx.")
    parser.add_argument("--pdf-dir", type=Path, default=Path(os.getenv("PDF_DIR") or DEFAULT_PDF_DIR),
                        help="Folder of invoice PDFs. Default: <project>\\invoice.")
    parser.add_argument("--xlsx", type=Path, default=Path(os.getenv("XLSX_PATH") or DEFAULT_XLSX),
                        help="invoices.xlsx to append to. Default: <project>\\invoices.xlsx.")
    parser.add_argument("--max-pages", type=int, default=None,
                        help="Only read the first N pages of each PDF (default: all pages).")
    parser.add_argument("--report", type=Path, default=None,
                        help="Conversion report path. Default: <project>\\conversion_report_<yyyymmdd>.xlsx")
    parser.add_argument("--no-report", action="store_true",
                        help="Do not write the conversion report.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Only read this many PDFs. Use for a first test.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Read every PDF and print what was found. Nothing is written.")
    parser.add_argument("--utility", action="store_true",
                        help="Process files as Utility Bills (Pepco, Gas, Water, etc.).")
    parser.add_argument("--workers", "--worker", dest="workers", type=int, default=1,
                        help="Kept for compatibility; local reading runs one PDF at a time "
                             "(about 0.1 s per text PDF, a few seconds per scanned page).")
    args = parser.parse_args()

    if not args.pdf_dir.exists():
        sys.exit(f"STOPPED: no PDF folder at {args.pdf_dir}")
    if not args.xlsx.exists():
        sys.exit(f"STOPPED: no spreadsheet at {args.xlsx}")
    workbook = load_workbook(args.xlsx)
    sheet = get_invoice_sheet(workbook)
    columns = _columns(sheet)
    seen = existing_image_names(sheet)

    all_pdfs = sorted(p for p in args.pdf_dir.glob("*.pdf"))
    if not all_pdfs:
        sys.exit(f"STOPPED: no PDFs found in {args.pdf_dir}")

    # Skip byte-identical copies ("- Copy", "(1)") so one invoice is not read
    # and recorded twice under different file names.
    hashes, unique_pdfs = {}, []
    dup_content = {}
    # Originals first: names containing "copy" or "(1)" sort last, so the
    # clean file name is the one kept.
    for p in sorted(all_pdfs, key=lambda x: (
            "copy" in x.stem.lower() or "(" in x.stem, x.name)):
        h = file_hash(p)
        if h in hashes:
            print(f"  [Duplicate content skipped: {p.name} == {hashes[h]}]")
            dup_content[p.name] = hashes[h]
            continue
        hashes[h] = p.name
        unique_pdfs.append(p)
    all_pdfs = unique_pdfs

    # Prioritize unseen PDFs so --limit tests actually process new files
    unseen_pdfs = [p for p in all_pdfs if _normalise(p.stem) not in seen]
    if unseen_pdfs:
        pdfs = unseen_pdfs[:args.limit] if args.limit is not None else unseen_pdfs
    else:
        pdfs = all_pdfs[:args.limit] if args.limit is not None else all_pdfs

    master_rows = load_master(VENDORS_CSV, name_col="vendor", address_col="vendorAddress")
    aliases = load_aliases(ALIASES_CSV)
    prop_vendor_rows = load_master(MASTERS_DIR / "property_vendor_map.csv")
    alias_examples = load_alias_examples(ALIASES_CSV, limit=5)
    lookup = None
    known = None
    if master_rows:
        lookup = Lookup(master_rows, aliases, name_col="vendor", address_col="vendorAddress", code_col="vendorCode", property_vendor_rows=prop_vendor_rows)
    else:
        print(f"\n*** WARNING: could not read the vendor master at "
              f"{VENDORS_CSV} (missing or empty). Falling back to the OLD "
              "vendor matching (known vendors already in this spreadsheet) "
              "so this run does not crash. Ask Puneet/Manish to check the "
              "masters folder. ***\n")
        known = load_known_vendors(args.xlsx)

    property_rows = load_master(PROPERTIES_CSV, name_col="Property Name")
    property_lookup = None
    package_index = {}
    if property_rows:
        property_aliases = load_aliases(PROPERTY_ALIASES_CSV)
        property_lookup = Lookup(property_rows, property_aliases,
                                 name_col="Property Name",
                                 code_col="Property Code")
        package_index = build_index(property_rows, "Property Name", "Integration Package")
    else:
        print(f"\n*** NOTE: no property master at {PROPERTIES_CSV}; property "
              "is written RAW for review and Integration Package left blank "
              "(today's behaviour). ***\n")

    utility_lookup = None
    try:
        utility_lookup = UtilityLookup(MASTERS_DIR)
        print("  [Utility crosswalk & address master loaded]")
    except Exception as ue:
        print(f"  [Notice: utility crosswalk not loaded: {ue}]")

    # Load Query Resolution Sheet rules
    query_resolutions = {}
    query_res_file = MASTERS_DIR / "query_resolutions.json"
    if query_res_file.exists():
        try:
            with open(query_res_file, "r", encoding="utf-8") as qf:
                query_resolutions = json.load(qf)
            print(f"  [Query Resolutions Loaded: {len(query_resolutions.get('property_redirects', {}))} prop redirects, {len(query_resolutions.get('vendor_codes', {}))} vendor codes]")
        except Exception as qe:
            print(f"  Notice: could not load query resolutions ({qe})")

    return _run_local(args, workbook, sheet, columns, seen, all_pdfs, pdfs, master_rows, lookup, known,
                      property_lookup, package_index, utility_lookup, query_resolutions, dup_content)


# --------------------------------------------------------------------------
# Local (no-Ollama) run: read -> post-process -> write rows / sheets / report
# --------------------------------------------------------------------------
_NON_INVOICE = {"statement", "past_due_notice", "notice", "check_request", "receipt", "unreadable"}
_HARD_DOUBT = re.compile(r"^(missing|total .* differs|line items add up|scanned|pdf holds|unreadable)", re.I)
LINE_ITEM_HEADERS = ["Image Name", "Invoice Number", "Vendor", "Source", "Line #", "Page",
                     "Item Code", "Description", "Qty", "Unit Price", "Amount"]
DETAIL_HEADERS = ["Image Name", "Invoice Number", "Doc Type", "Vendor (printed)", "Vendor Master Match",
                  "Vendor Evidence", "Invoice Date", "Due Date", "PO Number", "Account Number",
                  "Subtotal", "Tax", "Freight", "Discount", "Total", "Bill To", "Ship To",
                  "Service Address", "Vendor Address", "Pages", "OCR Used", "Other Invoices In PDF",
                  "Notes", "Doubts"]
REPORT_HEADERS = ["File Name", "Status", "Doc Type", "Vendor", "Invoice #", "Invoice Date", "Due Date",
                  "Total", "Line Items", "Property", "Missing / Doubtful Fields (reason)", "Notes",
                  "invoices.xlsx Action", "Existing Row Check", "OCR Used", "Seconds"]


def _num(v):
    try:
        return round(float(str(v).replace(",", "").replace("$", "").strip()), 2)
    except (TypeError, ValueError):
        return None


def _inv_key(v):
    return re.sub(r"[^0-9a-z]", "", str(v or "").lower()).lstrip("0")


def _vendor_close(a, b):
    a = re.sub(r"[^a-z]", "", str(a or "").lower())
    b = re.sub(r"[^a-z]", "", str(b or "").lower())
    return not a or not b or a[:4] == b[:4]


def _existing_rows(sheet, columns):
    """Rows already in the Invoices sheet: by image name, and by
    (invoice #, amount) for the vendor+invoice+amount duplicate guard."""
    by_image, keys = {}, {}
    for row in sheet.iter_rows(min_row=2):
        def g(k):
            c = columns.get(k)
            return row[c - 1].value if c else None
        img = g("image_name")
        if img in (None, ""):
            continue
        info = {"row": row[0].row, "vendor": g("vendor"), "invoice_number": g("invoice_number"),
                "amount": g("amount"), "invoice_date": g("invoice_date"), "status": g("status")}
        by_image[_normalise(img)] = info
        inv, amt = _inv_key(info["invoice_number"]), _num(info["amount"])
        if inv and amt is not None:
            keys.setdefault((inv, amt), []).append((info["row"], str(info["vendor"] or ""), str(img)))
    return by_image, keys


def _compare_existing(ex, fields):
    """Text describing where an existing (older) row disagrees with this read."""
    diffs = []
    if _inv_key(ex.get("invoice_number")) != _inv_key(fields.get("invoice_number")):
        diffs.append(f"invoice # '{ex.get('invoice_number')}' vs read '{fields.get('invoice_number')}'")
    if _num(ex.get("amount")) != _num(fields.get("amount")):
        diffs.append(f"amount '{ex.get('amount')}' vs read '{fields.get('amount')}'")
    if str(_to_mmddyyyy(ex.get("invoice_date")) or "") != str(_to_mmddyyyy(fields.get("invoice_date")) or ""):
        diffs.append(f"date '{ex.get('invoice_date')}' vs read '{fields.get('invoice_date')}'")
    if not _vendor_close(ex.get("vendor"), fields.get("vendor")):
        diffs.append(f"vendor '{ex.get('vendor')}' vs read '{fields.get('vendor')}'")
    return ("existing row differs: " + "; ".join(diffs)) if diffs else "existing row matches"


def _note_for(fields):
    parts = []
    if fields.get("doc_class") in _NON_INVOICE:
        parts.append(f"NOT AN INVOICE ({fields.get('doc_type') or fields.get('doc_class')})")
    parts += [d for d in fields.get("doubts") or []]
    parts += [n for n in fields.get("notes") or [] if n not in parts]
    text = " | ".join(parts)
    return text[:900] if text else None


def _report_status(fields, err):
    if err or fields is None:
        return "failed"
    if fields.get("doc_class") in _NON_INVOICE:
        return "not-an-invoice"
    if any(_HARD_DOUBT.match(d) for d in fields.get("doubts") or []):
        return "partial"
    return "success"


def _rewrite_keyed_sheet(workbook, title, headers, new_rows):
    """(Re)write a helper sheet; rows of images read in THIS run replace that
    image's earlier rows, rows of other images are kept."""
    new_images = {r.get("Image Name") for r in new_rows}
    kept = []
    if title in workbook.sheetnames:
        ws = workbook[title]
        old_headers = [c.value for c in ws[1]]
        for vals in ws.iter_rows(min_row=2, values_only=True):
            if not vals or all(v in (None, "") for v in vals):
                continue
            rec = dict(zip(old_headers, vals))
            if rec.get("Image Name") not in new_images:
                kept.append(rec)
        workbook.remove(ws)
    ws = workbook.create_sheet(title)
    ws.append(headers)
    for rec in kept + list(new_rows):
        ws.append([rec.get(h) for h in headers])
    for col in ("B",):
        for cell in ws[col][1:]:
            cell.number_format = "@"
    ws.freeze_panes = "A2"
    return ws


def _write_report(path, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append(REPORT_HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
    fills = {"success": "C6EFCE", "partial": "FFEB9C", "failed": "FFC7CE",
             "not-an-invoice": "D9D9D9", "duplicate": "DDEBF7"}
    for r in rows:
        ws.append([r.get(h) for h in REPORT_HEADERS])
        st = r.get("Status")
        if st in fills:
            ws.cell(row=ws.max_row, column=2).fill = PatternFill("solid", fgColor=fills[st])
    widths = [48, 14, 22, 34, 18, 12, 12, 12, 10, 30, 70, 70, 40, 60, 9, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet("Summary")
    counts = {}
    for r in rows:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1
    s.append(["Status", "PDFs"])
    for k in ("success", "partial", "failed", "not-an-invoice", "duplicate"):
        s.append([k, counts.get(k, 0)])
    s.append(["TOTAL", len(rows)])
    s.append([])
    s.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M IST")])
    s.append(["Reader", "reader_local.py (PyMuPDF text + RapidOCR, no Ollama)"])
    wb.save(path)


def _run_local(args, workbook, sheet, columns, seen, all_pdfs, pdfs, master_rows, lookup, known,
               property_lookup, package_index, utility_lookup, query_resolutions, dup_content):
    vendor_index = VendorIndex(master_rows) if master_rows else None
    to_read = all_pdfs if args.limit is None else pdfs
    by_image, dup_keys = _existing_rows(sheet, columns)

    def process_single_pdf(pdf):
        """Read one PDF locally, then the same post-processing as before
        (vendor/property masters, utility rules, query resolutions)."""
        started = time.perf_counter()
        try:
            fields = extract_invoice_local(pdf, vendor_index=vendor_index, max_pages=args.max_pages)
        except Exception as e:
            print(f"  [Error reading {pdf.name}]: {type(e).__name__}: {e} -- marked for REVIEW")
            return pdf, None, None, None, None, None, time.perf_counter() - started, f"read_error: {type(e).__name__}: {e}"
        elapsed = time.perf_counter() - started
        if not fields or fields.get("doc_class") == "unreadable":
            return pdf, fields, None, None, None, None, elapsed, "unreadable: no text even after OCR"
        fields["vendor_printed"] = fields.get("vendor")

        # Remove whitespace to maintain continuity
        if fields.get("invoice_number"):
            fields["invoice_number"] = "".join(str(fields["invoice_number"]).split())
        if fields.get("account_number"):
            fields["account_number"] = "".join(str(fields["account_number"]).split())

        # Phase 1: Pydantic Validation & Currency/Date Sanitization
        try:
            validated = NexusExtractedInvoice(
                vendor_name=fields.get("vendor") or "",
                invoice_number=fields.get("invoice_number") or "",
                invoice_date=fields.get("invoice_date"),
                due_date=fields.get("due_date"),
                amount=fields.get("amount") or "0.00",
                property_name=fields.get("property"),
                service_address=fields.get("service_address"),
                vendor_address=fields.get("vendor_address"),
                account_number=fields.get("account_number")
            )
            fields["amount"] = validated.amount
            fields["invoice_number"] = validated.invoice_number
            if validated.vendor_address and not fields.get("vendor_address"):
                fields["vendor_address"] = validated.vendor_address
            if validated.service_address and not fields.get("service_address"):
                fields["service_address"] = validated.service_address
            if validated.account_number and not fields.get("account_number"):
                fields["account_number"] = validated.account_number
        except Exception as pe:
            print(f"    (Notice Pydantic validation: {pe})")

        raw_vendor = fields.get("vendor")
        if lookup is not None:
            m = lookup.match(raw_vendor)
            how = str(fields.get("vendor_how") or "")
            if m.status != "READ_OK" and fields.get("vendor_master_match") and \
                    "printed company line" not in how and \
                    any(k in how for k in ("name", "filename", "domain", "logo")):
                # the local reader confirmed this master row from name / remit
                # address / zip / web domain -- let the lookup resolve it
                full = fields["vendor_master_match"]
                for cand in (full, re.sub(r"\s*\(.*?\)\s*$", "", full)):
                    m2 = lookup.match(cand)
                    if m2.status == "READ_OK":
                        m = m2
                        break
            fields["vendor"] = m.value
            fields["vendor_code"] = m.code
            qwen_status = m.status
        else:
            canonical = match_vendor(raw_vendor, known)
            if canonical:
                fields["vendor"] = canonical
                qwen_status = "READ_OK"
            else:
                qwen_status = "NEEDS_REVIEW"

        raw_property = fields.get("property")
        property_status = None
        package = None

        v_check = str(raw_vendor or fields.get("vendor") or "").lower()
        # Utility handling only when the CONTENT is a utility bill (kWh /
        # therms / meter readings), not merely because the vendor name has
        # "Power" or "Water" in it (that wrongly turned Fidelity Power
        # Systems letters into "N/A-20261001" rows before).
        is_util_bill = args.utility or bool(fields.get("is_utility_bill"))
        fields["bill_type"] = bill_type_for(fields, is_util_bill)

        if is_util_bill:
            raw_acct = str(fields.get("account_number") or "").strip()
            if not raw_acct:
                raw_acct = str(fields.get("invoice_number") or "").strip()
            acct = "".join(raw_acct.split())
            inv_dt = fields.get("invoice_date")
            date_suffix = parse_date_to_yyyymmdd(inv_dt)

            if acct and date_suffix:
                if acct.endswith(f"-{date_suffix}"):
                    fields["invoice_number"] = acct
                else:
                    fields["invoice_number"] = f"{acct}-{date_suffix}"
            elif acct:
                fields["invoice_number"] = acct

            if acct:
                fields["account_number"] = acct

            raw_curr = str(fields.get("current_charges") or fields.get("amount") or "").strip()
            clean_val = raw_curr.replace("$", "").replace(",", "").strip()
            is_negative = False
            if clean_val.startswith("-") or clean_val.endswith("-") or (clean_val.startswith("(") and clean_val.endswith(")")) or "cr" in clean_val.lower():
                is_negative = True
            else:
                try:
                    num_val = float(re.sub(r"[^\d.-]", "", clean_val))
                    if num_val < 0:
                        is_negative = True
                except ValueError:
                    pass

            if is_negative:
                fields["amount"] = "0"
            else:
                try:
                    num_val = float(re.sub(r"[^\d.]", "", clean_val)) if clean_val else 0.0
                    fields["amount"] = f"{num_val:.2f}"
                except ValueError:
                    fields["amount"] = clean_val

            srv_addr = fields.get("service_address") or raw_property
            raw_property = srv_addr
            if utility_lookup is not None and srv_addr:
                um = utility_lookup.match(srv_addr, utility_type=fields.get("utility_type"))
                fields["property"] = um.property_name
                fields["property_code"] = um.property_code
                package = um.integration_package
                property_status = um.status
                if um.gl_code:
                    fields["gl_code"] = um.gl_code

            if not fields.get("gl_code"):
                u_type = str(fields.get("utility_type") or "").lower()
                if "elec" in u_type or "pepco" in v_check:
                    fields["gl_code"] = "436313"
                elif "gas" in u_type or "gas" in v_check:
                    fields["gl_code"] = "436323"
                elif "water" in u_type or "sewer" in u_type or "wssc" in v_check:
                    fields["gl_code"] = "436311"
        else:
            if property_lookup is not None and raw_property:
                pm = property_lookup.match(raw_property)
                fields["property"] = pm.value
                fields["property_code"] = pm.code
                property_status = pm.status
                package = extra_for_confident_match(pm, package_index) or None

            srv_addr = fields.get("service_address")
            if (not fields.get("property") or property_status != "READ_OK") and utility_lookup is not None and srv_addr:
                um = utility_lookup.match(srv_addr)
                if um.status == "READ_OK":
                    fields["property"] = um.property_name
                    fields["property_code"] = um.property_code
                    package = um.integration_package
                    property_status = um.status

        if not fields.get("vendor_address") and lookup is not None and fields.get("vendor"):
            v_norm = _normalise(fields["vendor"])
            for n_name, canon_name, n_addr, _ in lookup._candidates:
                if n_name == v_norm and n_addr:
                    fields["vendor_address"] = n_addr.title()
                    break

        if query_resolutions:
            v_codes = query_resolutions.get("vendor_codes", {})
            v_redirs = query_resolutions.get("vendor_redirects", {})
            curr_v = fields.get("vendor") or raw_vendor
            for k_vendor, forced_code in v_codes.items():
                if k_vendor.lower() in str(curr_v).lower():
                    fields["vendor_code"] = forced_code
                    break
            for k_vendor, target_v in v_redirs.items():
                if k_vendor.lower() in str(curr_v).lower():
                    fields["vendor"] = target_v
                    break

        if lookup is not None and fields.get("vendor"):
            refined_code = lookup.resolve_code_for_package(
                fields["vendor"],
                package=package,
                property_name=fields.get("property"),
                property_code=fields.get("property_code"),
                fallback_code=fields.get("vendor_code")
            )
            if refined_code:
                fields["vendor_code"] = refined_code

        return pdf, fields, raw_vendor, qwen_status, raw_property, property_status, elapsed, None

    review, written, skipped = [], 0, 0
    per_invoice_times = []
    batch_start = time.perf_counter()
    next_row = last_data_row(sheet, columns) + 1
    item_rows, detail_rows, report_rows = [], [], []

    for name, orig in dup_content.items():
        report_rows.append({"File Name": name, "Status": "duplicate",
                            "Missing / Doubtful Fields (reason)": f"byte-identical copy of {orig}",
                            "invoices.xlsx Action": "skipped (same file content as another PDF)"})

    print(f"\n--- Reading {len(to_read)} PDF(s) locally (PyMuPDF + RapidOCR, no Ollama) ---\n")
    for pdf in to_read:
        pdf, fields, raw_vendor, qwen_status, raw_property, property_status, elapsed, err = process_single_pdf(pdf)
        per_invoice_times.append(elapsed)
        rep = {"File Name": pdf.name, "Seconds": round(elapsed, 1)}
        if err or fields is None:
            review.append(pdf.name)
            rep.update({"Status": "failed", "Missing / Doubtful Fields (reason)": err,
                        "invoices.xlsx Action": "not written (read failed)"})
            report_rows.append(rep)
            print(f"  FAILED: {pdf.name} ({err}) [{elapsed:.1f}s]")
            continue

        doc_class = fields.get("doc_class")
        status_x = "NOT_INVOICE" if doc_class in _NON_INVOICE else "PENDING"
        note = _note_for(fields)
        rstatus = _report_status(fields, None)
        norm = _normalise(pdf.stem)
        existing_check = None
        if norm in seen:
            ex = by_image.get(norm)
            skipped += 1
            action = (f"already in invoices.xlsx (row {ex['row']}, Status {ex.get('status')}) -- row left unchanged"
                      if ex else "already in invoices.xlsx -- left unchanged")
            existing_check = _compare_existing(ex, fields) if ex else None
        else:
            dkey = (_inv_key(fields.get("invoice_number")), _num(fields.get("amount")))
            hits = []
            if dkey[0] and dkey[1] is not None:
                hits = [h for h in dup_keys.get(dkey, []) if _vendor_close(h[1], fields.get("vendor"))]
            if hits:
                rstatus = "duplicate"
                action = (f"not written: same vendor + invoice # + amount as row {hits[0][0]} "
                          f"({hits[0][2]})")
            elif args.dry_run:
                action = f"dry run -- would append (Status {status_x})"
            else:
                append_invoice(sheet, columns, fields, pdf.stem, next_row,
                               raw_vendor=raw_vendor, qwen_status=qwen_status,
                               raw_property=raw_property, property_status=property_status,
                               integration_package=fields.get("integration_package"),
                               status=status_x, extraction_note=note)
                action = f"appended row {next_row} (Status {status_x})"
                if dkey[0] and dkey[1] is not None:
                    dup_keys.setdefault(dkey, []).append((next_row, str(fields.get("vendor") or ""), pdf.stem))
                next_row += 1
                seen.add(norm)
                written += 1
                save_workbook(workbook, args.xlsx)

        inv_no = fields.get("invoice_number") or None
        vend = fields.get("vendor") or fields.get("vendor_printed")
        src = "statement line" if doc_class == "statement" else "line item"
        for i, it in enumerate(fields.get("line_items") or [], start=1):
            item_rows.append({"Image Name": pdf.stem, "Invoice Number": inv_no, "Vendor": vend, "Source": src,
                              "Line #": i, "Page": it.get("page"), "Item Code": it.get("item_code"),
                              "Description": it.get("description"), "Qty": it.get("qty"),
                              "Unit Price": it.get("unit_price"), "Amount": it.get("amount")})
        others = []
        if fields.get("multi_invoice"):
            for sd in fields.get("sub_documents") or []:
                others.append(f"#{sd.get('invoice_number')} {sd.get('invoice_date')} ${sd.get('amount')} (p.{sd.get('pages')})")
                for i, it in enumerate(sd.get("line_items") or [], start=1):
                    item_rows.append({"Image Name": pdf.stem, "Invoice Number": sd.get("invoice_number"),
                                      "Vendor": vend, "Source": "line item (other invoice in same PDF)",
                                      "Line #": i, "Page": it.get("page"), "Item Code": it.get("item_code"),
                                      "Description": it.get("description"), "Qty": it.get("qty"),
                                      "Unit Price": it.get("unit_price"), "Amount": it.get("amount")})
        detail_rows.append({
            "Image Name": pdf.stem, "Invoice Number": inv_no, "Doc Type": fields.get("doc_type"),
            "Vendor (printed)": fields.get("vendor_printed"), "Vendor Master Match": fields.get("vendor_master_match"),
            "Vendor Evidence": fields.get("vendor_how"), "Invoice Date": _to_mmddyyyy(fields.get("invoice_date")),
            "Due Date": _to_mmddyyyy(fields.get("due_date")), "PO Number": fields.get("po_number"),
            "Account Number": fields.get("account_number"), "Subtotal": fields.get("subtotal"),
            "Tax": fields.get("tax"), "Freight": fields.get("freight"), "Discount": fields.get("discount"),
            "Total": fields.get("amount"), "Bill To": fields.get("bill_to"), "Ship To": fields.get("ship_to"),
            "Service Address": fields.get("service_address"), "Vendor Address": fields.get("vendor_address"),
            "Pages": fields.get("pages"), "OCR Used": "yes" if fields.get("ocr_used") else "no",
            "Other Invoices In PDF": "; ".join(others) or None,
            "Notes": " | ".join(fields.get("notes") or []) or None,
            "Doubts": " | ".join(fields.get("doubts") or []) or None})
        rep.update({"Status": rstatus, "Doc Type": fields.get("doc_type"), "Vendor": vend,
                    "Invoice #": inv_no, "Invoice Date": _to_mmddyyyy(fields.get("invoice_date")),
                    "Due Date": _to_mmddyyyy(fields.get("due_date")), "Total": fields.get("amount"),
                    "Line Items": len(fields.get("line_items") or []), "Property": fields.get("property"),
                    "Missing / Doubtful Fields (reason)": "; ".join(fields.get("doubts") or []) or None,
                    "Notes": "; ".join(fields.get("notes") or []) or None,
                    "invoices.xlsx Action": action, "Existing Row Check": existing_check,
                    "OCR Used": "yes" if fields.get("ocr_used") else "no"})
        report_rows.append(rep)
        print(f"  {rstatus:14s} {pdf.name} -> {vend} | #{inv_no} | {fields.get('invoice_date')} | "
              f"{fields.get('amount')} | {action} [{elapsed:.1f}s]")

    if not args.dry_run:
        _rewrite_keyed_sheet(workbook, "LineItems", LINE_ITEM_HEADERS, item_rows)
        _rewrite_keyed_sheet(workbook, "InvoiceDetails", DETAIL_HEADERS, detail_rows)
        if "Invoices" in workbook.sheetnames:
            workbook.active = workbook.sheetnames.index("Invoices")
        save_workbook(workbook, args.xlsx)

    if not args.no_report:
        rpath = args.report or (PROJECT / f"conversion_report_{datetime.now():%Y%m%d}.xlsx")
        try:
            _write_report(rpath, report_rows)
            print(f"\nReport: {rpath}")
        except PermissionError:
            print(f"  !! Could not save report {rpath} -- is it open in Excel?")

    counts = {}
    for r in report_rows:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1
    batch_elapsed = time.perf_counter() - batch_start
    print(f"\nDone. {len(report_rows)} PDF(s): " + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    print(f"Rows appended to {args.xlsx.name}: {written}; already-recorded PDFs left unchanged: {skipped}; "
          f"line items written: {len(item_rows)}" + (" (DRY RUN: nothing written)" if args.dry_run else ""))
    if per_invoice_times:
        print(f"Time: {batch_elapsed:.1f}s total, {sum(per_invoice_times) / len(per_invoice_times):.2f}s average per PDF.")
    if review:
        print("Failed (needs a human): " + ", ".join(review))


if __name__ == "__main__":
    main()
