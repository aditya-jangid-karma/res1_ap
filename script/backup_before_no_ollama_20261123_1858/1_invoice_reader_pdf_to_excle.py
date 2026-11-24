"""TASK 1 of 2 -- folder of PDFs -> invoices.xlsx. Never touches the portal;
TASK 2 (invoice_recorder.py) types the spreadsheet into Nexus.

Read invoice PDFs with the local Qwen model and append their details to
invoices.xlsx. Fills only what an invoice actually shows. Property is filled
with the RAW location text Qwen reads (Ship-To preferred) for a human to
confirm; Integration Package is left blank for a human. Never touches Bottomline.

Run (the USER runs this -- it calls the live Qwen on the Mac):
    .venv\\Scripts\\python.exe script\\invoice_reader.py ^
        --host http://<YOUR-MAC>:11434

Offline preview (no Qwen, no writing):  add --dry-run
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

from reader_pdf import render_pdf
from reader_qwen import extract_invoice
from reader_qwen_utility import extract_utility_invoice
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
                   integration_package=None):
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
        "status": "PENDING",
        "integration_package": integration_package,
        "qwen_vendor_raw": raw_vendor,
        "qwen_status": qwen_status,
        "qwen_property_raw": raw_property,
        "property_status": property_status,
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
        description="Read invoice PDFs with local Qwen; fill invoices.xlsx.")
    parser.add_argument("--host", default=os.getenv("OLLAMA_HOST"),
                        help="Ollama address on the Mac, e.g. "
                             "http://192.168.1.42:11434. Required unless --dry-run.")
    parser.add_argument("--pdf-dir", type=Path, default=Path(os.getenv("PDF_DIR") or DEFAULT_PDF_DIR),
                        help="Folder of invoice PDFs. Default: <project>\\invoice.")
    parser.add_argument("--xlsx", type=Path, default=Path(os.getenv("XLSX_PATH") or DEFAULT_XLSX),
                        help="invoices.xlsx to append to. Default: <project>\\invoices.xlsx.")
    parser.add_argument("--dpi", type=int, default=140,
                        help="Render DPI for invoice pages. 140 balances speed and OCR clarity.")
    parser.add_argument("--max-pages", type=int, default=2,
                        help="Maximum pages sent to Qwen. Front pages contain all critical invoice details.")
    parser.add_argument("--poppler-path", default=None,
                        help="Poppler bin folder, if it is not on PATH.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Only read this many PDFs. Use for a first test.")
    parser.add_argument("--dry-run", action="store_true",
                        help="List the PDFs that would be read. No Qwen, no write.")
    parser.add_argument("--utility", action="store_true",
                        help="Process files as Utility Bills (Pepco, Gas, Water, etc.).")
    parser.add_argument("--workers", "--worker", dest="workers", type=int, default=4,
                        help="Number of concurrent reading workers (default: 4, recommended: 4 or 5).")
    args = parser.parse_args()

    if not args.pdf_dir.exists():
        sys.exit(f"STOPPED: no PDF folder at {args.pdf_dir}")
    if not args.xlsx.exists():
        sys.exit(f"STOPPED: no spreadsheet at {args.xlsx}")
    if not args.dry_run and not args.host:
        sys.exit("STOPPED: --host is required for a real run (or configure OLLAMA_HOST in .env).")
    if args.host:
        args.host = args.host.strip()
        if not args.host.startswith("http://") and not args.host.startswith("https://"):
            args.host = f"http://{args.host}"
        # If port omitted, default to Ollama standard 11434
        host_no_proto = args.host.split("://", 1)[1]
        if ":" not in host_no_proto:
            args.host = f"{args.host}:11434"

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
    # Originals first: names containing "copy" or "(1)" sort last, so the
    # clean file name is the one kept.
    for p in sorted(all_pdfs, key=lambda x: (
            "copy" in x.stem.lower() or "(" in x.stem, x.name)):
        h = file_hash(p)
        if h in hashes:
            print(f"  [Duplicate content skipped: {p.name} == {hashes[h]}]")
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

    if args.dry_run:
        print(f"\nDRY RUN -- {len(pdfs)} PDF(s) in {args.pdf_dir}. Nothing "
              "sent to Qwen, nothing written.\n")
        for pdf in pdfs:
            mark = "  (already recorded -- would skip)" if _normalise(
                pdf.stem) in seen else ""
            print(f"  {pdf.name}{mark}")
        return

    review, written, skipped = [], 0, 0
    per_invoice_times = []                          # seconds per PDF actually read
    batch_start = time.perf_counter()

    # Filter out already seen PDFs
    todo_pdfs = []
    for pdf in pdfs:
        if _normalise(pdf.stem) in seen:
            skipped += 1
        else:
            todo_pdfs.append(pdf)

    num_workers = max(1, getattr(args, "workers", 4))
    print(f"\n--- Reading {len(todo_pdfs)} invoice(s) using {num_workers} parallel workers ---\n")

    def process_single_pdf(pdf):
        """Process one PDF: render, Qwen OCR extraction, and field post-processing."""
        started = time.perf_counter()
        try:
            images = to_base64(render_pdf(pdf, dpi=args.dpi,
                                          max_pages=args.max_pages,
                                          poppler_path=args.poppler_path))
            if args.utility:
                fields = extract_utility_invoice(args.host, images, alias_examples=alias_examples)
            else:
                fields = extract_invoice(args.host, images, alias_examples=alias_examples)
                if fields:
                    if _looks_like_utility(fields.get("vendor")):
                        try:
                            util_fields = extract_utility_invoice(args.host, images, alias_examples=alias_examples)
                            if util_fields:
                                fields = util_fields
                                print(f"    [Auto-Detected Utility Bill ({fields.get('vendor')}) -> Applied Utility Extraction Schema]")
                        except Exception as ue:
                            print(f"    (Notice utility re-extraction fallback: {ue})")
        except Exception as e:
            print(f"  [Error reading {pdf.name}]: {e} -- marked for REVIEW")
            return pdf, None, None, None, None, None, time.perf_counter() - started, f"read_error: {e}"

        elapsed = time.perf_counter() - started
        if fields is None:
            return pdf, None, None, None, None, None, elapsed, "qwen_none"

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
        is_util_bill = args.utility or _looks_like_utility(v_check)
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

    from concurrent.futures import ThreadPoolExecutor, as_completed
    import threading

    excel_lock = threading.Lock()
    next_row = last_data_row(sheet, columns) + 1

    if num_workers > 1 and len(todo_pdfs) > 1:
        with ThreadPoolExecutor(max_workers=num_workers) as executor:
            future_to_pdf = {executor.submit(process_single_pdf, p): p for p in todo_pdfs}
            for future in as_completed(future_to_pdf):
                pdf, fields, raw_vendor, qwen_status, raw_property, property_status, elapsed, err = future.result()
                per_invoice_times.append(elapsed)
                if err or fields is None:
                    review.append(pdf.name)
                    print(f"  review: {pdf.name} (extraction failed or unusable) [{elapsed:.1f}s]")
                    continue

                with excel_lock:
                    append_invoice(sheet, columns, fields, pdf.stem, next_row,
                                   raw_vendor=raw_vendor, qwen_status=qwen_status,
                                   raw_property=raw_property, property_status=property_status,
                                   integration_package=fields.get("integration_package"))
                    next_row += 1
                    seen.add(_normalise(pdf.stem))
                    written += 1
                    save_workbook(workbook, args.xlsx)
                    print(f"  filled: {pdf.name} -> {fields['vendor']}  [{elapsed:.1f}s]")
    else:
        for pdf in todo_pdfs:
            pdf, fields, raw_vendor, qwen_status, raw_property, property_status, elapsed, err = process_single_pdf(pdf)
            per_invoice_times.append(elapsed)
            if err or fields is None:
                review.append(pdf.name)
                print(f"  review: {pdf.name} (extraction failed or unusable) [{elapsed:.1f}s]")
                continue

            append_invoice(sheet, columns, fields, pdf.stem, next_row,
                           raw_vendor=raw_vendor, qwen_status=qwen_status,
                           raw_property=raw_property, property_status=property_status,
                           integration_package=fields.get("integration_package"))
            next_row += 1
            seen.add(_normalise(pdf.stem))
            written += 1
            save_workbook(workbook, args.xlsx)
            print(f"  filled: {pdf.name} -> {fields['vendor']}  [{elapsed:.1f}s]")

    save_workbook(workbook, args.xlsx)
    batch_elapsed = time.perf_counter() - batch_start
    print(f"\nDone. {written} written, {skipped} already-recorded skipped, "
          f"{len(review)} sent to human review.")
    processed = len(per_invoice_times)
    if processed:
        average = sum(per_invoice_times) / processed
        print(f"Time: {batch_elapsed:.1f}s total for {processed} invoice(s) "
              f"actually read -- {average:.1f}s average "
              f"(fastest {min(per_invoice_times):.1f}s, "
              f"slowest {max(per_invoice_times):.1f}s). "
              "Skipped duplicates are not counted; they never reach Qwen.")
    if review:
        print("Human review (Qwen reply unusable): " + ", ".join(review))
    print("Property is matched to masters\\properties_data.csv when possible "
          "(canonical name + Integration Package hint); anything uncertain is "
          "written RAW with Property Status = NEEDS_REVIEW for you to confirm. "
          "Then run invoice_recorder.py.")


if __name__ == "__main__":
    main()
