"""
extract_deleted_invoices_v2.py -- Step 2 of the deleted-invoice extraction.

GOAL
    For EVERY image named in "Deleted invoices as of 10.16.xlsx" (column
    "Image name"), produce one row with:
        Image Name | Invoice Type | Invoice Number | Amount | Vendor Name |
        Property Name | Deleted Date | Deleted By | Scanned Date | Status | Note
    saved to deleted_invoices_last_2_months.xlsx (+ a JSON copy used to resume).

HOW (plain English)
    1. Reads the image names from the Excel file above.
    2. Looks each one up in deleted_grid_all_pages.json (made by
       scan_deleted_pages.py) to get Nexus's short PDF code (image_index_GUID).
    3. Downloads each PDF straight from Nexus
         https://residentialone.nexuspayables.com/showImage.php?ID=<code>
       from inside the logged-in browser page (no clicking, no grid paging).
    4. Local Qwen (via Ollama) reads the first + last page and returns the
       document type, invoice number, amount, vendor and property.
    5. Prints a final check: "Excel has N images, done X, not done Y" and writes
       the not-done ones, with the reason, to deleted_invoices_not_done.csv.

STATUS VALUES   (only OK counts as finished; anything else is retried next run)
    OK            read successfully
    NOT IN NEXUS  image name not found in the grid scan
    NO PDF        PDF could not be downloaded
    BAD PDF       PDF downloaded but would not open
    QWEN FAILED   Qwen could not read it (Ollama off / timeout / bad reply)

RUN (from the Invoice_Indexer folder)
    python automation/scan_deleted_pages.py                          <- once, first
    python automation/extract_deleted_invoices_v2.py --limit 5       <- small test
    python automation/extract_deleted_invoices_v2.py                 <- everything
    add --download-only to just fetch the PDFs (Qwen not needed for that).
    Safe to stop (Ctrl+C) and re-run at any time: it resumes where it stopped.
"""

import argparse
import base64
import csv
import json
import os
import re
import shutil
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
sys.path.insert(0, str(HERE))
from reader_pdf import render_pdf                      # noqa: E402  (fallback renderer)
from reader_qwen import MODEL, NUM_CTX, _post          # noqa: E402  (same Qwen + same HTTP call as the main reader)
from scan_deleted_pages import open_deleted_grid, GRID_CACHE  # noqa: E402

TARGET_XLSX = Path(os.getenv("TARGET_XLSX", r"C:\Users\Meet Patel\Downloads\Deleted as of 10.20.xlsx"))
PDF_DIR = HERE / "temp_deleted_pdfs"
OUTPUT_EXCEL = HERE / "deleted_invoices_last_2_months.xlsx"
PROGRESS_JSON = HERE / "deleted_invoices_progress.json"
NOT_DONE_CSV = HERE / "deleted_invoices_not_done.csv"
EXEC_LOG = HERE / "execution_tracking.log"
SHOW_IMAGE = "https://residentialone.nexuspayables.com/showImage.php?ID="

COLUMNS = ["Image Name", "Invoice Type", "Invoice Number", "Amount", "Vendor Name",
           "Property Name", "Deleted Date", "Deleted By", "Scanned Date", "Status", "Note"]

CATEGORIES = ["Termination Notice", "Past Dues Notice", "Statement", "Utility Invoices",
              "Payment Request", "Information", "Invoices"]

PDF_DPI = 90               # 90 DPI: ultra-fast inference with sharp text legibility
EXCEL_EVERY = 20           # rewrite the Excel every N images (JSON is saved every image)
STOP_AFTER_FAILS = 10      # failure tolerance before stopping
DEFAULT_WORKERS = 4        # parallel Qwen threads against Ollama

PROMPT = (
    "You are reading ONE document, shown as page image(s): its first page, and its last "
    "page if it has more than one. Return ONLY the fields in the schema. "
    "_reasoning: ONE short sentence - what kind of document this is and who issued it. "
    "doc_type: choose exactly one of: "
    "'Termination Notice' = notice that service will be or was disconnected, shut off, terminated or cancelled; "
    "'Past Dues Notice' = overdue / past due / delinquent / final / demand notice about an unpaid balance; "
    "'Statement' = statement of account or billing statement that lists activity or balances rather than one bill; "
    "'Utility Invoices' = a bill from an electric, gas, water/sewer, trash, internet/phone/cable provider; "
    "'Payment Request' = a request for payment that is not an invoice (draw request, application for payment, "
    "reimbursement or check request); "
    "'Information' = nothing to pay: letters, reports, insurance certificates, W-9s, notices for information only; "
    "'Invoices' = any other vendor invoice or bill (credit memos too). "
    "vendor = the company that ISSUED the document (seller / remit-to payee), never the customer being billed. "
    "invoice_number = as printed, keep leading zeros; null if there is none. "
    "amount = the total amount due / invoice total as printed; null if there is none. "
    "property = the short NAME of the property or building billed (e.g. 'Battery Lane' or 'Victory Court'). "
    "If both a Bill To and a Ship To are shown, take property from the SHIP TO. If only a street address "
    "is shown, leave property null and put the address in service_address. "
    "Never describe the work done as the property, and never invent values."
)


def build_schema():
    return {
        "type": "object",
        "properties": {
            "_reasoning": {"type": "string"},
            "doc_type": {"type": "string", "enum": CATEGORIES},
            "vendor": {"type": ["string", "null"]},
            "invoice_number": {"type": ["string", "null"]},
            "amount": {"type": ["string", "null"]},
            "property": {"type": ["string", "null"]},
            "service_address": {"type": ["string", "null"]},
        },
        "required": ["_reasoning", "doc_type"],
        "additionalProperties": False,
    }


# ---------------------------------------------------------------- small helpers
def to_base64(blist):
    return [base64.b64encode(b).decode("ascii") for b in blist]


def pdf_path(name):
    return PDF_DIR / (re.sub(r"[^A-Za-z0-9._-]", "_", name) + ".pdf")


def clean_amount(raw):
    """'$1,234.50' -> 1234.5 ; '(45.00)' / '45.00-' / '45.00 CR' -> -45.0 ;
    nothing/unreadable -> None (or the original text if it cannot be parsed)."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    neg = (s.startswith("(") and s.endswith(")")) or s.startswith("-") or s.endswith("-") \
        or s.upper().endswith("CR")
    digits = re.sub(r"[^0-9.]", "", s)
    if not re.search(r"\d", digits) or digits.count(".") > 1:
        return s
    try:
        value = float(digits)
    except ValueError:
        return s
    return -value if neg else value


def clean_dt(value):
    """'2026-10-15 07:02:17.007000+00:00' -> '2026-10-15 07:02:17' (as Nexus stores it)."""
    s = str(value or "").strip()
    return s[:19] if re.match(r"\d{4}-\d{2}-\d{2}", s) else s


def pick_property(q_prop, q_addr, nexus_prop):
    """Prefer the name Qwen read from the PDF; fall back to Nexus's value, then the
    service address. A 'name' that is really a sentence (Qwen sometimes copies the job
    description) is rejected. Returns (property, note)."""
    p = (q_prop or "").strip()
    a = (q_addr or "").strip()
    n = (nexus_prop or "").strip()
    nexus_is_real = bool(n) and "property not found" not in n.lower()
    looks_bad = len(p) > 60 or p.endswith(".") or len(p.split()) > 8
    if p and not looks_bad:
        return p, ""
    note = "PDF 'property' looked like a description, ignored. " if p else ""
    if nexus_is_real:
        return n, note
    if a:
        return a, note + "no property name; used service address."
    return n, note


def make_row(name, rec, status, note, **fields):
    """One output row. Starts from Nexus's own values, overridden by what Qwen read."""
    rec = rec or {}
    row = {
        "Image Name": name,
        "Invoice Type": "",
        "Invoice Number": str(rec.get("Image_Index_Ref") or ""),
        "Amount": None,
        "Vendor Name": str(rec.get("vendor_name") or ""),
        "Property Name": str(rec.get("property_name") or ""),
        "Deleted Date": clean_dt(rec.get("image_index_deleted_datetm")),
        "Deleted By": str(rec.get("deletedby_username") or ""),
        "Scanned Date": clean_dt(rec.get("Image_Index_Date_Entered")),
        "Status": status,
        "Note": note,
    }
    row.update(fields)
    return row


# ------------------------------------------------------------------ input loading
def load_targets():
    if not TARGET_XLSX.exists():
        sys.exit(f"Cannot find the target list: {TARGET_XLSX}")
    df = pd.read_excel(TARGET_XLSX)
    col = next((c for c in df.columns if str(c).strip().lower() == "image name"), None)
    if col is None:
        sys.exit(f"No 'Image name' column in {TARGET_XLSX.name}. Headers found: {list(df.columns)}")
    seen, names, rows = set(), [], 0
    for v in df[col].dropna():
        name = str(v).strip()
        if not name:
            continue
        rows += 1
        if name.lower() not in seen:
            seen.add(name.lower())
            names.append(name)
    print(f"Target list: {rows} rows in '{TARGET_XLSX.name}' -> {len(names)} unique images")
    return names


def load_grid():
    if not GRID_CACHE.exists():
        sys.exit(f"{GRID_CACHE.name} not found. Run first:  python automation/scan_deleted_pages.py")
    with open(GRID_CACHE, "r", encoding="utf-8") as f:
        data = json.load(f)
    records = data["records"] if isinstance(data, dict) else data
    by_name = {}
    for r in records:
        key = str(r.get("Image_Index_Name") or "").strip().lower()
        if key:
            by_name[key] = r
    when = data.get("scanned_at", "unknown time") if isinstance(data, dict) else "unknown time"
    print(f"Grid scan: {len(by_name)} images (scanned {when})")
    return by_name


def load_progress():
    """Return {lowercase image name: row} for rows already finished (Status == OK).
    Rows from the old version of the script have no Status and are redone."""
    data = []
    if PROGRESS_JSON.exists():
        try:
            with open(PROGRESS_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as err:
            sys.exit(f"Progress file {PROGRESS_JSON.name} could not be read ({err}).\n"
                     f"Stopped so nothing gets overwritten. Rename or repair it, then re-run.")
        shutil.copy2(PROGRESS_JSON, PROGRESS_JSON.with_suffix(".bak.json"))
    elif OUTPUT_EXCEL.exists():
        try:
            data = pd.read_excel(OUTPUT_EXCEL).fillna("").to_dict("records")
        except Exception:
            data = []
    done = {}
    for r in data:
        if r.get("Status") == "OK" and r.get("Image Name"):
            done[str(r["Image Name"]).strip().lower()] = r
    print(f"Progress: {len(data)} saved rows, {len(done)} already OK (the rest will be redone)")
    return done


def save_progress(rows, targets, force_excel=False, counter=[0]):
    """JSON every time (written safely: temp file then rename). Excel every EXCEL_EVERY images."""
    ordered = [rows[t.lower()] for t in targets if t.lower() in rows]
    tmp = PROGRESS_JSON.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(ordered, f, indent=1, default=str)
    os.replace(tmp, PROGRESS_JSON)

    counter[0] += 1
    if force_excel or counter[0] % EXCEL_EVERY == 0:
        try:
            pd.DataFrame(ordered, columns=COLUMNS).to_excel(OUTPUT_EXCEL, index=False)
        except Exception as err:
            print(f"    (!! Could not write {OUTPUT_EXCEL.name}: {err} -- close it in Excel. "
                  f"Progress is safe in the JSON.)")


# ----------------------------------------------------------------- phase 1: PDFs
FETCH_B64_JS = """
async (url) => {
  const r = await fetch(url, {credentials: 'include'});
  const buf = new Uint8Array(await r.arrayBuffer());
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < buf.length; i += chunk) {
    bin += String.fromCharCode.apply(null, buf.subarray(i, i + chunk));
  }
  return {status: r.status, ctype: r.headers.get('content-type') || '', b64: btoa(bin)};
}
"""


def download_pdfs(todo, by_name):
    """Download the PDF of every name in `todo`. Returns {name: why it failed}."""
    fails = {}
    with sync_playwright() as playwright:
        browser, context, page = open_deleted_grid(playwright)
        try:
            in_a_row = 0
            for i, name in enumerate(todo, 1):
                guid = str(by_name[name.lower()].get("image_index_GUID") or "").strip()
                if not guid:
                    fails[name] = "grid record has no image_index_GUID"
                    print(f"[{i}/{len(todo)}] x {name}: {fails[name]}")
                    continue
                try:
                    res = page.evaluate(FETCH_B64_JS, SHOW_IMAGE + guid)
                    body = base64.b64decode(res["b64"]) if res.get("b64") else b""
                    if body.startswith(b"%PDF"):
                        part = pdf_path(name).with_suffix(".part")
                        part.write_bytes(body)
                        os.replace(part, pdf_path(name))
                        in_a_row = 0
                        print(f"[{i}/{len(todo)}] ok {name} ({len(body) // 1024} KB)")
                        continue
                    fails[name] = (f"Nexus returned HTTP {res.get('status')}, type "
                                   f"'{res.get('ctype')}', {len(body)} bytes - not a PDF")
                except Exception as err:
                    fails[name] = f"download error: {err}"
                in_a_row += 1
                print(f"[{i}/{len(todo)}] x {name}: {fails[name]}")
                if in_a_row >= STOP_AFTER_FAILS:
                    print(f"\n!! {STOP_AFTER_FAILS} downloads failed in a row - the Nexus login probably "
                          f"expired. Stopping the download phase. Re-run and log in again.")
                    break
        finally:
            browser.close()
    return fails


# ------------------------------------------------------------------ phase 2: Qwen
def check_ollama(host):
    """(ok, message). Fails fast if the Qwen machine is off or lacks the model."""
    try:
        with urllib.request.urlopen(host.rstrip("/") + "/api/tags", timeout=10) as resp:
            names = [m.get("name", "") for m in json.loads(resp.read().decode()).get("models", [])]
    except Exception as err:
        return False, f"cannot reach Ollama at {host} ({err})"
    if MODEL not in names:
        return False, f"Ollama at {host} is up but model '{MODEL}' is not installed (has: {names})"
    return True, f"Ollama OK at {host} with {MODEL}"


def render_first_and_last(path, dpi=PDF_DPI):
    """First page, plus the last page when there is more than one (totals are usually last)."""
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
    except ImportError:
        return render_pdf(path, dpi=dpi, max_pages=1)
    doc = fitz.open(path)
    try:
        scale = dpi / 72.0
        return [doc.load_page(0).get_pixmap(matrix=fitz.Matrix(scale, scale)).tobytes("png")]
    finally:
        doc.close()


def ask_qwen(host, images):
    """(reply dict, None) or (None, reason). Two tries."""
    payload = {"model": MODEL, "prompt": PROMPT, "images": images, "stream": False,
               "format": build_schema(), "options": {"num_ctx": NUM_CTX}}
    reason = ""
    for attempt in (1, 2):
        try:
            obj = json.loads(_post(host, payload).get("response", ""))
            if obj.get("doc_type") not in CATEGORIES:
                raise ValueError("reply had no valid doc_type")
            return obj, None
        except Exception as err:
            reason = f"{type(err).__name__}: {err}"
            print(f"    (Qwen attempt {attempt} failed: {reason})")
    return None, reason


def row_from_qwen(name, rec, obj):
    nexus_amt = clean_amount(rec.get("Image_Index_Amount"))
    amount = clean_amount(obj.get("amount"))
    if amount is None and nexus_amt not in (None, 0, 0.0, ""):
        amount = nexus_amt
    vendor = (obj.get("vendor") or "").strip() or str(rec.get("vendor_name") or "")
    inv_no = (obj.get("invoice_number") or "").strip() or str(rec.get("Image_Index_Ref") or "")
    prop, note = pick_property(obj.get("property"), obj.get("service_address"), rec.get("property_name"))
    if not (vendor or inv_no or amount):
        note += "nothing readable on the document."
    return make_row(name, rec, "OK", note.strip(), **{
        "Invoice Type": obj["doc_type"], "Invoice Number": inv_no, "Amount": amount,
        "Vendor Name": vendor, "Property Name": prop})


# ------------------------------------------------------------------------ report
def final_report(rows, targets):
    counts = {}
    not_done = []
    for t in targets:
        row = rows.get(t.lower())
        status = row["Status"] if row else "NOT PROCESSED YET"
        counts[status] = counts.get(status, 0) + 1
        if status != "OK":
            not_done.append((t, status, (row or {}).get("Note", "")))
    with open(NOT_DONE_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Image Name", "Status", "Note"])
        w.writerows(not_done)

    ok = counts.get("OK", 0)
    print("\n" + "=" * 65)
    print(f"  CHECK AGAINST THE EXCEL LIST")
    print(f"  Excel has {len(targets)} unique images")
    print(f"  Done (OK): {ok}     Not done: {len(not_done)}")
    for status, n in sorted(counts.items()):
        if status != "OK":
            print(f"      {status}: {n}")
    if not_done:
        print(f"  Details (with reasons): {NOT_DONE_CSV.name}")
        print("  Re-run the same command to retry them.")
    else:
        print("  ALL IMAGES DONE.")
    print(f"  Results: {OUTPUT_EXCEL.name}  (+ {PROGRESS_JSON.name})")
    print("=" * 65 + "\n")


# -------------------------------------------------------------------------- main
def main():
    parser = argparse.ArgumentParser(description="Extract details of deleted invoices listed in the Excel file.")
    parser.add_argument("--limit", type=int, default=0, help="only process this many images (test run)")
    parser.add_argument("--download-only", action="store_true", help="fetch PDFs, skip the Qwen reading")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help=f"number of concurrent Qwen threads (default {DEFAULT_WORKERS})")
    args = parser.parse_args()

    load_dotenv(PROJECT / ".env")
    host = os.getenv("OLLAMA_HOST", "http://192.168.175.74:11434")
    PDF_DIR.mkdir(parents=True, exist_ok=True)

    print("\n=======================================================")
    print("  DELETED INVOICE EXTRACTOR  (Excel list -> PDFs -> Qwen)")
    print(f"  Parallel Mode: {args.workers} workers | DPI: {PDF_DPI}")
    print("=======================================================\n")

    targets = load_targets()
    by_name = load_grid()
    rows = load_progress()

    qwen_ok = False
    if not args.download_only:
        qwen_ok, msg = check_ollama(host)
        print(("OK: " if qwen_ok else "!! ") + msg)
        if not qwen_ok:
            print("   PDFs will still be downloaded; fix Ollama and re-run to do the reading.")

    # work out what each remaining image needs
    todo_download = []
    for t in targets:
        if t.lower() in rows or t.lower() not in by_name:
            continue
        if not pdf_path(t).exists():
            todo_download.append(t)
    if args.limit:
        todo_download = todo_download[:args.limit]

    fails = {}
    if todo_download:
        print(f"\nPhase 1: downloading {len(todo_download)} PDFs (already-saved ones are skipped)...")
        fails = download_pdfs(todo_download, by_name)
    else:
        print("\nPhase 1: no PDFs to download (all cached on disk).")

    if args.download_only or not qwen_ok:
        have = sum(1 for t in targets if pdf_path(t).exists())
        print(f"\nPDFs on disk: {have} of {len(targets)}. Stopping here (no Qwen reading this run).")
        return

    # Phase 2 - read with Qwen, in parallel
    pending = []
    for t in targets:
        key = t.lower()
        if key in rows:
            continue
        if key not in by_name:
            rows[key] = make_row(t, None, "NOT IN NEXUS",
                                 "image name not found in the grid scan (re-run scan_deleted_pages.py if it is recent)")
        elif not pdf_path(t).exists():
            rows[key] = make_row(t, by_name[key], "NO PDF", fails.get(t, "PDF not downloaded"))
        else:
            pending.append(t)
    if args.limit:
        pending = pending[:args.limit]

    num_workers = max(1, args.workers)
    print(f"\nPhase 2: reading {len(pending)} PDFs with {num_workers} parallel workers (Qwen {MODEL})...")

    from concurrent.futures import ThreadPoolExecutor, as_completed
    import threading

    lock = threading.Lock()
    done_count = 0
    total_pending = len(pending)
    consecutive_fails = 0

    def process_one(name):
        key = name.lower()
        rec = by_name[key]
        t0 = time.time()
        try:
            images = to_base64(render_first_and_last(pdf_path(name)))
        except Exception as err:
            return name, make_row(name, rec, "BAD PDF", f"could not open the PDF: {err}"), time.time() - t0, False

        obj, why = ask_qwen(host, images)
        if obj is None:
            return name, make_row(name, rec, "QWEN FAILED", why), time.time() - t0, False
        return name, row_from_qwen(name, rec, obj), time.time() - t0, True

    try:
        with ThreadPoolExecutor(max_workers=num_workers) as executor:
            future_to_name = {executor.submit(process_one, name): name for name in pending}
            for future in as_completed(future_to_name):
                name, row, elapsed, success = future.result()
                with lock:
                    done_count += 1
                    key = name.lower()
                    rows[key] = row
                    if success:
                        consecutive_fails = 0
                    else:
                        consecutive_fails += 1

                    amt = row["Amount"] if row["Amount"] is not None else ""
                    line = (f"[{row['Status']}] {row['Invoice Type'] or '-'} | Inv#: {row['Invoice Number']} | "
                            f"Amt: {amt} | Vendor: {row['Vendor Name']} | Prop: {row['Property Name']}"
                            f"{' | NOTE: ' + row['Note'] if row['Note'] else ''}")
                    print(f"[{done_count}/{total_pending}] {name}\n  -> {line}  ({elapsed:.1f}s)")
                    try:
                        with open(EXEC_LOG, "a", encoding="utf-8") as lf:
                            lf.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {name} {line}\n")
                    except Exception:
                        pass
                    save_progress(rows, targets)

                    if consecutive_fails >= STOP_AFTER_FAILS:
                        print(f"\n!! {STOP_AFTER_FAILS} Qwen reads failed in a row - stopping. Fix and resume.")
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
    except KeyboardInterrupt:
        print("\nStopped by you. Progress is saved; re-run to continue.")

    save_progress(rows, targets, force_excel=True)
    final_report(rows, targets)


if __name__ == "__main__":
    main()

