r"""TASK 3 -- Bill Verification (FIFO). Sorts every bill in invoice\ into the
"query bills" exception folders BEFORE script 2 runs. Clean bills stay put.

Checks, in Kapil's priority order (a bill with several issues goes to the first
match; the report lists them all):
  1. property not found   -> query bills\property not found
  2. vendor not found     -> query bills\vendor not found
  3. duplicate invoice    -> query bills\dublicate invocie   (folder name as created)
  4. bill not in platform -> query bills\bill not found      (not in the portal's
                             Images to be Indexed queue)

Data: invoices.xlsx (written by script 1) + masters\ + the print-name maps
masters\vendor_print_name_map.csv / property_print_name_map.csv.
FIFO: bills are processed oldest first (timestamp prefix of the image name).
Portal access is READ-ONLY: log in, open Images to be Indexed, read the names.
It never indexes, saves, deletes or opens an image. Nothing is ever deleted;
a move never overwrites (a numeric suffix is added).

  python script\3_bill_verification.py --dry-run         plan + report, no move
  python script\3_bill_verification.py                   real run
  python script\3_bill_verification.py --no-portal       skip the bill-not-found check
"""
import argparse
import csv
import hashlib
import importlib.util
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
sys.path.insert(0, str(HERE))
INVOICE_DIR = PROJECT / "invoice"
QUERY_DIR = PROJECT / "query bills"
MASTERS = PROJECT / "masters"
XLSX = PROJECT / "invoices.xlsx"
FOLDERS = {                       # priority order = Kapil's order
    "property": "property not found",
    "vendor": "vendor not found",
    "duplicate": "dublicate invocie",
    "queue": "bill not found",
}
VENDOR_MAP = MASTERS / "vendor_print_name_map.csv"
PROPERTY_MAP = MASTERS / "property_print_name_map.csv"


def norm(t):
    return re.sub(r"[^a-z0-9]+", " ", str(t or "").lower()).strip()


def stem_key(name):
    t = str(name or "").strip()
    if t.lower().endswith(".pdf"):
        t = t[:-4]
    return norm(t)


def fifo_key(stem, path):
    m = re.match(r"(\d{12})", stem)
    return (0, m.group(1)) if m else (1, datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d%H%M"))


def read_csv(path):
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def file_hash(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def sheet_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    head = [str(c or "").strip() for c in rows[0]]
    return head, [dict(zip(head, r)) for r in rows[1:]]


def g(d, k):
    v = d.get(k)
    return "" if v is None else str(v).strip()


# ----------------------------------------------------------- mapping files
def build_maps(records):
    """Create the print-name -> full-name CSVs from the sheet if absent."""
    if not VENDOR_MAP.exists():
        seen, out = set(), []
        for d in records:
            raw, full, code = g(d, "Qwen Vendor Raw"), g(d, "Vendor"), g(d, "Vendor Code")
            if raw and full and g(d, "Qwen Status") == "READ_OK" and (norm(raw), norm(full)) not in seen:
                seen.add((norm(raw), norm(full)))
                out.append([raw, full, code, "invoices.xlsx (Qwen Vendor Raw -> Vendor, READ_OK)"])
        with open(VENDOR_MAP, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["print_name_on_invoice", "full_name_mapped", "code", "source"])
            w.writerows(out)
        print(f"  created {VENDOR_MAP.name}: {len(out)} rows")
    if not PROPERTY_MAP.exists():
        seen, out = set(), []
        for r in read_csv(MASTERS / "property_aliases.csv"):
            raw, full = r.get("raw", ""), r.get("correct", "")
            if raw and full and (norm(raw), norm(full)) not in seen:
                seen.add((norm(raw), norm(full)))
                out.append([raw, full, "", "masters/property_aliases.csv"])
        for d in records:
            full, code = g(d, "Property"), g(d, "Property Code")
            for raw in (g(d, "Qwen Property Raw"), g(d, "Service Address")):
                if raw and full and g(d, "Property Status") == "READ_OK" and (norm(raw), norm(full)) not in seen:
                    seen.add((norm(raw), norm(full)))
                    out.append([raw, full, code, "invoices.xlsx (Property Status READ_OK)"])
        with open(PROPERTY_MAP, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["print_name_on_invoice", "full_name_mapped", "code", "source"])
            w.writerows(out)
        print(f"  created {PROPERTY_MAP.name}: {len(out)} rows")


# ------------------------------------------------------------ portal queue
def read_portal_queue(headless, log):
    """Read-only: names in Images to be Indexed, all pages."""
    load_dotenv(PROJECT / ".env")
    spec = importlib.util.spec_from_file_location("s2", HERE / "2_invoice_recorder_fill_web.py")
    s2 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(s2)
    import json
    from playwright.sync_api import sync_playwright
    url = os.getenv("BOTTOMLINE_URL")
    if not url:
        raise RuntimeError("BOTTOMLINE_URL missing in .env")
    selectors = json.loads((HERE / "selectors.json").read_text(encoding="utf-8"))
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        ctx = (browser.new_context(storage_state=str(s2.SESSION_FILE))
               if s2.SESSION_FILE.exists() else browser.new_context())
        page = ctx.new_page()
        try:
            if not s2.login(page, url, log, auto=True):
                raise RuntimeError("portal login did not complete (2FA or wrong login)")
            page.locator(selectors["queue"]["menu_opener"] +
                         f':has-text("{selectors["queue"]["menu_opener_text"]}")'
                         ).first.wait_for(state="visible", timeout=45000)
            if not s2.return_to_queue(page, selectors, log):
                raise RuntimeError("could not open Images to be Indexed")
            s2.set_page_size(page, selectors, log)
            index = s2.read_queue_fast(page, selectors, log)
            return list(index.keys()), s2.same_image
        finally:
            browser.close()


# --------------------------------------------------------------- main logic
def main():
    ap = argparse.ArgumentParser(description="Bill verification (FIFO) -> query bills folders.")
    ap.add_argument("--dry-run", action="store_true", help="plan + report only; move nothing")
    ap.add_argument("--no-portal", action="store_true", help="skip the bill-not-found (portal queue) check")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--force-bill-not-found", action="store_true",
                    help="move bill-not-found bills even if the queue read looks unreliable")
    args = ap.parse_args()
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    if not XLSX.exists():
        sys.exit(f"STOPPED: no {XLSX}")
    wb = load_workbook(XLSX)
    ws = wb["Invoices"]
    head, records = sheet_rows(ws)
    by_stem = {stem_key(g(d, "Image Name")): (i + 2, d) for i, d in enumerate(records) if g(d, "Image Name")}

    build_maps(records)
    vmap = {norm(r["print_name_on_invoice"]): r["full_name_mapped"] for r in read_csv(VENDOR_MAP)}
    pmap = {norm(r["print_name_on_invoice"]): r["full_name_mapped"] for r in read_csv(PROPERTY_MAP)}

    vend_master = {}
    for r in read_csv(MASTERS / "vendors_data.csv"):
        for k in ("vendor", "defaultErpVendor"):
            if r.get(k):
                vend_master.setdefault(norm(r[k]), r[k])
    prop_master = {}
    for r in read_csv(MASTERS / "properties_data.csv"):
        for k in ("Property Name", "Property Short Name", "Dropdown Name"):
            if r.get(k):
                prop_master.setdefault(norm(r[k]), r.get("Property Name") or r[k])
    print(f"  masters: {len(vend_master)} vendor names, {len(prop_master)} property names; "
          f"maps: {len(vmap)} vendor, {len(pmap)} property print names")

    pdfs = sorted(INVOICE_DIR.glob("*.pdf"), key=lambda p: fifo_key(p.stem, p))
    pdfs += [p for p in INVOICE_DIR.glob("*.PDF") if p not in pdfs]
    pdfs = sorted(set(pdfs), key=lambda p: fifo_key(p.stem, p))
    print(f"  {len(pdfs)} PDF(s) in invoice\\ (oldest first)")

    # portal queue
    queue_names, same_image, queue_note = None, None, ""
    if args.no_portal:
        queue_note = "skipped (--no-portal)"
    else:
        try:
            queue_names, same_image = read_portal_queue(args.headless, lambda m: print(f"  [portal] {m}"))
            print(f"  queue names read: {len(queue_names)}; first 15:")
            for n in queue_names[:15]:
                print(f"     {n}")
            queue_note = f"{len(queue_names)} names read"
        except Exception as e:
            queue_note = f"portal read FAILED: {e}"
            print(f"  !! {queue_note} -- bill-not-found check skipped")

    # classify
    hashes = {}
    results = []
    filled_keys = {}
    for i, d in enumerate(records):
        if g(d, "Status") == "FILLED":
            filled_keys.setdefault((norm(g(d, "Vendor")), norm(g(d, "Invoice Number"))), g(d, "Image Name"))
    seen_keys = {}
    for p in pdfs:
        stem = p.stem
        hit = by_stem.get(stem_key(stem))
        res = {"image": stem, "file": p, "reasons": [], "folder": None, "row": None,
               "vendor_printed": "", "vendor_master": "", "property_printed": "", "property_master": "",
               "invoice": "", "amount": "", "status": "", "note": ""}
        h = file_hash(p)
        if hit is None:
            res["note"] = "no row in invoices.xlsx -- run script 1 first; left in place"
            results.append(res)
            hashes.setdefault(h, stem)
            continue
        row, d = hit
        res.update(row=row, vendor_printed=g(d, "Qwen Vendor Raw"), property_printed=g(d, "Qwen Property Raw") or g(d, "Service Address"),
                   invoice=g(d, "Invoice Number"), amount=g(d, "Amount"), status=g(d, "Status"))
        if g(d, "Status") in ("NOT_INVOICE", "FILLED"):
            res["note"] = f"Status {g(d, 'Status')} -- not checked, left in place"
            results.append(res)
            hashes.setdefault(h, stem)
            continue
        reasons = []
        # property
        pv = g(d, "Property")
        pfull = prop_master.get(norm(pv)) or prop_master.get(norm(pmap.get(norm(pv), ""))) \
            or prop_master.get(norm(pmap.get(norm(res["property_printed"]), "")))
        res["property_master"] = pfull or ""
        if not pfull:
            reasons.append(("property", f"property '{pv or res['property_printed'] or 'blank'}' not in properties master (Property Status {g(d, 'Property Status') or 'blank'})"))
        # vendor
        vv = g(d, "Vendor")
        vfull = vend_master.get(norm(vv)) or vend_master.get(norm(vmap.get(norm(vv), ""))) \
            or vend_master.get(norm(vmap.get(norm(res["vendor_printed"]), "")))
        res["vendor_master"] = vfull or ""
        if not vfull:
            reasons.append(("vendor", f"vendor '{vv or res['vendor_printed'] or 'blank'}' not in vendors master (Qwen Status {g(d, 'Qwen Status') or 'blank'})"))
        # duplicate
        inv = norm(g(d, "Invoice Number"))
        key = (norm(vv), inv)
        dup = None
        if inv and key[0]:
            placeholder = inv.startswith("n a")
            amt = g(d, "Amount")
            if key in filled_keys and filled_keys[key] != g(d, "Image Name"):
                dup = f"same vendor+invoice # already FILLED ({filled_keys[key]})"
            elif key in seen_keys and (not placeholder or seen_keys[key][1] == amt):
                dup = f"same vendor+invoice # as earlier bill {seen_keys[key][0]}"
            else:
                seen_keys.setdefault(key, (g(d, "Image Name"), amt))
        if h in hashes:
            dup = f"byte-identical to {hashes[h]}"
        hashes.setdefault(h, stem)
        if dup:
            reasons.append(("duplicate", dup))
        # queue
        if queue_names is not None:
            if not any(same_image(q, g(d, "Image Name") or stem) or same_image(q, stem) for q in queue_names):
                reasons.append(("queue", "image name not found in portal Images to be Indexed queue"))
        res["reasons"] = reasons
        for k in FOLDERS:
            if any(r[0] == k for r in reasons):
                res["folder"] = k
                break
        results.append(res)

    # reliability of queue check
    checked = [r for r in results if r["row"] and r["status"] not in ("NOT_INVOICE", "FILLED")]
    qmiss = [r for r in checked if any(x[0] == "queue" for x in r["reasons"])]
    unreliable = bool(queue_names is not None and checked and
                      (len(checked) - len(qmiss)) / len(checked) < 0.2)
    if queue_names is not None:
        print(f"  queue check: {len(checked) - len(qmiss)} of {len(checked)} checked bills found in the queue"
              + ("  -> LOOKS UNRELIABLE / bills not uploaded" if unreliable else ""))
    hold_queue = unreliable and not args.force_bill_not_found

    # move
    moved = 0
    for r in results:
        r["moved"] = "no"
        if not r["folder"]:
            continue
        # when the queue check is held back, a bill whose ONLY reason is 'queue' is not moved
        if r["folder"] == "queue" and hold_queue:
            r["note"] = "bill-not-found NOT moved: queue read looks unreliable (see report)"
            continue
        dest = QUERY_DIR / FOLDERS[r["folder"]]
        r["destination"] = str(dest)
        if args.dry_run:
            r["moved"] = "dry-run"
            continue
        dest.mkdir(parents=True, exist_ok=True)
        target = dest / r["file"].name
        n = 1
        while target.exists():
            target = dest / f"{r['file'].stem} ({n}){r['file'].suffix}"
            n += 1
        shutil.move(str(r["file"]), str(target))
        r["moved"] = "yes"
        moved += 1

    # verification note column in the sheet (only when really moving)
    if not args.dry_run:
        col = head.index("Verification Note") + 1 if "Verification Note" in head else len(head) + 1
        if "Verification Note" not in head:
            ws.cell(1, col).value = "Verification Note"
        for r in results:
            if r["row"] and r["reasons"]:
                ws.cell(r["row"], col).value = (f"{datetime.now():%Y-%m-%d %H:%M} -> " +
                                                (FOLDERS[r["folder"]] if r["moved"] == "yes" else "flagged, not moved") +
                                                ": " + "; ".join(x[1] for x in r["reasons"]))[:500]
        try:
            wb.save(XLSX)
        except PermissionError:
            print("  !! could not save invoices.xlsx (open in Excel?) -- report only")

    # report
    out = PROJECT / f"verification_report_{stamp}{'_dryrun' if args.dry_run else ''}.xlsx"
    rw = Workbook()
    v = rw.active
    v.title = "Verification"
    v.append(["Order (FIFO)", "Image", "Sheet row", "Status", "Vendor printed on invoice", "Vendor master (full name)",
              "Property printed on invoice", "Property master (full name)", "Invoice #", "Amount",
              "Reasons (all)", "Destination folder", "Moved", "Note"])
    for c in v[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for n, r in enumerate(results, 1):
        v.append([n, r["image"], r["row"], r["status"], r["vendor_printed"], r["vendor_master"],
                  r["property_printed"], r["property_master"], r["invoice"], r["amount"],
                  " | ".join(x[1] for x in r["reasons"]), FOLDERS.get(r["folder"], "") if r["folder"] else "",
                  r["moved"], r["note"]])
    s = rw.create_sheet("Summary")
    s.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")])
    s.append(["Mode", "DRY-RUN" if args.dry_run else "REAL"])
    s.append(["Portal queue", queue_note])
    for k, f in FOLDERS.items():
        s.append([f, sum(1 for r in results if r["folder"] == k), "moved: " + str(sum(1 for r in results if r["folder"] == k and r["moved"] == "yes"))])
    s.append(["clean (stay in invoice\\)", sum(1 for r in results if r["row"] and not r["reasons"] and r["status"] not in ("NOT_INVOICE", "FILLED"))])
    s.append(["not checked (no row / NOT_INVOICE / FILLED)", sum(1 for r in results if not r["row"] or r["status"] in ("NOT_INVOICE", "FILLED"))])
    if queue_names is not None:
        q = rw.create_sheet("Queue names (first 60)")
        for n in queue_names[:60]:
            q.append([n])
    rw.save(out)
    print(f"\n  Report: {out}")
    for k, f in FOLDERS.items():
        print(f"  {f}: {sum(1 for r in results if r['folder'] == k)} flagged, {sum(1 for r in results if r['folder'] == k and r['moved'] == 'yes')} moved")
    print(f"  clean: {sum(1 for r in results if r['row'] and not r['reasons'] and r['status'] not in ('NOT_INVOICE', 'FILLED'))}")
    for r in results:
        print(f"   {r['image'][:60]:60} | {r['status'] or '-':11} | {FOLDERS.get(r['folder'], 'clean') if r['folder'] else ('clean' if r['row'] else 'no row')} | {r['moved']} | {' | '.join(x[1] for x in r['reasons'])[:150]}")


if __name__ == "__main__":
    main()

