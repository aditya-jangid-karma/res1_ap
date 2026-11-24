"""
extract_deleted_invoices_with_qwen.py -- Resume-ready, parallel Excel & JSON writing,
fast PDF extraction, and local Qwen reading for deleted invoices in the last 2 months.
"""

import os
import sys
import json
import time
import base64
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
OUTPUT_EXCEL = HERE / "deleted_invoices_last_2_months.xlsx"
PROGRESS_JSON = HERE / "deleted_invoices_progress.json"
CACHE_JSON = HERE / "deleted_records_last_2_months.json"
EXCEL_TARGET_FILE = HERE / "Deleted invoices as of 10.16.xlsx"
TEMP_PDF_DIR = HERE / "temp_deleted_pdfs"
SESSION_FILE = HERE / ".bottomline_session.json"
EXEC_LOG_FILE = HERE / "execution_tracking.log"
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://192.168.175.74:11434")


sys.path.insert(0, str(HERE))
from reader_pdf import render_pdf
from reader_qwen import extract_invoice

to_base64 = lambda blist: [base64.b64encode(b).decode("ascii") for b in blist]

def get_grid_records(page):
    """Extracts all invoice records from the active ExtJS grid store on the current page."""
    script = """
    () => {
        if (typeof Ext === 'undefined') return [];
        const grids = Ext.ComponentQuery.query('gridpanel, grid');
        for (let g of grids) {
            if (g.isVisible && g.isVisible()) {
                const store = g.getStore ? g.getStore() : null;
                if (store && store.data && store.data.items && store.data.items.length > 0) {
                    return store.data.items.map(it => ({
                        image_name: it.data.Image_Index_Name || it.data.image_index_GUID,
                        tablekey_id: it.data.Tablekey_Id || it.data.Image_Index_Id,
                        image_index_id: it.data.Image_Index_Id,
                        deleted_date: it.data.image_index_deleted_datetm || it.data.Image_Index_Date_Entered,
                        deleted_by: it.data.deletedby_username,
                        property: it.data.property_name,
                        vendor: it.data.vendor_name,
                        ref: it.data.Image_Index_Ref,
                        doctype: it.data.image_doctype_name
                    }));
                }
            }
        }
        return [];
    }
    """
    try:
        return page.evaluate(script)
    except Exception:
        return []

def save_progress(extracted_rows, force_excel=False):
    """Saves progress to JSON instantly on every row, and flushes to Excel every 10 rows or on force_excel."""
    # 1. Save JSON (instant in-memory dump, no disk lag)
    with open(PROGRESS_JSON, "w", encoding="utf-8") as f:
        json.dump(extracted_rows, f, indent=2, default=str)

    # 2. Flush Excel periodically to avoid 1-2s blocking file I/O per invoice
    if force_excel or len(extracted_rows) % 10 == 0:
        try:
            out_df = pd.DataFrame(extracted_rows)
            for c in out_df.columns:
                if pd.api.types.is_datetime64_any_dtype(out_df[c]):
                    out_df[c] = out_df[c].dt.tz_localize(None)
            out_df.to_excel(OUTPUT_EXCEL, index=False)
        except Exception as e:
            print(f"    (Notice saving Excel: {e})")


def main():
    load_dotenv(PROJECT / ".env")
    url = os.getenv("BOTTOMLINE_URL", "").strip()
    username = os.getenv("BOTTOMLINE_USERNAME", "").strip()
    password = os.getenv("BOTTOMLINE_PASSWORD", "").strip()

    TEMP_PDF_DIR.mkdir(parents=True, exist_ok=True)

    print("\n=======================================================")
    print("  RESUME-READY DELETED INVOICE EXTRACTOR WITH QWEN")
    print("=======================================================\n")

    # 1. Load target Image Names from "Deleted invoices as of 10.16.xlsx"
    target_image_order = []
    target_image_set = set()
    if EXCEL_TARGET_FILE.exists():
        try:
            target_df = pd.read_excel(EXCEL_TARGET_FILE)
            if "Image name" in target_df.columns:
                for img in target_df["Image name"].dropna():
                    clean_img = str(img).strip()
                    if clean_img and clean_img.lower() not in target_image_set:
                        target_image_order.append(clean_img)
                        target_image_set.add(clean_img.lower())
            print(f"✓ Target list: loaded {len(target_image_order)} unique images from '{EXCEL_TARGET_FILE.name}'")
        except Exception as te:
            print(f"Notice loading target Excel: {te}")

    # Load candidate records from cache, filtered strictly to the target list
    candidate_records = []
    if CACHE_JSON.exists():
        try:
            with open(CACHE_JSON, "r", encoding="utf-8") as f:
                raw_records = json.load(f)
            cached_by_name = {str(r.get("image_name") or r.get("Image_Index_Name") or "").strip().lower(): r for r in raw_records}
            if target_image_order:
                # Order candidates in portal grid order (Page 1 -> Page 2 -> ...)
                # so the browser walks through the grid naturally without jumping
                in_cache_targets = [r for r in raw_records if str(r.get("image_name") or r.get("Image_Index_Name") or "").strip().lower() in target_image_set]
                candidate_records.extend(in_cache_targets)

                # Append any remaining target images that weren't in cache
                found_set = set(str(r.get("image_name") or r.get("Image_Index_Name") or "").strip().lower() for r in in_cache_targets)
                for t_img in target_image_order:
                    if t_img.lower() not in found_set:
                        candidate_records.append({"image_name": t_img})
                print(f"✓ Prepared {len(candidate_records)} candidate records matching '{EXCEL_TARGET_FILE.name}' (portal-grid ordered)")
            else:
                candidate_records = raw_records
        except Exception:
            pass

    # Load existing progress (JSON preferred, fallback to Excel)
    extracted_rows = []
    if PROGRESS_JSON.exists():
        try:
            with open(PROGRESS_JSON, "r", encoding="utf-8") as f:
                extracted_rows = json.load(f)
            print(f"✓ Resuming: loaded {len(extracted_rows)} processed records from {PROGRESS_JSON.name}")
        except Exception:
            pass
    elif OUTPUT_EXCEL.exists():
        try:
            prev_df = pd.read_excel(OUTPUT_EXCEL)
            extracted_rows = prev_df.to_dict('records')
            print(f"✓ Resuming: loaded {len(extracted_rows)} processed records from {OUTPUT_EXCEL.name}")
        except Exception:
            pass

    processed_images = set(str(r.get("Image Name") or "").strip().lower() for r in extracted_rows)
    print(f"Previously completed: {len(processed_images)} invoices.")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        if SESSION_FILE.exists():
            context = browser.new_context(storage_state=str(SESSION_FILE))
            print("Reusing existing saved session.")
        else:
            context = browser.new_context()

        # Intercept background PDF responses
        captured_pdfs = {}
        def on_response(response):
            try:
                ct = response.headers.get("content-type", "").lower()
                r_url = response.url.lower()
                if "application/pdf" in ct or ".pdf" in r_url or "showimage" in r_url or "getimage" in r_url:
                    if response.status == 200:
                        body = response.body()
                        if body and body.startswith(b"%PDF"):
                            captured_pdfs[response.url] = body
            except Exception:
                pass

        page = context.new_page()
        page.on("response", handle_response := on_response)

        target_url = "https://residentialone.nexuspayables.com/#Images:showMain:deleted"
        print(f"Navigating to {target_url}...")
        page.goto(target_url)
        page.wait_for_timeout(3000)

        if page.locator("input[type='password']").count() > 0:
            print("Logging in...")
            try:
                page.locator("input[type='text'], input[type='email']").first.fill(username)
                page.locator("input[type='password']").first.fill(password)
                page.keyboard.press("Enter")
            except Exception as e:
                print(f"Login error: {e}")

        print("\n>>> Please complete login/2FA if prompted, and ensure the 'Deleted Images' grid is visible.")
        input("Press ENTER once you are on the 'Deleted Images' screen... ")

        try:
            context.storage_state(path=str(SESSION_FILE))
        except Exception:
            pass

        total_target_count = len(target_image_order) if target_image_order else len(candidate_records)
        print(f"\nProcessing target invoices ({total_target_count} total, {len(processed_images)} already done)...")
        print("Walking grid pages (Page 1 -> 59) to find and extract remaining invoices...\n")

        page_num = 1
        consecutive_empty_pages = 0

        while page_num <= 59:
            # Check if all targets are completed
            if target_image_set and target_image_set.issubset(processed_images):
                print("\n✓ ALL TARGET INVOICES HAVE BEEN PROCESSED!")
                break

            # 1. Read the active ExtJS grid records on current page
            page_records = get_grid_records(page)
            if not page_records:
                consecutive_empty_pages += 1
                if consecutive_empty_pages > 2:
                    print(f"Page {page_num}: No records detected in grid. Finished or end of pages.")
                    break
                page.wait_for_timeout(2000)
                continue
            else:
                consecutive_empty_pages = 0

            first_dt = page_records[0].get("deleted_date", "")[:10]
            last_dt = page_records[-1].get("deleted_date", "")[:10]
            print(f"\n=== [Page {page_num}/59] ({len(page_records)} rows, Range: {first_dt} -> {last_dt}) ===")

            # Find which records on this page belong to our target list and are not yet processed
            page_targets = []
            for r in page_records:
                im_name = str(r.get("image_name") or "").strip()
                if not im_name:
                    continue
                if target_image_set and im_name.lower() not in target_image_set:
                    continue
                if im_name.lower() in processed_images:
                    continue
                page_targets.append(r)

            print(f"  Target items on this page remaining: {len(page_targets)}")

            for rec in page_targets:
                image_name = rec.get("image_name")
                if not image_name or str(image_name).strip().lower() in processed_images:
                    continue

                deleted_date = str(rec.get("deleted_date") or "")
                deleted_by = str(rec.get("deleted_by") or "")
                ref = str(rec.get("ref") or "")
                amt = "$0.00"
                vendor = str(rec.get("vendor") or "")
                prop = str(rec.get("property") or "")
                doc_type = "Invoices"

                current_done = len(processed_images) + 1
                print("-" * 65)
                print(f"[{current_done}/{total_target_count}] (Page {page_num}) Image: {image_name}")

                pdf_path = TEMP_PDF_DIR / f"{image_name}.pdf"

                # 1. Fetch PDF binary stream
                if not pdf_path.exists():
                    captured_pdfs.clear()
                    fetched = False

                    try:
                        row_sel = page.locator(f"tr:has-text('{image_name}')").first
                        if row_sel.count() > 0:
                            row_sel.scroll_into_view_if_needed(timeout=2000)
                            page.wait_for_timeout(300)

                            # Locate paperclip link in the View Image column
                            paperclip = row_sel.locator("td").last.locator("a, img, div").first
                            try:
                                with context.expect_page(timeout=5000) as pinfo:
                                    paperclip.click(timeout=3000)
                                new_tab = pinfo.value
                                new_tab.wait_for_timeout(1000)

                                tab_url = new_tab.url
                                if "showimage.php" in tab_url.lower():
                                    t_res = context.request.get(tab_url)
                                    if t_res.status == 200 and t_res.body().startswith(b"%PDF"):
                                        with open(pdf_path, "wb") as f:
                                            f.write(t_res.body())
                                        print(f"  [OK] PDF downloaded: {pdf_path.name} ({len(t_res.body())} bytes)")
                                        fetched = True

                                if not fetched:
                                    embed = new_tab.locator("embed, object, iframe").first
                                    orig = embed.get_attribute("original-url") or embed.get_attribute("src") or ""
                                    if orig:
                                        abs_src = urljoin(new_tab.url, orig)
                                        s_res = context.request.get(abs_src)
                                        if s_res.status == 200 and s_res.body().startswith(b"%PDF"):
                                            with open(pdf_path, "wb") as f:
                                                f.write(s_res.body())
                                            print(f"  [OK] PDF downloaded from tab embed: {pdf_path.name}")
                                            fetched = True

                                new_tab.close()
                            except Exception as pe:
                                page.wait_for_timeout(600)
                                if captured_pdfs:
                                    for purl, pbytes in captured_pdfs.items():
                                        with open(pdf_path, "wb") as f:
                                            f.write(pbytes)
                                        print(f"  [OK] PDF captured from network: {pdf_path.name}")
                                        fetched = True
                                        break
                                else:
                                    print(f"  Notice on tab open: {pe}")
                        else:
                            print(f"  Row not found in DOM for: {image_name}")
                    except Exception as ce:
                        print(f"  Notice on click: {ce}")

                # If PDF could not be fetched, SKIP and DO NOT mark as completed!
                if not pdf_path.exists():
                    print(f"  (Skipping {image_name}: PDF not yet downloaded - will retry)")
                    continue

                # 2. Extract with Qwen if PDF exists (1 page at 115 DPI cuts VLM processing time in half)
                if pdf_path.exists():
                    try:
                        imgs = to_base64(render_pdf(pdf_path, dpi=115, max_pages=1))
                        q_res = extract_invoice(OLLAMA_HOST, imgs)

                        if q_res:
                            ref = q_res.get("invoice_number") or ref
                            amt = q_res.get("amount") or amt
                            vendor = q_res.get("vendor") or vendor
                            prop = q_res.get("property") or prop

                            v_lower = str(vendor).lower()
                            ref_lower = str(ref).lower()
                            reasoning_text = str(q_res.get("_reasoning") or "").lower()
                            raw_combined = f"{v_lower} {ref_lower} {reasoning_text}"

                            # Strict classification into the 7 user-requested categories
                            if any(k in raw_combined for k in ["disconnect notice", "disconnection notice", "shut off", "cut off notice", "termination notice", "notice of termination", "service termination", "cancel notice"]):
                                doc_type = "Termination Notice"
                            elif any(k in raw_combined for k in ["overdue notice", "past due notice", "past due", "delinquent notice", "delinquent", "late notice", "urgent notice", "demand notice", "final notice"]):
                                doc_type = "Past Dues Notice"
                            elif any(k in raw_combined for k in ["statement of account", "monthly statement", "billing statement", "account statement", "summary statement"]) or ("statement" in raw_combined and "invoice" not in raw_combined):
                                doc_type = "Statement"
                            elif any(k in raw_combined for k in ["payment request", "request for payment", "draw request", "application for payment", "reimbursement request", "check request"]):
                                doc_type = "Payment Request"
                            elif any(k in raw_combined for k in ["information only", "for your records", "info notice", "annual report", "tax document", "w-9", "certificate of insurance", "inspection report", "memorandum"]):
                                doc_type = "Information"
                            elif any(u in v_lower for u in ["pepco", "bge", "potomac", "electric", "power", "washington gas", "columbia gas", "gas light", "wssc", "water", "sewer", "trash", "waste", "recycling", "republic", "comcast", "verizon", "prestige tel", "telecom", "at&t", "spectrum"]):
                                doc_type = "Utility Invoices"
                            else:
                                doc_type = "Invoices"
                    except Exception as qe:
                        print(f"  Qwen read notice: {qe}")

                row_data = {
                    "Image Name": image_name,
                    "Invoice Type": doc_type,
                    "Invoice Number": ref,
                    "Amount": amt,
                    "Vendor Name": vendor,
                    "Property Name": prop,
                    "Deleted Date": deleted_date,
                    "Deleted By": deleted_by,
                    "Scanned Date": str(rec.get("Image_Index_Date_Entered") or "")
                }
                extracted_rows.append(row_data)
                processed_images.add(str(image_name).strip().lower())

                # Print clean status to console
                log_line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{len(processed_images)}/{total_target_count}] Image: {image_name} | Type: {doc_type} | Inv#: {ref} | Amt: {amt} | Vendor: {vendor} | Prop: {prop}"
                print(f"  -> Extracted: [{doc_type}] | Inv#: {ref} | Amt: {amt} | Vendor: {vendor} | Prop: {prop}")

                try:
                    with open(EXEC_LOG_FILE, "a", encoding="utf-8") as lf:
                        lf.write(log_line + "\n")
                except Exception:
                    pass

                # Parallel Save to Excel and JSON
                save_progress(extracted_rows)

            # Advance to Next Page in Bottomline Nexus
            next_btn = page.locator("a.x-tbar-page-next:visible, .x-btn:has(.x-tbar-page-next):visible, [data-qtip='Next Page']:visible").first
            if next_btn.count() == 0 or "disabled" in (next_btn.get_attribute("class") or ""):
                print("\nReached last page of Deleted Images.")
                break

            print(f"-> Advancing to Page {page_num + 1}...")
            next_btn.click()
            page.wait_for_timeout(2500)
            page_num += 1

        # Final flush to ensure all remaining records are written to Excel
        save_progress(extracted_rows, force_excel=True)

        print("\n" + "=" * 65)
        print(f"  [OK] COMPLETE! All records saved to:")
        print(f"    - Excel: {OUTPUT_EXCEL}")
        print(f"    - JSON:  {PROGRESS_JSON}")
        print("=" * 65 + "\n")
        browser.close()

if __name__ == "__main__":
    main()
