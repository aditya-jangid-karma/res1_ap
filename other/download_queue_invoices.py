"""
download_queue_invoices.py -- Download candidate invoices from Bottomline Nexus queue.

Inspects the 'Images to be Indexed' queue in Bottomline Nexus.
Filters and downloads up to N candidate invoices to Z:\\Res-1-indexer\\input\\pdfs.
Strictly verifies document type via local Qwen vision AI before saving:
Only 'Invoices' (Standard Vendor Invoices) and 'Utility Invoices' are saved.
Other types (Statements, Past Due Notices, Termination Notices, Payment Requests, Information)
are skipped.
"""

import os
import sys
import json
import time
import re
import base64
import urllib.request
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
PDF_DIR = Path(r"Z:\Res-1-indexer\input\pdfs")
SESSION_FILE = HERE / ".bottomline_session.json"
SELECTORS_FILE = HERE / "selectors.json"
MASTERS_DIR = PROJECT / "masters"

sys.path.insert(0, str(HERE))
from queue_navigator import (
    wait_for_queue_screen,
    read_queue,
    go_to_queue_page,
    normalise_name,
    return_to_queue,
    wait_for_queue_rows
)

ALLOWED_DOC_TYPES = {"Invoices", "Utility Invoices", "Standard Invoices", "Utility Invoice"}
CATEGORIES = [
    "Termination Notice", "Past Dues Notice", "Statement", "Utility Invoices",
    "Payment Request", "Information", "Invoices"
]

def render_first_page_base64(pdf_bytes, dpi=90):
    """Render page 0 of in-memory PDF bytes to a base64-encoded PNG string."""
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(doc) == 0:
            doc.close()
            return None
        page = doc.load_page(0)
        scale = dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=matrix)
        png_data = pix.tobytes("png")
        doc.close()
        return base64.b64encode(png_data).decode("ascii")
    except Exception as e:
        print(f"    (Notice rendering PDF for classification: {e})")
        return None

def classify_doc_type(pdf_bytes, host="http://192.168.175.74:11434"):
    """
    Classify document type using Qwen vision model.
    Returns (doc_type, reasoning).
    If classification fails or host is unreachable, returns ('Invoices', 'fallback_allowed').
    """
    b64_img = render_first_page_base64(pdf_bytes, dpi=90)
    if not b64_img:
        return "Invoices", "unable to render first page; defaulting to allow"

    prompt = (
        "You are reading ONE document page. Return ONLY the JSON schema. "
        "_reasoning: ONE short sentence explaining what document this is. "
        "doc_type: choose exactly one of: "
        "'Termination Notice' = notice that service will be or was disconnected, shut off, terminated or cancelled; "
        "'Past Dues Notice' = overdue / past due / delinquent / final / demand notice about an unpaid balance; "
        "'Statement' = statement of account or billing statement listing activity/balances instead of a single bill; "
        "'Utility Invoices' = utility bill from electric, gas, water/sewer, trash, internet/phone provider; "
        "'Payment Request' = draw request, application for payment, check request or reimbursement; "
        "'Information' = letters, insurance certs, W-9s, reports, general information without amount to pay; "
        "'Invoices' = standard vendor invoice, bill, or credit memo."
    )
    schema = {
        "type": "object",
        "properties": {
            "_reasoning": {"type": "string"},
            "doc_type": {"type": "string", "enum": CATEGORIES}
        },
        "required": ["_reasoning", "doc_type"],
        "additionalProperties": False
    }
    payload = {
        "model": "qwen2.5vl:7b",
        "prompt": prompt,
        "images": [b64_img],
        "stream": False,
        "format": schema,
        "options": {"num_ctx": 4096}
    }
    try:
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            host.rstrip("/") + "/api/generate",
            data=data,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=40) as response:
            res_json = json.loads(response.read().decode("utf-8"))
            body = json.loads(res_json.get("response", "{}"))
            return body.get("doc_type", "Invoices"), body.get("_reasoning", "")
    except Exception as e:
        print(f"    (Notice document classification call: {e})")
        # On error/timeout, return 'Invoices' to avoid blocking queue downloads
        return "Invoices", f"classification_fallback: {e}"

def inspect_and_download_invoice(page, wanted_name):
    """
    Looks inside the indexing screen for the active PDF / document and captures raw bytes.
    Returns (bytes, source_description) or (None, None).
    """
    # 1. Look for iframes or embeds with PDF sources
    pdf_url = None
    try:
        embeds = page.locator("embed, object, iframe")
        for i in range(embeds.count()):
            el = embeds.nth(i)
            orig = el.get_attribute("original-url")
            if orig:
                pdf_url = orig
                break
            src = el.get_attribute("src") or el.get_attribute("data") or ""
            if "showimage.php" in src.lower() or "getimage" in src.lower() or ".pdf" in src.lower():
                pdf_url = src
                break

        if not pdf_url:
            for frame in page.frames:
                f_url = frame.url
                if "showimage.php" in f_url.lower() or ".pdf" in f_url.lower() or "getimage" in f_url.lower():
                    pdf_url = f_url
                    break
    except Exception as e:
        print(f"    Error scanning document elements: {e}")

    # 2. If URL is found, download via browser context request
    if pdf_url:
        try:
            from urllib.parse import urljoin
            abs_url = urljoin(page.url, pdf_url)
            res = page.context.request.get(abs_url)
            if res.status == 200:
                body = res.body()
                if body.startswith(b"%PDF"):
                    return body, f"stream: {abs_url[:50]}"
        except Exception as e:
            print(f"    Failed direct stream fetch: {e}")

    # 3. Check for Download button on screen
    try:
        dl_btn = page.locator('a[data-qtip*="Download"], button[data-qtip*="Download"], a:has-text("Download"), button:has-text("Download"), .x-btn-inner:has-text("Download")').first
        if dl_btn.is_visible(timeout=2000):
            with page.expect_download(timeout=8000) as download_info:
                dl_btn.click()
            download = download_info.value
            temp_path = HERE / f"temp_{wanted_name}.pdf"
            download.save_as(str(temp_path))
            with open(temp_path, "rb") as f:
                b = f.read()
            temp_path.unlink(missing_ok=True)
            if b.startswith(b"%PDF"):
                return b, "download button"
    except Exception:
        pass

    return None, None

def main():
    load_dotenv(PROJECT / ".env")
    load_dotenv(HERE / ".env")
    url = os.getenv("BOTTOMLINE_URL", "").strip()
    username = os.getenv("BOTTOMLINE_USERNAME", "").strip()
    password = os.getenv("BOTTOMLINE_PASSWORD", "").strip()
    ollama_host = os.getenv("OLLAMA_HOST", "http://192.168.175.74:11434").strip()

    if not url:
        sys.exit("STOPPED: .env is missing BOTTOMLINE_URL.")

    if not PDF_DIR.exists():
        sys.exit(f"STOPPED: Target PDF directory does not exist: {PDF_DIR}")

    target_count = 5
    if len(sys.argv) > 1:
        try:
            target_count = int(sys.argv[1])
        except ValueError:
            pass

    print("\n" + "=" * 65)
    print(f"  BOTTOMLINE NEXUS - QUALIFIED INVOICE DOWNLOADER (Target: {target_count})")
    print("  Allowed Types: 'Standard Invoices' & 'Utility Invoices' ONLY")
    print(f"  Classifier Host: {ollama_host}")
    print("=" * 65 + "\n")

    # Load selectors
    selectors = json.loads(SELECTORS_FILE.read_text(encoding="utf-8"))

    # Load existing downloaded PDFs
    existing_pdfs = set(f.lower().replace(".pdf", "") for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf"))
    print(f"Current PDFs in Z drive: {len(existing_pdfs)}")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        if SESSION_FILE.exists():
            context = browser.new_context(storage_state=str(SESSION_FILE))
            print("Reusing existing saved session.")
        else:
            context = browser.new_context()

        # Intercept PDF requests in the background
        captured_pdfs = {}
        def handle_response(response):
            try:
                ct = response.headers.get("content-type", "").lower()
                r_url = response.url.lower()
                if "application/pdf" in ct or ".pdf" in r_url or "getimage" in r_url or "showimage" in r_url:
                    if response.status == 200:
                        body = response.body()
                        if body and body.startswith(b"%PDF"):
                            captured_pdfs[response.url] = body
            except Exception:
                pass

        page = context.new_page()
        page.on("response", handle_response)

        print(f"Navigating to {url}...")
        page.goto(url)
        page.wait_for_timeout(2000)

        if page.locator("input[type='password']").count() > 0:
            print("Logging in...")
            try:
                page.locator("input[type='text'], input[type='email']").first.fill(username)
                page.locator("input[type='password']").first.fill(password)
                page.keyboard.press("Enter")
            except Exception as e:
                print(f"Login error: {e}")

        print("\n>>> Please complete login/2FA if prompted.")
        input("Press ENTER once you are on the portal (queue or dashboard)... ")

        try:
            context.storage_state(path=str(SESSION_FILE))
        except Exception:
            pass

        def note(msg):
            print(f"  [Queue] {msg}")

        note("Ensuring we are on 'Images to be Indexed'...")
        return_to_queue(page, selectors, note)
        wait_for_queue_rows(page, selectors, note)

        # Read queue
        queue_index, _ = read_queue(page, selectors, note)
        print(f"\nDiscovered {len(queue_index)} invoices in the queue.")

        # Candidate selection
        candidates = []
        for name, page_num in queue_index.items():
            if normalise_name(name) not in existing_pdfs:
                candidates.append((name, page_num))

        print(f"Candidate invoices not yet in Z drive: {len(candidates)}")
        if not candidates:
            print("No new candidate invoices found in queue.")
            browser.close()
            return

        downloaded = 0
        for name, page_num in candidates:
            if downloaded >= target_count:
                break

            print("\n" + "-" * 55)
            print(f"Processing candidate ({downloaded + 1}/{target_count}): {name}")
            
            # Navigate to page and open invoice
            captured_pdfs.clear()
            go_to_queue_page(page, selectors, page_num)
            
            # Find row and click Index Selected
            rows = page.locator(selectors["queue"]["grid_row"])
            target = normalise_name(name)
            found = False
            for position in range(rows.count()):
                row = rows.nth(position)
                cells = row.locator("td")
                r_text = ""
                for c_idx in range(cells.count()):
                    try:
                        t = cells.nth(c_idx).inner_text(timeout=1000).strip()
                        if t:
                            r_text = t
                            break
                    except Exception:
                        continue
                if normalise_name(r_text) == target:
                    tick = row.locator(selectors["queue"]["row_tick_box"]).first
                    tick.click()
                    page.wait_for_timeout(400)
                    page.locator(selectors["queue"]["index_selected"]).first.click()
                    found = True
                    break

            if not found:
                print(f"  Could not find row for {name} on page {page_num}.")
                continue

            # Wait for invoice indexing screen to settle
            page.wait_for_timeout(4000)

            # Check if network response caught the PDF
            pdf_bytes = None
            if captured_pdfs:
                for pdf_url, b in captured_pdfs.items():
                    pdf_bytes = b
                    break

            if not pdf_bytes:
                pdf_bytes, _ = inspect_and_download_invoice(page, name)

            if not pdf_bytes:
                print(f"  ⚠️ Could not download or inspect PDF for {name}")
                return_to_queue(page, selectors, note)
                page.wait_for_timeout(2000)
                continue

            # Verify document type: only standard invoices & utility invoices
            print("  Inspecting document type with vision AI...")
            doc_type, reasoning = classify_doc_type(pdf_bytes, host=ollama_host)
            print(f"  Classification: [{doc_type}] ({reasoning})")

            if doc_type in ALLOWED_DOC_TYPES:
                target_file = PDF_DIR / f"{name}.pdf"
                with open(target_file, "wb") as f:
                    f.write(pdf_bytes)
                print(f"  ✓ SAVED [{doc_type}]: {target_file.name} ({len(pdf_bytes)} bytes)")
                downloaded += 1
            else:
                print(f"  ⏩ SKIPPED [{doc_type}]: Not a standard or utility invoice ({name})")

            # Return to queue for next invoice
            return_to_queue(page, selectors, note)
            page.wait_for_timeout(2000)

        print("\n" + "=" * 65)
        print(f"  FINISHED: Downloaded {downloaded} qualified invoices to {PDF_DIR}")
        print("=" * 65 + "\n")
        browser.close()

if __name__ == "__main__":
    main()
