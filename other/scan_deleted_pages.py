"""
scan_deleted_pages.py -- Step 1 of the deleted-invoice extraction.

WHAT IT DOES
    Opens the Nexus "Deleted Images" grid, loads EVERY page of it from code
    (no clicking "Next"), and saves the full record of every deleted image to
    deleted_grid_all_pages.json. Each record includes `image_index_GUID`, the
    short code Nexus uses to serve the PDF:
        https://residentialone.nexuspayables.com/showImage.php?ID=<image_index_GUID>

WHY
    extract_deleted_invoices_with_qwen.py (Step 2) looks up each image named in
    "Deleted invoices as of 10.16.xlsx" in this file, then downloads its PDF
    directly. It no longer needs the grid on screen.

RUN
    python automation/scan_deleted_pages.py        (from the Invoice_Indexer folder)
    Log in if asked, wait for the Deleted Images grid, press ENTER. ~2-4 minutes.

READ-ONLY: it only reads the grid. It changes nothing in Nexus.
"""

import json
import math
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
SESSION_FILE = HERE / ".bottomline_session.json"
GRID_CACHE = HERE / "deleted_grid_all_pages.json"
DELETED_URL = "https://residentialone.nexuspayables.com/#Images:showMain:deleted"
TARGET_XLSX = Path(os.getenv("TARGET_XLSX", r"C:\Users\Meet Patel\Downloads\Deleted as of 10.20.xlsx"))

# Reads (and, for n >= 2, first loads) page n of the visible Ext grid's store.
LOAD_PAGE_JS = """
async (n) => {
  if (typeof Ext === 'undefined') return {error: 'ExtJS not found on this page'};
  let st = null;
  for (const g of Ext.ComponentQuery.query('gridpanel, grid')) {
    if (g.isVisible && g.isVisible()) {
      const s = g.getStore ? g.getStore() : null;
      if (s && s.data && s.data.items && s.data.items.length) { st = s; break; }
    }
  }
  if (!st) return {error: 'no visible grid with records'};
  if (n && st.currentPage !== n) {
    await new Promise(resolve => {
      const timer = setTimeout(resolve, 30000);
      st.on('load', () => { clearTimeout(timer); resolve(); }, null, {single: true});
      st.loadPage(n);
    });
  }
  return {currentPage: st.currentPage, total: st.getTotalCount(), pageSize: st.pageSize,
          records: st.data.items.map(it => it.data)};
}
"""


def open_deleted_grid(playwright):
    """Open Nexus on the Deleted Images grid (reusing the saved login session if
    there is one), let the human log in / pass 2FA, and wait for ENTER.
    Returns (browser, context, page). Shared with the extractor script."""
    load_dotenv(PROJECT / ".env")
    username = os.getenv("BOTTOMLINE_USERNAME", "").strip()
    password = os.getenv("BOTTOMLINE_PASSWORD", "").strip()

    browser = playwright.chromium.launch(headless=False)
    if SESSION_FILE.exists():
        context = browser.new_context(storage_state=str(SESSION_FILE))
        print("Reusing existing saved session.")
    else:
        context = browser.new_context()

    page = context.new_page()
    print(f"Navigating to {DELETED_URL} ...")
    page.goto(DELETED_URL)
    page.wait_for_timeout(3000)

    if page.locator("input[type='password']").count() > 0:
        print("Logging in...")
        try:
            page.locator("input[type='text'], input[type='email']").first.fill(username)
            page.locator("input[type='password']").first.fill(password)
            page.keyboard.press("Enter")
        except Exception as e:
            print(f"Login error: {e}")

    print("\n>>> Complete login/2FA if prompted and make sure the 'Deleted Images' grid is visible.")
    input("Press ENTER once you can see the Deleted Images grid... ")

    try:
        context.storage_state(path=str(SESSION_FILE))
    except Exception:
        pass
    return browser, context, page


def _first_key(records):
    r = records[0] if records else {}
    return str(r.get("image_index_GUID") or r.get("Image_Index_Name") or "")


def load_target_names():
    """Load target Image Names from TARGET_XLSX if it exists."""
    if not TARGET_XLSX.exists():
        print(f"Target Excel not found at: {TARGET_XLSX} (will scan without target early-exit)")
        return set()
    try:
        df = pd.read_excel(TARGET_XLSX)
        col = next((c for c in df.columns if str(c).strip().lower() == "image name"), None)
        if not col:
            print(f"Warning: 'Image name' column not found in {TARGET_XLSX.name}. Headers: {list(df.columns)}")
            return set()
        targets = set(str(v).strip().lower() for v in df[col].dropna() if str(v).strip())
        print(f"Loaded {len(targets)} unique target image names from {TARGET_XLSX.name}")
        return targets
    except Exception as e:
        print(f"Warning loading {TARGET_XLSX.name}: {e}")
        return set()


def collect_all_pages(page, target_names=None):
    """Walk pages of the grid. If target_names is given, stops early once all targets are found."""
    first = page.evaluate(LOAD_PAGE_JS, None)
    if first.get("error"):
        raise RuntimeError(first["error"])

    total = int(first["total"] or 0)
    page_size = int(first["pageSize"] or len(first["records"]) or 200)
    n_pages = max(1, math.ceil(total / page_size))
    print(f"Grid says {total} deleted images, {page_size} per page -> {n_pages} pages.\n")

    all_records, seen = [], set()
    found_targets = set()

    def add(records):
        added = 0
        for r in records:
            key = str(r.get("image_index_GUID") or r.get("Image_Index_Name") or "")
            name = str(r.get("Image_Index_Name") or "").strip().lower()
            if key and key not in seen:
                seen.add(key)
                all_records.append(r)
                added += 1
                if target_names and name in target_names:
                    found_targets.add(name)
        return added

    add(first["records"])
    targets_status = f" | {len(found_targets)}/{len(target_names)} targets found" if target_names else ""
    print(f"Page  1/{n_pages} | {len(first['records']):3d} records | running total {len(all_records)}{targets_status}")
    
    if target_names and len(found_targets) >= len(target_names):
        print(f"\n✓ Found all {len(target_names)} target images in page 1! Early stopping.")
        return all_records, total

    prev_key = _first_key(first["records"])

    for n in range(2, n_pages + 1):
        got = None
        for attempt in range(1, 4):
            res = page.evaluate(LOAD_PAGE_JS, n)
            recs = res.get("records") or []
            # a real page change: records exist and are not the previous page again
            if not res.get("error") and recs and _first_key(recs) != prev_key:
                got = recs
                break
            print(f"  (page {n}: not loaded properly, attempt {attempt}/3 ...)")
            page.wait_for_timeout(2000)
        if got is None:
            print(f"  !! Page {n} could not be loaded - stopping the scan here.")
            break
        add(got)
        prev_key = _first_key(got)
        targets_status = f" | {len(found_targets)}/{len(target_names)} targets found" if target_names else ""
        print(f"Page {n:2d}/{n_pages} | {len(got):3d} records | running total {len(all_records)}{targets_status}")

        if target_names and len(found_targets) >= len(target_names):
            print(f"\n✓ Found all {len(target_names)} target images by page {n}! Early stopping (no need to scan older pages).")
            break

    return all_records, total


def main():
    print("\n=======================================================")
    print(f"  SCAN DELETED IMAGES -> {GRID_CACHE.name}")
    print(f"  Target Excel: {TARGET_XLSX}")
    print("=======================================================\n")

    target_names = load_target_names()

    with sync_playwright() as playwright:
        browser, context, page = open_deleted_grid(playwright)
        try:
            records, grid_total = collect_all_pages(page, target_names=target_names)
        finally:
            browser.close()

    # Merge newly scanned records with existing cache so historical scans are never lost
    existing_map = {}
    if GRID_CACHE.exists():
        try:
            with open(GRID_CACHE, "r", encoding="utf-8") as f:
                old_data = json.load(f)
            old_records = old_data.get("records", []) if isinstance(old_data, dict) else old_data
            for r in old_records:
                k = str(r.get("image_index_GUID") or r.get("Image_Index_Name") or "")
                if k:
                    existing_map[k] = r
            print(f"Loaded {len(existing_map)} previous records from existing cache.")
        except Exception as e:
            print(f"Notice loading existing cache: {e}")

    # New records overwrite old ones
    for r in records:
        k = str(r.get("image_index_GUID") or r.get("Image_Index_Name") or "")
        if k:
            existing_map[k] = r

    merged_records = list(existing_map.values())

    print("\n" + "=" * 60)
    print(f"  Grid total reported by Nexus : {grid_total}")
    print(f"  Scanned in this session      : {len(records)}")
    print(f"  Total records in cache       : {len(merged_records)}")
    if target_names:
        cache_names = set(str(r.get("Image_Index_Name") or "").strip().lower() for r in merged_records)
        found_in_cache = target_names.intersection(cache_names)
        print(f"  Target images in cache       : {len(found_in_cache)} / {len(target_names)}")
    print("=" * 60)

    payload = {
        "scanned_at": datetime.now().isoformat(timespec="seconds"),
        "grid_total": grid_total,
        "records": merged_records,
    }
    tmp = GRID_CACHE.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1, default=str)
    os.replace(tmp, GRID_CACHE)
    print(f"\nSaved {len(merged_records)} records to: {GRID_CACHE.name}")
    print("Next: python automation/extract_deleted_invoices_v2.py --limit 5   (or run full extraction)")


if __name__ == "__main__":
    main()
