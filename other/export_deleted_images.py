"""
export_deleted_images.py -- Scrape deleted images grid from Bottomline Nexus
and download/read image details for the last 2 months.
"""

import os
import sys
import json
import time
import re
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
OUTPUT_EXCEL = HERE / "deleted_invoices_last_2_months.xlsx"
SESSION_FILE = HERE / ".bottomline_session.json"

def main():
    load_dotenv(PROJECT / ".env")
    url = os.getenv("BOTTOMLINE_URL", "").strip()
    username = os.getenv("BOTTOMLINE_USERNAME", "").strip()
    password = os.getenv("BOTTOMLINE_PASSWORD", "").strip()

    print("\n=======================================================")
    print("  BOTTOMLINE NEXUS - DELETED INVOICES EXTRACTOR (LAST 2 MONTHS)")
    print("=======================================================\n")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        if SESSION_FILE.exists():
            context = browser.new_context(storage_state=str(SESSION_FILE))
            print("Reusing existing saved session.")
        else:
            context = browser.new_context()

        page = context.new_page()

        # Listen for internal API calls containing grid data
        captured_grid_data = []
        def on_response(response):
            try:
                r_url = response.url.lower()
                if "getimage" in r_url or "deleted" in r_url or "imagestore" in r_url or "grid" in r_url or "list" in r_url:
                    if "json" in response.headers.get("content-type", "").lower():
                        data = response.json()
                        if isinstance(data, dict) and ("data" in data or "records" in data or "rows" in data or "total" in data):
                            captured_grid_data.append(data)
                            print(f"  [API Capture] Captured grid JSON response from {response.url[:60]}...")
            except Exception:
                pass

        page.on("response", on_response)

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

        print("\n>>> Please complete login/2FA if prompted, and ensure the 'Deleted Images' tab is open.")
        input("Press ENTER once you are on the 'Deleted Images' grid... ")

        try:
            context.storage_state(path=str(SESSION_FILE))
        except Exception:
            pass

        # 1. Inspect grid via ExtJS store in memory
        print("\nInspecting ExtJS in memory for the active Deleted Images store...")
        script = """
        () => {
            const results = [];
            if (typeof Ext === 'undefined') return { error: 'No ExtJS' };

            // Find all Ext grid panels or stores
            const grids = Ext.ComponentQuery.query('gridpanel, grid');
            for (let g of grids) {
                if (g.isVisible && g.isVisible()) {
                    const store = g.getStore ? g.getStore() : null;
                    if (store && store.data && store.data.items && store.data.items.length > 0) {
                        const items = store.data.items.map(it => it.data);
                        return { found: true, id: g.id, count: items.length, records: items };
                    }
                }
            }

            // Fallback: search StoreManager
            if (Ext.StoreManager) {
                let bestStore = null;
                Ext.StoreManager.each(st => {
                    const sid = String(st.storeId || st.id).toLowerCase();
                    if (sid.includes('image') && st.getCount && st.getCount() > 0) {
                        bestStore = st;
                    }
                });
                if (bestStore) {
                    const items = bestStore.data.items.map(it => it.data);
                    return { found: true, id: bestStore.storeId, count: items.length, records: items };
                }
            }
            return { found: false };
        }
        """
        eval_res = page.evaluate(script)
        records = []
        if eval_res.get("found"):
            print(f"✓ Discovered ExtJS store '{eval_res.get('id')}' with {eval_res.get('count')} records on screen!")
            records = eval_res.get("records", [])
        else:
            print("Notice: Could not read direct ExtJS store. Attempting Export / DOM scan...")

        # 2. Try clicking the built-in Export All / Excel button to get the complete dataset
        export_clicked = False
        try:
            # Check Export All if present
            export_all_cb = page.locator('label:has-text("Export All"), input[type="checkbox"]:has-text("Export All")').first
            if export_all_cb.is_visible(timeout=2000):
                export_all_cb.click()
                print("  Checked 'Export All'...")
                page.wait_for_timeout(500)

            # Target the green Excel icon on the paging toolbar
            excel_btn = page.locator('a[data-qtip*="Excel"]:visible, .x-btn:has([class*="excel"]):visible, .x-btn:has([class*="page-excel"]):visible, a.x-btn:has(span[style*="excel"]):visible').first
            if excel_btn.is_visible(timeout=3000):
                print(f"  -> Triggering Excel Export button...")
                with page.expect_download(timeout=15000) as dl_info:
                    excel_btn.click()
                dl = dl_info.value
                raw_dl_path = HERE / "nexus_deleted_export_raw.xlsx"
                dl.save_as(str(raw_dl_path))
                print(f"  ✓ Downloaded complete export from Nexus: {raw_dl_path.name}")
                export_clicked = True
        except Exception as e:
            print(f"  Notice: Direct export download attempt ({e})")

        # If we have the ExtJS records or captured records
        if records:
            df = pd.DataFrame(records)
            print(f"\nDiscovered {len(df)} total records from store.")

            # Filter for last 2 months (62 days) using image_index_deleted_datetm
            cutoff_date = datetime.now() - timedelta(days=62)
            date_col = 'image_index_deleted_datetm' if 'image_index_deleted_datetm' in df.columns else next((c for c in df.columns if "deleted" in c.lower() and "date" in c.lower()), None)

            if date_col:
                try:
                    # Convert to datetime and strip timezone
                    df['parsed_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.tz_localize(None)
                    df_filtered = df[df['parsed_date'] >= cutoff_date]
                    if not df_filtered.empty:
                        df = df_filtered
                    print(f"Filtered to {len(df)} records for the last 2 months (since {cutoff_date.strftime('%Y-%m-%d')}).")
                except Exception as de:
                    print(f"Date filter notice: {de}")

            # Structure desired clean columns
            clean_columns = {
                'Image_Index_Name': 'Image Name',
                'image_doctype_name': 'Invoice Type',
                'Image_Index_Ref': 'Invoice Number',
                'Image_Index_Amount': 'Amount',
                'vendor_name': 'Vendor Name',
                'vendortype_name': 'Vendor Type',
                'property_name': 'Property Name',
                'Property_Id': 'Property ID',
                'integration_package_name': 'Integration Package',
                'image_index_deleted_datetm': 'Deleted Date',
                'deletedby_username': 'Deleted By',
                'Image_Index_Invoice_Date': 'Invoice Date',
                'Image_Index_Due_Date': 'Due Date',
                'Image_Index_Exception_reason': 'Exception Reason'
            }

            # Select available columns
            export_cols = [c for c in clean_columns.keys() if c in df.columns]
            export_df = df[export_cols].copy()
            export_df.rename(columns=clean_columns, inplace=True)

            # Strip any remaining timezone info across all columns
            for col in export_df.columns:
                if pd.api.types.is_datetime64_any_dtype(export_df[col]):
                    export_df[col] = export_df[col].dt.tz_localize(None)
                elif export_df[col].dtype == 'object':
                    # Handle any stringified dates or complex types
                    pass

            export_df.to_excel(OUTPUT_EXCEL, index=False)
            print(f"\n=======================================================")
            print(f"  ✓ SUCCESS: Saved {len(export_df)} deleted invoice records to:")
            print(f"    {OUTPUT_EXCEL}")
            print(f"=======================================================\n")

        print("\nProcess finished.")
        browser.close()

if __name__ == "__main__":
    main()
