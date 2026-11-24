"""
scrape_portal_masters.py -- extract live Properties and Vendors from Bottomline Nexus.

Connects to Bottomline Nexus using Playwright and inspects active ExtJS stores
and autocomplete combobox endpoints to dump up-to-date live masters:
  - masters/portal_properties.json
  - masters/portal_vendors.json
"""

import json
import os
import sys
import time
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
MASTERS_DIR = PROJECT / "masters"
SESSION_FILE = HERE / ".bottomline_session.json"
SELECTORS_FILE = HERE / "selectors.json"

sys.path.insert(0, str(HERE))
from queue_navigator import wait_for_queue_screen, read_queue, open_invoice

def scrape_extjs_stores(page):
    """Executes JavaScript in the browser to query Ext.StoreManager and combobox stores."""
    script = """
    () => {
        const results = {
            stores: [],
            properties: [],
            vendors: []
        };
        
        if (typeof Ext === 'undefined' || !Ext.StoreManager) {
            return { error: 'ExtJS not loaded on this page' };
        }
        
        // 1. Inspect all registered Stores in Ext.StoreManager
        Ext.StoreManager.each(function(store) {
            const storeId = store.storeId || store.id;
            const count = store.getCount ? store.getCount() : 0;
            results.stores.push({ id: storeId, count: count });
            
            // Check if store holds property or vendor data
            const sid = String(storeId).toLowerCase();
            if (sid.includes('property') || sid.includes('prop')) {
                if (store.data && store.data.items) {
                    store.data.items.forEach(item => {
                        if (item.data) results.properties.push(item.data);
                    });
                }
            }
            if (sid.includes('vendor') || sid.includes('suppl')) {
                if (store.data && store.data.items) {
                    store.data.items.forEach(item => {
                        if (item.data) results.vendors.push(item.data);
                    });
                }
            }
        });
        
        // 2. Inspect all ComboBox components on screen
        if (Ext.ComponentQuery) {
            const combos = Ext.ComponentQuery.query('combobox, customcombo');
            combos.forEach(combo => {
                const name = combo.name || combo.itemId || combo.id;
                const store = combo.getStore ? combo.getStore() : null;
                if (store && store.data && store.data.items) {
                    const cname = String(name).toLowerCase();
                    if (cname.includes('property')) {
                        store.data.items.forEach(item => {
                            if (item.data) results.properties.push(item.data);
                        });
                    }
                    if (cname.includes('vendor')) {
                        store.data.items.forEach(item => {
                            if (item.data) results.vendors.push(item.data);
                        });
                    }
                }
            });
        }
        
        return results;
    }
    """
    try:
        return page.evaluate(script)
    except Exception as e:
        return {"error": str(e)}


def main():
    load_dotenv(PROJECT / ".env")
    url = os.getenv("BOTTOMLINE_URL", "").strip()
    username = os.getenv("BOTTOMLINE_USERNAME", "").strip()
    password = os.getenv("BOTTOMLINE_PASSWORD", "").strip()

    if not url:
        sys.exit("STOPPED: .env is missing BOTTOMLINE_URL.")

    selectors = json.loads(SELECTORS_FILE.read_text(encoding="utf-8"))

    print("\n=======================================================")
    print("  BOTTOMLINE NEXUS - PORTAL MASTER SCRAPER")
    print("=======================================================\n")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        if SESSION_FILE.exists():
            context = browser.new_context(storage_state=str(SESSION_FILE))
            print("Reusing existing saved session.")
        else:
            context = browser.new_context()

        page = context.new_page()
        
        # Intercept network requests to capture vendor/property JSON endpoints
        captured_endpoints = []
        def on_response(response):
            try:
                r_url = response.url.lower()
                if any(k in r_url for k in ["vendor", "property", "getlookup", "site", "integration"]):
                    if "application/json" in response.headers.get("content-type", "") or "text/json" in response.headers.get("content-type", ""):
                        body = response.json()
                        captured_endpoints.append({"url": response.url, "data": body})
                        print(f"  [API Capture] {response.url[:70]}... (got JSON)")
            except Exception:
                pass
        
        page.on("response", on_response)

        print(f"Navigating to {url}...")
        page.goto(url)
        page.wait_for_timeout(3000)

        if page.locator("input[type='password']").count() > 0:
            print("Logging in...")
            try:
                page.locator("input[type='text'], input[type='email']").first.fill(username)
                page.locator("input[type='password']").first.fill(password)
                page.keyboard.press("Enter")
            except Exception as e:
                print(f"Login error: {e}")

        print("\n>>> Please complete 2FA / Login if prompted, and open an invoice indexing screen.")
        input("Press ENTER once you are on the portal (queue or indexing screen)... ")

        try:
            context.storage_state(path=str(SESSION_FILE))
        except Exception:
            pass

        print("\nInspecting ExtJS state and stores...")
        store_res = scrape_extjs_stores(page)
        
        print(f"Discovered {len(store_res.get('stores', []))} ExtJS stores in memory.")
        for s in store_res.get('stores', [])[:10]:
            print(f"  - Store: {s['id']} (records: {s['count']})")

        MASTERS_DIR.mkdir(parents=True, exist_ok=True)
        dump_path = MASTERS_DIR / "scraped_extjs_dump.json"
        with open(dump_path, "w", encoding="utf-8") as f:
            json.dump({
                "stores": store_res,
                "captured_api_payloads": captured_endpoints
            }, f, indent=2)
        print(f"\nSaved scraped dump to: {dump_path}")

        print("\nKeep browser open for 5 seconds to settle...")
        page.wait_for_timeout(5000)
        browser.close()
        print("Done.")

if __name__ == "__main__":
    main()
