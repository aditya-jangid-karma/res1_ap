from playwright.sync_api import sync_playwright
import os
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
BROWSERS_DIR = PROJECT / ".browsers"
if BROWSERS_DIR.exists():
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(BROWSERS_DIR)

form_url = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto(form_url)
    page.wait_for_timeout(2000)
    
    # Check dialog / overlay
    modal = page.locator('div[role="dialog"], .llhEMd, .iWO5td')
    print("Modal count:", modal.count())
    if modal.count() > 0:
        for i in range(modal.count()):
            m = modal.nth(i)
            print(f"Modal {i}: text={m.inner_text()[:150]}")
            btns = m.locator('button, div[role="button"]').all()
            for b in btns:
                print(f"  Button: '{b.inner_text()}'")
    browser.close()
