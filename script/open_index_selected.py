"""Open the portal and run the working navigation steps ONLY:

    1. Log in (credentials from .env)
    2. Image Management -> Images to be Indexed
    3. Records per page: 200
    4. Go to the LAST page
    5. Select all (Image Name header checkbox)
    6. Click "Index Selected"  -> the first image opens in the indexing form

It fills nothing and never presses Save / Index / Convert / Delete.
The browser stays open so you can work in it.

Run:
    C:\\res1_venv\\Scripts\\python.exe script\\open_index_selected.py
"""
import json
import os
import sys
import time
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from queue_navigator import return_to_queue  # noqa: E402

load_dotenv(HERE.parent / ".env")
SEL = json.loads((HERE / "selectors.json").read_text(encoding="utf-8"))
USER = os.getenv("BOTTOMLINE_USERNAME")
PASSWORD = os.getenv("BOTTOMLINE_PASSWORD")
URL = os.getenv("BOTTOMLINE_URL")
STAY_OPEN_SECONDS = 1200


def log(msg):
    print("[portal]", msg, flush=True)


def main():
    if not (URL and USER and PASSWORD):
        sys.exit("Set BOTTOMLINE_URL, BOTTOMLINE_USERNAME, BOTTOMLINE_PASSWORD in .env")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        pg = browser.new_context(viewport={"width": 1500, "height": 900}).new_page()

        # 1) log in
        pg.goto(URL, timeout=60000)
        pg.fill('input[type="text"]', USER)
        pg.fill('input[type="password"]', PASSWORD)
        pg.keyboard.press("Enter")
        pg.locator(SEL["queue"]["menu_opener"] + ':has-text("Image Management")'
                   ).first.wait_for(state="visible", timeout=60000)

        # 2) Image Management -> Images to be Indexed
        if not return_to_queue(pg, SEL, log):
            sys.exit("Could not open 'Images to be Indexed'.")
        pg.wait_for_timeout(4000)

        # 3) Records per page: 200
        boxes = pg.locator('input[id^="combo-"][id$="-inputEl"]')
        for i in range(boxes.count()):
            box = boxes.nth(i)
            if box.is_visible() and box.input_value().strip().isdigit():
                box.click()
                pg.wait_for_timeout(800)
                pg.locator(SEL["dropdown_item_selector"],
                           has_text="200").first.click()
                break
        pg.wait_for_timeout(5000)
        log("Records per page: 200")

        # 4) last page
        pg.locator(".x-tbar-page-last").first.click()
        pg.wait_for_timeout(5000)
        log("Last page: " +
            pg.locator(SEL["queue"]["page_number_box"]).first.input_value())

        # 5) select all (Image Name header checkbox)
        pg.locator(".x-column-header-checkbox, .x-column-header-first"
                   ).first.click()
        pg.wait_for_timeout(2000)
        log(f"Selected rows: {pg.locator('.x-grid-item-selected').count()}")

        # 6) Index Selected (never the Delete button)
        pg.locator(SEL["queue"]["index_selected"]).first.click()
        pg.wait_for_timeout(5000)
        log("Index Selected clicked -- the first image is open.")

        time.sleep(STAY_OPEN_SECONDS)
        browser.close()


if __name__ == "__main__":
    main()
