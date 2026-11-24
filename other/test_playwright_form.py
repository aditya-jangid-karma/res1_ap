import os
from pathlib import Path
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
BROWSERS_DIR = PROJECT / ".browsers"
if BROWSERS_DIR.exists():
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(BROWSERS_DIR)

form_url = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform"

def submit_via_playwright(image_name, invoice_number, image_type="Yardi - Utility"):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.goto(form_url)
        page.wait_for_load_state("domcontentloaded")
        
        # 1. Fill Image Name: targets the input inside the question
        img_input = page.locator('div[data-params*="Image Name"] input[type="text"], input[aria-labelledby*="i1"], div[role="listitem"] input[type="text"]').first
        img_input.fill(image_name)
        print("Filled Image Name:", image_name)
        
        # 2. Click 'Approve'
        approve_radio = page.locator('div[role="radio"][data-value="Approve"]').first
        approve_radio.click()
        print("Selected Approve")
        
        # 3. Click 'Next'
        next_btn = page.locator('div[role="button"]:has-text("Next"), span:has-text("Next")').first
        next_btn.click()
        page.wait_for_timeout(1500)
        
        # 4. Fill 'Write invoice number'
        inv_input = page.locator('div[role="listitem"] input[type="text"], input[type="text"]:visible').first
        inv_input.fill(invoice_number)
        print("Filled Invoice Number:", invoice_number)
        
        # 5. Click 'Next'
        next_btn = page.locator('div[role="button"]:has-text("Next"), span:has-text("Next")').first
        next_btn.click()
        page.wait_for_timeout(1500)
        
        # 6. Select Image Type dropdown
        dropdown = page.locator('div[role="listbox"]').first
        if dropdown.is_visible(timeout=3000):
            dropdown.click()
            page.wait_for_timeout(500)
            option = page.locator(f'div[role="option"]:has-text("{image_type}"), span:has-text("{image_type}")').first
            option.click()
            print("Selected Image Type from dropdown:", image_type)
        else:
            page.locator(f'div[role="radio"][data-value="{image_type}"]').first.click()
            print("Selected Image Type radio:", image_type)
            
        page.wait_for_timeout(500)
        # 7. Click 'Submit'
        submit_btn = page.locator('div[role="button"]:has-text("Submit"), span:has-text("Submit")').first
        submit_btn.click()
        page.wait_for_timeout(2500)
        
        # Check confirmation
        content = page.content()
        is_success = "Your response has been recorded" in content
        print("Playwright submit result:", "SUCCESS - Your response has been recorded!" if is_success else "Check page content")
        browser.close()
        return is_success

if __name__ == "__main__":
    submit_via_playwright("58a8e909-27ca-480c-a784-16654d20b19d", "200262000466", "Yardi - Utility")
