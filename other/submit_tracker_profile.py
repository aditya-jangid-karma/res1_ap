import os
import shutil
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

_HOME = os.path.expanduser("~")
REAL_USER_DATA_DIR = os.path.join(_HOME, "AppData", "Local", "Google", "Chrome", "User Data")
REAL_PROFILE_DIR = "Default"
AUTOMATION_USER_DATA_DIR = os.path.join(_HOME, "ChromeAutomationProfile")
AUTOMATION_PROFILE_DIR = "Default"
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform"

def sync_profile_copy():
    print("[SYSTEM] Syncing browser profile from local Chrome (exact HelloGM / Harleysville pattern)...")
    src = os.path.join(REAL_USER_DATA_DIR, REAL_PROFILE_DIR)
    dst = os.path.join(AUTOMATION_USER_DATA_DIR, AUTOMATION_PROFILE_DIR)
    os.makedirs(AUTOMATION_USER_DATA_DIR, exist_ok=True)

    local_state_src = os.path.join(REAL_USER_DATA_DIR, "Local State")
    local_state_dst = os.path.join(AUTOMATION_USER_DATA_DIR, "Local State")
    if os.path.exists(local_state_src):
        try:
            shutil.copy2(local_state_src, local_state_dst)
        except Exception:
            pass

    SKIP_NAMES = {
        "Cache", "Code Cache", "GPUCache", "ShaderCache", "GrShaderCache",
        "DawnCache", "DawnGraphiteCache", "DawnWebGPUCache", "Media Cache",
        "Service Worker", "Crashpad", "Component Cache", "extensions_crx_cache",
        "optimization_guide_model_store", "lockfile",
    }

    def ignore_locks(directory, contents):
        return [c for c in contents if c.startswith("Singleton") or c in SKIP_NAMES]

    def safe_copy(src_file, dst_file):
        try:
            shutil.copy2(src_file, dst_file)
        except Exception:
            pass

    try:
        shutil.copytree(src, dst, ignore=ignore_locks, dirs_exist_ok=True, copy_function=safe_copy)
        print("[SYSTEM] Chrome profile sync complete (live locked files safely bypassed).")
    except Exception as e:
        print(f"[SYSTEM WARNING] Partial profile copy warning (non-fatal): {e}")

def submit_tracker(image_name, invoice_number, image_type="Intacct - Expense"):
    sync_profile_copy()
    
    print(f"\n[Maker Tracker] Submitting for: {image_name} | {invoice_number} | {image_type}")
    
    with sync_playwright() as p:
        launch_kwargs = dict(
            user_data_dir=AUTOMATION_USER_DATA_DIR,
            headless=False,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
                "--start-maximized"
            ]
        )
        try:
            context = p.chromium.launch_persistent_context(channel="chrome", **launch_kwargs)
        except Exception:
            context = p.chromium.launch_persistent_context(**launch_kwargs)
            
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(FORM_URL)
        page.wait_for_timeout(2500)
        
        # 1. Fill Image Name
        # Find active question 1 text box
        q0 = page.locator('div[role="listitem"]').first
        # If there's an enabled input
        inp_img = q0.locator('input[type="text"]:visible, input:not([disabled])').first
        if inp_img.is_visible():
            inp_img.fill(image_name)
        else:
            q0.click()
            page.keyboard.type(image_name)
        print("  -> Filled Image Name:", image_name)
        
        # 2. Select Approve
        page.locator('div[role="radio"][data-value="Approve"], span:has-text("Approve")').first.click()
        print("  -> Selected 'Approve'")
        
        # 3. Click Next
        page.locator('div[role="button"]:has-text("Next"), span:has-text("Next")').first.click()
        page.wait_for_timeout(1200)
        
        # 4. Fill Invoice Number
        q_inv = page.locator('div[role="listitem"]').first
        inp_inv = q_inv.locator('input[type="text"]:visible, input:not([disabled])').first
        if inp_inv.is_visible():
            inp_inv.fill(invoice_number)
        else:
            q_inv.click()
            page.keyboard.type(invoice_number)
        print("  -> Filled Invoice Number:", invoice_number)
        
        # 5. Click Next
        page.locator('div[role="button"]:has-text("Next"), span:has-text("Next")').first.click()
        page.wait_for_timeout(1200)
        
        # 6. Select Image Type
        dropdown = page.locator('div[role="listbox"]').first
        if dropdown.is_visible(timeout=2000):
            dropdown.click()
            page.wait_for_timeout(500)
            page.locator(f'div[role="option"]:has-text("{image_type}"), span:has-text("{image_type}")').first.click()
        else:
            page.locator(f'div[role="radio"][data-value="{image_type}"]').first.click()
        print("  -> Selected Image Type:", image_type)
        page.wait_for_timeout(500)
        
        # 7. Submit
        page.locator('div[role="button"]:has-text("Submit"), span:has-text("Submit")').first.click()
        page.wait_for_timeout(2500)
        
        content = page.content()
        success = "Your response has been recorded" in content
        
        # Log to file
        log_entry = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Image: {image_name} | Inv: {invoice_number} | Type: {image_type} | Result: {'SUCCESS' if success else 'FAILED'}\n"
        log_path = Path(HERE / "maker_tracker_submissions.log")
        with open(log_path, "a", encoding="utf-8") as lf:
            lf.write(log_entry)
            
        context.close()
        
        if success:
            print(f"✓ [Maker Tracker SUCCESS] Your response has been recorded for {image_name}")
        else:
            print(f"⚠️ [Maker Tracker Notice] Check confirmation page.")
        return success

if __name__ == "__main__":
    # Test submission on the first of your 3 recently recorded invoices (Excel Line 255)
    submit_tracker(
        image_name="3a6424da-73f9-4361-8059-716b0e8392c9",
        invoice_number="90289216",
        image_type="Intacct - Expense"
    )
