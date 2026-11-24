import os
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
CHROME_USER_DATA = Path(r"C:\Users\Meet Patel\AppData\Local\Google\Chrome\User Data")
AUTOMATION_PROFILE = HERE / ".chrome_user_profile"
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform"

def launch_with_chrome_profile():
    print("=" * 60)
    print("Launching Chrome directly with your 'Default' Profile:")
    print("meet.patel@kgsoffice.com")
    print("=" * 60)
    
    # We use Playwright's persistent context pointing directly to Chrome's executable
    # with your real profile directory
    real_user_data = r"C:\Users\Meet Patel\AppData\Local\Google\Chrome\User Data"
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=real_user_data,
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            headless=False,
            ignore_default_args=["--enable-automation", "--no-sandbox", "--disable-blink-features=AutomationControlled"],
            args=[
                "--profile-directory=Default",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled"
            ]
        )
        page = context.pages[0] if context.pages else context.new_page()
        # Remove navigator.webdriver flag so Google allows sign-in
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        """)
        page.goto(FORM_URL)
        
        print("\n>>> Chrome has opened with your real profile.")
        print(">>> If asked to sign in, do it once right here in your profile.")
        input("\nPress ENTER here once you are logged in and see the form... ")
        
        # Save storage state for background headlessly
        session_file = HERE / ".google_session.json"
        context.storage_state(path=str(session_file))
        print(f"\n✓ Session state saved to: {session_file.name}")
        context.close()

if __name__ == "__main__":
    launch_with_chrome_profile()
