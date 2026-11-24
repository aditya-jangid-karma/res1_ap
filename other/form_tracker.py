"""
Maker Tracking Form Submitter
Submits recorded invoice details directly to Google Forms in the background.

Google Form:
https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform
"""

import urllib.request
import urllib.parse
import sys

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/formResponse"

ENTRY_IMAGE_NAME = "entry.253702755"
ENTRY_IMAGE_ACTION = "entry.408781098"
ENTRY_INVOICE_NUMBER = "entry.1032666432"
ENTRY_REASON_DELETE = "entry.106333593"
ENTRY_IMAGE_TYPE = "entry.205748566"


def determine_image_type(package: str, is_utility: bool) -> str:
    """
    Returns one of the 4 exact options expected by the form:
      - 'Intacct - Expense'
      - 'Intacct - Utility'
      - 'Yardi - Expense'
      - 'Yardi - Utility'
    """
    pkg = (package or "").strip().upper()
    if "INTACCT" in pkg or "SAGE" in pkg:
        prefix = "Intacct"
    else:
        prefix = "Yardi"

    suffix = "Utility" if is_utility else "Expense"
    return f"{prefix} - {suffix}"


def submit_maker_tracker(
    image_name: str,
    invoice_number: str,
    package: str = "YARDI",
    is_utility: bool = False,
    timeout: int = 10,
    dry_run: bool = False
) -> dict:
    """
    Submits an Approved invoice record to the Maker Tracking form.

    Returns:
      {"success": True/False, "status_code": int, "image_type": str, "message": str}
    """
    image_name = str(image_name or "").strip()
    invoice_number = str(invoice_number or "").strip()
    image_type = determine_image_type(package, is_utility)

    if not image_name:
        return {"success": False, "status_code": 0, "image_type": image_type, "message": "Missing image_name"}

    payload = {
        ENTRY_IMAGE_NAME: image_name,
        ENTRY_IMAGE_ACTION: "Approve",
        ENTRY_INVOICE_NUMBER: invoice_number,
        ENTRY_IMAGE_TYPE: image_type,
        "pageHistory": "0,1,6",
    }

    if dry_run:
        return {
            "success": True,
            "status_code": 200,
            "image_type": image_type,
            "message": f"[DRY RUN] Would submit: image='{image_name}', action='Approve', inv='{invoice_number}', type='{image_type}'",
        }

    encoded_data = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(
        FORM_URL,
        data=encoded_data,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            status = response.getcode()
            if status == 200:
                return {
                    "success": True,
                    "status_code": status,
                    "image_type": image_type,
                    "message": f"Submitted successfully: image='{image_name}', inv='{invoice_number}', type='{image_type}'",
                }
            else:
                return {
                    "success": False,
                    "status_code": status,
                    "image_type": image_type,
                    "message": f"Unexpected HTTP status {status}",
                }
    except Exception as e:
        return {
            "success": False,
            "status_code": 0,
            "image_type": image_type,
            "message": f"Submission failed: {e}",
        }


def get_prefilled_url(
    image_name: str,
    invoice_number: str,
    package: str = "YARDI",
    is_utility: bool = False,
    action: str = "Approve",
    delete_reason: str = None
) -> str:
    """Returns a direct Google Form URL prefilled with image_name, action, invoice_number, image_type, and optional delete reason."""
    image_name = str(image_name or "").strip()
    invoice_number = str(invoice_number or "").strip()
    image_type = determine_image_type(package, is_utility)
    params = {
        "usp": "pp_url",
        ENTRY_IMAGE_NAME: image_name,
        ENTRY_IMAGE_ACTION: action,
        ENTRY_INVOICE_NUMBER: invoice_number,
    }
    if action == "Delete" and delete_reason:
        params[ENTRY_REASON_DELETE] = delete_reason
    else:
        params[ENTRY_IMAGE_TYPE] = image_type

    base = "https://docs.google.com/forms/d/e/1FAIpQLScIqyCFFiE4XSYCxGsMWdpkzQkidzsnX8Yzk-4b4CeOpo7mWQ/viewform"
    return f"{base}?{urllib.parse.urlencode(params)}"


if __name__ == "__main__":
    # Test script in dry-run mode
    res = submit_maker_tracker(
        image_name="test-uuid-12345",
        invoice_number="55029030362-20260928",
        package="YARDI",
        is_utility=True,
        dry_run=True,
    )
    print("Self-test result:")
    print(res)
