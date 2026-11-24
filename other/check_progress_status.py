import json
from pathlib import Path

p = Path("Invoice_Indexer/automation/deleted_invoices_progress.json")
if p.exists():
    with open(p, "r", encoding="utf-8") as f:
        d = json.load(f)
    print("TOTAL_RECORDS_PROCESSED:", len(d))
    if d:
        last = d[-1]
        print("LAST_IMAGE:", last.get("Image Name"))
        print("LAST_TYPE:", last.get("Invoice Type"))
        print("LAST_REF:", last.get("Invoice Number"))
        print("LAST_AMT:", last.get("Amount"))
        print("LAST_VENDOR:", last.get("Vendor Name"))
        print("LAST_PROP:", last.get("Property Name"))
else:
    print("File not found")
