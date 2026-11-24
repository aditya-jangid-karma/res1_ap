import openpyxl
import csv
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent
MASTERS = PROJECT / "masters"
QUERY_XLSX = MASTERS / "Query Resolution Sheet.xlsx"

wb = openpyxl.load_workbook(QUERY_XLSX, data_only=True)

# 1. Parse Property - Query Resolution
prop_ws = wb["Property - Query Resolution"]
prop_mappings = {}
unmanaged_props = set()

for row in prop_ws.iter_rows(values_only=True):
    if not any(row):
        continue
    status = str(row[0] or "").strip()
    raw_prop = str(row[1] or "").strip()
    action = str(row[2] or "").strip()
    resol = str(row[3] or "").strip()

    if not raw_prop or raw_prop.lower() in ("property name", "name"):
        continue

    text_all = f"{status} {action} {resol}".lower()
    if any(k in text_all for k in ["delete", "no longer managed", "not managed"]):
        unmanaged_props.add(raw_prop)
    elif "index to" in text_all or "index as" in text_all or "index with" in text_all:
        # Extract target property name
        m = re.search(r"(?i)index\s+(?:it\s+)?(?:to|as|with)\s+([A-Za-z0-9\s\-&]+)", f"{action} {resol}")
        if m:
            target = m.group(1).strip()
            # Clean trailing words like 'Property'
            target = re.sub(r"(?i)\s+property$", "", target).strip()
            prop_mappings[raw_prop] = target
        elif action and not any(k in action.lower() for k in ["delete", "managed", "yardi property"]):
            prop_mappings[raw_prop] = action

# Also check Ask Query for property mappings
if "Ask Query" in wb.sheetnames:
    ask_ws = wb["Ask Query"]
    for row in ask_ws.iter_rows(values_only=True):
        if not any(row):
            continue
        q_type = str(row[0] or "")
        vp = str(row[1] or "")
        sol = str(row[3] or "")
        if "Property Not Found" in q_type and vp and sol:
            m = re.search(r"(?i)index\s+(?:to|as|with)\s+([A-Za-z0-9\s\-&]+)", sol)
            if m:
                prop_mappings[vp] = m.group(1).strip()

print(f"Loaded {len(prop_mappings)} property redirects and {len(unmanaged_props)} unmanaged properties.")

# 2. Parse Vendor - Query Resolution
vendor_ws = wb["Vendor - Query Resolution"]
vendor_mappings = {}
vendor_code_mappings = {}
do_not_index_vendors = set()

for row in vendor_ws.iter_rows(values_only=True):
    if not any(row):
        continue
    q_type = str(row[0] or "").strip()
    raw_vendor = str(row[1] or "").strip()
    action = str(row[2] or "").strip()
    resol = str(row[3] or "").strip()

    if not raw_vendor or raw_vendor.lower() in ("vendor name", "name"):
        continue

    text_all = f"{q_type} {action} {resol}".lower()
    if "do not index" in text_all:
        do_not_index_vendors.add(raw_vendor)
        continue

    # Extract vendor code if specified (e.g. 5604-001, 3600-001)
    code_match = re.search(r"\b(\d{2,4}-\d{3})\b", f"{action} {resol}")
    if code_match:
        vendor_code_mappings[raw_vendor] = code_match.group(1)

    # Extract target vendor name if mapped to a different name
    if "index with" in text_all or "index to" in text_all or "index as" in text_all:
        m = re.search(r"(?i)index\s+(?:to|with|as)\s+([A-Za-z0-9\s\.,&]+)", f"{action} {resol}")
        if m:
            target = m.group(1).strip()
            # If target is vendor code only, don't use as name
            if not re.match(r"^\d{2,4}-\d{3}$", target) and "vendor code" not in target.lower():
                vendor_mappings[raw_vendor] = target

# Check Ask Query for vendor mappings
if "Ask Query" in wb.sheetnames:
    ask_ws = wb["Ask Query"]
    for row in ask_ws.iter_rows(values_only=True):
        if not any(row):
            continue
        q_type = str(row[0] or "")
        vp = str(row[1] or "")
        sol = str(row[3] or "")
        if "vendor" in q_type.lower() or "multiple vendors" in q_type.lower():
            code_match = re.search(r"\b(\d{2,4}-\d{3})\b", sol)
            if code_match and vp:
                vendor_code_mappings[vp] = code_match.group(1)

print(f"Loaded {len(vendor_mappings)} vendor redirects, {len(vendor_code_mappings)} vendor code mappings, and {len(do_not_index_vendors)} do-not-index vendors.")

# 3. Export to masters/query_resolutions.json
import json
output = {
    "property_redirects": prop_mappings,
    "unmanaged_properties": sorted(list(unmanaged_props)),
    "vendor_redirects": vendor_mappings,
    "vendor_codes": vendor_code_mappings,
    "do_not_index_vendors": sorted(list(do_not_index_vendors))
}

out_file = MASTERS / "query_resolutions.json"
with open(out_file, "w", encoding="utf-8") as f:
    json.dump(output, f, indent=2)

print(f"Saved query resolution index to: {out_file}")

# 4. Append confirmed property redirects to property_aliases.csv
prop_aliases_file = MASTERS / "property_aliases.csv"
existing_aliases = set()
if prop_aliases_file.exists():
    with open(prop_aliases_file, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        for r in reader:
            if len(r) >= 2:
                existing_aliases.add(r[0].strip().lower())

new_aliases = []
for raw, target in prop_mappings.items():
    if raw.lower() not in existing_aliases:
        new_aliases.append([raw, target])

if new_aliases:
    with open(prop_aliases_file, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(new_aliases)
    print(f"Appended {len(new_aliases)} new property aliases to {prop_aliases_file.name}")
