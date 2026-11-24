import openpyxl
from pathlib import Path

fpath = Path(r"C:\Users\Meet Patel\Desktop\automation\Res-1 Github\Res-1\Invoice_Indexer\masters\Vendor Proprty masters.xlsx")
wb = openpyxl.load_workbook(fpath, data_only=True)
s = wb.active

all_codes = []
for r in range(2, s.max_row + 1):
    raw_code = str(s.cell(row=r, column=4).value or '').strip()
    if not raw_code:
        continue
    parts = [p.strip() for p in raw_code.split(',') if p.strip()]
    intacct = []
    yardi = []
    for p in parts:
        if p.isdigit():
            intacct.append(p)
        else:
            yardi.append(p)
    all_codes.append((r, s.cell(row=r, column=3).value, raw_code, intacct, yardi))

print("Total records parsed:", len(all_codes))
samples_both = [x for x in all_codes if x[3] and x[4]]
print(f"Records with both Intacct and Yardi: {len(samples_both)}")
for x in samples_both[:10]:
    print(f"Row {x[0]}: Vendor='{x[1]}' | Raw='{x[2]}' -> Intacct={x[3]}, Yardi={x[4]}")
