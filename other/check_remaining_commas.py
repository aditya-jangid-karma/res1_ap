import openpyxl
from pathlib import Path

fpath = Path(r"C:\Users\Meet Patel\Desktop\automation\Res-1 Github\Res-1\Invoice_Indexer\masters\Vendor Proprty masters.xlsx")
wb = openpyxl.load_workbook(fpath, data_only=True)
s = wb.active

for r in range(2, s.max_row + 1):
    raw_code = str(s.cell(row=r, column=4).value or '').strip()
    if ',' in raw_code:
        parts = [p.strip() for p in raw_code.split(',') if p.strip()]
        has_pure_digit = any(p.isdigit() for p in parts)
        has_non_digit = any(not p.isdigit() for p in parts)
        if not (has_pure_digit and has_non_digit):
            print(f"Row {r}: '{raw_code}' -> parts={parts}")
