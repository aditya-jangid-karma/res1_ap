import shutil
from pathlib import Path

src_dir = Path("Invoice_Indexer/automation")
dest_dir = Path("Z:/Res-1-indexer")
dest_folder = dest_dir / "extractor_workspace"

dest_folder.mkdir(parents=True, exist_ok=True)

# 1. Copy zip file
zip_name = "deleted_invoices_extractor_bundle.zip"
src_zip = src_dir / zip_name
if src_zip.exists():
    shutil.copy2(src_zip, dest_dir / zip_name)
    print(f"SUCCESS: Copied {zip_name} to {dest_dir / zip_name}")

# 2. Copy individual files directly to folder for instant access
files = [
    "scan_deleted_pages.py",
    "extract_deleted_invoices_with_qwen.py",
    "reader_qwen.py",
    "reader_pdf.py",
    "invoice_schema.py",
    "fast_matcher.py",
    "deleted_invoices_last_2_months.xlsx",
    "deleted_invoices_progress.json",
    "deleted_records_last_2_months.json",
    "execution_tracking.log",
    "check_progress_status.py",
    ".bottomline_session.json"
]

copied_count = 0
for f in files:
    src_file = src_dir / f
    if src_file.exists():
        shutil.copy2(src_file, dest_folder / f)
        copied_count += 1

print(f"SUCCESS: Copied {copied_count} files directly to {dest_folder}")

