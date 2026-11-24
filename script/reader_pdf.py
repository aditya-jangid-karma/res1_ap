"""Render a PDF's pages to PNG bytes with Poppler (via pdf2image). This is the
ONLY file that touches Poppler. 200 DPI, at most 3 pages (see the spec)."""
import io


def render_pdf(pdf_path, dpi=200, max_pages=3, poppler_path=None, _convert=None):
    """Up to `max_pages` pages of `pdf_path` as a list of PNG byte-strings.
    Supports PyMuPDF (fitz) out-of-the-box without Poppler, or pdf2image as fallback."""
    # 1. Try PyMuPDF (pymupdf / fitz) first: fast, pure pip package, no Poppler binary required
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz

        doc = fitz.open(str(pdf_path))
        png_bytes = []
        # Calculate matrix scaling for requested DPI (72 is standard PDF base DPI)
        scale = dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        for page_num in range(min(len(doc), max_pages)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=matrix)
            png_bytes.append(pix.tobytes("png"))
        doc.close()
        if png_bytes:
            return png_bytes
    except Exception as e:
        # If PyMuPDF failed or isn't installed, continue to fallback
        pass


    # 2. Fallback to pdf2image + Poppler
    if _convert is None:
        from pdf2image import convert_from_path as _convert

    pages = _convert(pdf_path, dpi=dpi, first_page=1, last_page=max_pages,
                     poppler_path=poppler_path)

    png_bytes = []
    for page in pages[:max_pages]:
        buffer = io.BytesIO()
        page.save(buffer, format="PNG")
        png_bytes.append(buffer.getvalue())
    return png_bytes

