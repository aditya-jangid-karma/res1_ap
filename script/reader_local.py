"""reader_local.py -- read invoice PDFs with LOCAL Python only.

Replaces the Ollama/Qwen vision call (reader_qwen.py / reader_qwen_utility.py)
for TASK 1. Nothing here touches the network or another machine:

  * PDFs with a text layer are read with PyMuPDF (word positions -> visual
    lines -> column segments).
  * Pages with no usable text (scans / photos) are OCR'd on this PC with
    RapidOCR (ONNX models ship inside the pip package).
  * Fields are found with layout-aware label/value rules (value to the right
    of a label, or in the same column on the line below), cross-checked
    (subtotal + tax = total, line items = subtotal), and the vendor is
    confirmed against masters/vendors_data.csv (name, remit address, zip).

extract_invoice_local(pdf_path, vendor_rows=None) returns a dict with the
SAME keys the Qwen reader returned (vendor, invoice_number, invoice_date,
due_date, amount, vendor_address, service_address, property,
account_number, doc_type, + utility keys) plus extra keys used for the
LineItems sheet and the conversion report (doc_class, po_number, subtotal,
tax, freight, line_items, sub_documents, notes, ocr_used, pages).
"""
import re
import threading
from datetime import datetime
from pathlib import Path

try:
    import pymupdf as fitz
except ImportError:  # older PyMuPDF installs only expose 'fitz'
    import fitz

try:
    from rapidfuzz import fuzz
except ImportError:  # pragma: no cover
    fuzz = None

# --------------------------------------------------------------------------
# 1. TEXT LAYER / OCR  ->  pages of lines of segments
# --------------------------------------------------------------------------
MIN_TEXT_CHARS = 40     # fewer real characters than this on a page -> OCR it
OCR_DPI = 200
_OCR = None
_OCR_LOCK = threading.Lock()
_OCR_RUN_LOCK = threading.Lock()

_CHAR_FIX = str.maketrans({"\u2212": "-", "\u2013": "-", "\u2014": "-",
                           "\u2019": "'", "\u2018": "'", "\u201c": '"',
                           "\u201d": '"', "\u00a0": " ", "\uff0c": ",",
                           "\uff1a": ":"})


def _ocr_engine():
    global _OCR
    with _OCR_LOCK:
        if _OCR is None:
            import logging
            for name in ("RapidOCR", "rapidocr"):
                logging.getLogger(name).setLevel(logging.ERROR)
            from rapidocr import RapidOCR
            _OCR = RapidOCR()
    return _OCR


def _ocr_items(page, clip=None):
    """OCR one page (or a clip of it) -> word-ish items in PDF points."""
    import numpy as np
    pix = page.get_pixmap(dpi=OCR_DPI, clip=clip)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    img = img[:, :, :3].copy()
    with _OCR_RUN_LOCK:            # onnxruntime session: one call at a time
        res = _ocr_engine()(img)
    scale = 72.0 / OCR_DPI
    ox = clip.x0 if clip is not None else 0
    oy = clip.y0 if clip is not None else 0
    items = []
    if res is not None and getattr(res, "boxes", None) is not None:
        for box, txt, score in zip(res.boxes, res.txts, res.scores):
            txt = (txt or "").translate(_CHAR_FIX).strip()
            if not txt or float(score) < 0.5:
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            items.append((ox + min(xs) * scale, oy + min(ys) * scale,
                          ox + max(xs) * scale, oy + max(ys) * scale, txt))
    return items


_MONEY_TOKEN = re.compile(r"\(?-?\$\s?[\d,]*\.\d{2}\)?|\(?-?[\d,]*\.\d{2}\)?")


def _group_lines(items, page_no, page_h, page_w):
    """(x0,y0,x1,y1,text) items -> visual lines, each split into segments at
    big horizontal gaps (columns)."""
    items = sorted(items, key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    rows = []
    for w in items:
        cy = (w[1] + w[3]) / 2
        h = max(1.0, w[3] - w[1])
        for ln in reversed(rows[-6:]):
            if abs(ln["cy"] - cy) <= max(2.5, 0.45 * min(h, ln["h"])):
                ln["words"].append(w)
                break
        else:
            rows.append({"cy": cy, "h": h, "words": [w]})
    lines = []
    for ln in rows:
        ws = sorted(ln["words"], key=lambda w: w[0])
        avg_h = sum(w[3] - w[1] for w in ws) / len(ws)
        segs, cur = [], [ws[0]]
        for prev, w in zip(ws, ws[1:]):
            gap = w[0] - prev[2]
            money_pair = bool(_MONEY_TOKEN.fullmatch(prev[4]) and _MONEY_TOKEN.fullmatch(w[4]))
            if gap > max(6.0, 1.15 * avg_h) or money_pair:
                segs.append(cur)
                cur = [w]
            else:
                cur.append(w)
        segs.append(cur)
        seg_objs = [{"text": " ".join(x[4] for x in s).strip(),
                     "x0": min(x[0] for x in s), "x1": max(x[2] for x in s),
                     "y0": min(x[1] for x in s), "y1": max(x[3] for x in s)}
                    for s in segs]
        seg_objs = [s for s in seg_objs if s["text"]]
        if not seg_objs:
            continue
        lines.append({"page": page_no, "segs": seg_objs,
                      "text": "   ".join(s["text"] for s in seg_objs),
                      "x0": min(s["x0"] for s in seg_objs),
                      "x1": max(s["x1"] for s in seg_objs),
                      "y0": min(s["y0"] for s in seg_objs),
                      "y1": max(s["y1"] for s in seg_objs),
                      "page_h": page_h, "page_w": page_w})
    lines.sort(key=lambda l: (round(l["y0"] / 2), l["x0"]))
    return lines


def _split_money_runs(items):
    """OCR often glues neighbouring columns: '$170.00 $510.00' -> two words."""
    out = []
    for x0, y0, x1, y1, t in items:
        toks = t.split()
        if len(toks) >= 2 and all(re.fullmatch(r"\(?-?\$?[\d,]*\.?\d+\)?", k) for k in toks):
            total = sum(len(k) for k in toks) + len(toks) - 1
            pos = 0
            for k in toks:
                a0 = x0 + (x1 - x0) * pos / total
                a1 = x0 + (x1 - x0) * (pos + len(k)) / total
                out.append((a0, y0, a1, y1, k))
                pos += len(k) + 1
        else:
            out.append((x0, y0, x1, y1, t))
    return out


def read_pdf_lines(path, max_pages=None, force_ocr=False):
    """-> (pages, ocr_pages). pages = [[line, ...] per page]."""
    doc = fitz.open(str(path))
    pages, ocr_pages = [], []
    try:
        n = len(doc) if max_pages is None else min(len(doc), max_pages)
        for i in range(n):
            page = doc.load_page(i)
            words = page.get_text("words", sort=True)
            items = [(w[0], w[1], w[2], w[3], w[4].translate(_CHAR_FIX))
                     for w in words if w[4].strip()]
            real = sum(len(w[4]) for w in items)
            scanned = False
            if MIN_TEXT_CHARS <= real < 400:
                # a scan with a little typed text on top (e.g. a stamped name)
                try:
                    area = page.rect.width * page.rect.height
                    scanned = any((b["bbox"][2] - b["bbox"][0]) * (b["bbox"][3] - b["bbox"][1]) > 0.5 * area
                                  for b in page.get_image_info())
                except Exception:
                    scanned = False
            if force_ocr or real < MIN_TEXT_CHARS or scanned:
                ocr = _ocr_items(page)
                if sum(len(t[4]) for t in ocr) > real:
                    def _overlaps(w):
                        return any(not (w[2] < o[0] or w[0] > o[2] or w[3] < o[1] or w[1] > o[3]) for o in ocr)
                    items = _split_money_runs(ocr) + [w for w in items if not _overlaps(w)]
                    ocr_pages.append(i + 1)
            pages.append(_group_lines(items, i, page.rect.height, page.rect.width))
    finally:
        doc.close()
    return pages, ocr_pages


def ocr_page_header(path, page_index=0, frac=0.30):
    """OCR just the top part of a page (logos often hold the vendor name)."""
    doc = fitz.open(str(path))
    try:
        page = doc.load_page(page_index)
        r = page.rect
        clip = fitz.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * frac)
        return _group_lines(_ocr_items(page, clip), page_index, r.height, r.width)
    finally:
        doc.close()


# --------------------------------------------------------------------------
# 2. VALUE PARSERS
# --------------------------------------------------------------------------
_MONTHS = ("jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|"
           "february|march|april|june|july|august|september|october|"
           "november|december")
DATE_RE = re.compile(
    r"(?<![\d/])(\d{1,2}[/-]\d{1,2}[/-](?:\d{4}|\d{2}))(?![\d/])"
    r"|((?:" + _MONTHS + r")\.?\s+\d{1,2},?\s+\d{4})"
    r"|(\d{4}-\d{2}-\d{2})", re.I)
MONEY_RE = re.compile(
    r"(?<![\w.])\(?-?\$?\s?-?\(?\d{1,3}(?:,\d{3})*(?:\.\d{1,3})?\)?-?(?![\w%/])"
    r"|(?<![\w.])\(?-?\$?\s?-?\.\d{2}\)?(?![\w%/])")


def parse_date(text):
    """First date in text -> 'MM/DD/YYYY' (US month-first), else None."""
    if not text:
        return None
    m = DATE_RE.search(str(text))
    if not m:
        return None
    s = m.group(0).replace(",", "").replace(".", "")
    s = re.sub(r"\s+", " ", s).strip()
    fmts = ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y", "%Y-%m-%d",
            "%B %d %Y", "%b %d %Y")
    for f in fmts:
        try:
            d = datetime.strptime(s, f)
            if 1990 <= d.year <= 2100:
                return d.strftime("%m/%d/%Y")
        except ValueError:
            continue
    if s.lower().startswith("sept"):
        return parse_date(s.replace("Sept", "Sep").replace("sept", "sep"))
    return None


def parse_money(text, require_decimal=False):
    """'$1,234.50' -> 1234.5 ; '(555.08)' / '$-1.00' / '12.00-' -> negative.
    Returns None when text holds no amount."""
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    if DATE_RE.fullmatch(s):
        return None
    s2 = s.replace("USD", "").replace("US$", "").strip().rstrip("TNY*").strip()
    m = re.fullmatch(r"\(?\s*-?\s*\$?\s*-?\s*\(?\s*((?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d{1,3})?)\s*\)?\s*-?", s2)
    if not m or not m.group(1) or not re.search(r"\d", m.group(1)):
        return None
    num = m.group(1).replace(",", "")
    if require_decimal and "." not in num:
        return None
    try:
        val = float(num)
    except ValueError:
        return None
    neg = ("(" in s2 and ")" in s2) or s2.startswith("-") or "$-" in s2 \
        or s2.endswith("-") or "- " in s2[:3]
    return -val if neg else val


def money_in(text, require_decimal=True):
    """All amounts in a string (decimals required by default)."""
    out = []
    for m in MONEY_RE.finditer(str(text)):
        v = parse_money(m.group(0), require_decimal=require_decimal)
        if v is not None:
            out.append(v)
    return out


def fmt_money(v):
    return None if v is None else f"{v:.2f}"


# --------------------------------------------------------------------------
# 3. LABEL / VALUE SEARCH ON THE LAYOUT
# --------------------------------------------------------------------------
def _norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _all_lines(pages):
    return [ln for p in pages for ln in p]


def _x_overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def find_values(pages, label_re, validator, max_below=3, page_limit=None,
                below=True, right=True, mid=False, pages_subset=None):
    """Every (value, info) for a label regex. The label must START a segment
    (or, with mid=True, may follow other words in it). Value: rest of that
    segment, else the next segments on the line, else a segment directly
    below (same column) in the next few lines. validator(text) -> value|None."""
    if mid:
        rx = re.compile(r"(?:^|(?<=\s))(?:" + label_re + r")\s*[:#.]?\s*", re.I)
    else:
        rx = re.compile(r"^\s*(?:" + label_re + r")\s*[:#.]?\s*", re.I)
    hits = []
    for pi, page in enumerate(pages):
        if page_limit is not None and pi >= page_limit:
            break
        for li, ln in enumerate(page):
            for si, seg in enumerate(ln["segs"]):
                m = rx.search(seg["text"]) if mid else rx.match(seg["text"])
                if not m:
                    continue
                rest = seg["text"][m.end():].strip(" :#")
                cands = []
                if right and rest:
                    cands.append(("same", rest))
                if right:
                    for s2 in ln["segs"][si + 1:si + 4]:
                        cands.append(("right", s2["text"]))
                if below:
                    lab_x0, lab_x1 = seg["x0"], seg["x1"]
                    if mid and m.start() > 0:
                        frac = m.start() / max(1, len(seg["text"]))
                        lab_x0 = seg["x0"] + frac * (seg["x1"] - seg["x0"])
                    last_y = ln["y1"]
                    for lj in range(li + 1, min(len(page), li + 1 + max_below)):
                        nxt = page[lj]
                        if nxt["y0"] - last_y > 45:
                            break
                        last_y = nxt["y1"]
                        best, best_ov = None, 0.0
                        for s2 in nxt["segs"]:
                            ov = _x_overlap(lab_x0 - 4, lab_x1 + 4, s2["x0"], s2["x1"])
                            if ov > best_ov:
                                best, best_ov = s2, ov
                        if best is None and nxt["segs"]:
                            s2 = nxt["segs"][-1]   # right-aligned value just right of the label
                            if s2["x0"] >= lab_x0 and 0 <= s2["x0"] - lab_x1 <= 50:
                                best = s2
                        if best is not None:
                            cands.append(("below", best["text"]))
                for how, txt in cands:
                    v = validator(txt)
                    if v not in (None, ""):
                        hits.append((v, {"how": how, "page": pi, "line": li,
                                         "y": ln["y0"], "label": m.group(0).strip(),
                                         "seg": seg["text"], "line_text": ln["text"]}))
                        break
    return hits


# --- validators -----------------------------------------------------------
_LABEL_WORDS = re.compile(
    r"^(date|invoice|number|no\.?|terms|total|amount|due|page|customer|account|"
    r"po|p\.o\.|order|ship|bill|to|from|qty|description|balance|reference|"
    r"salesperson|salesman|contract|type|rep|net|remit|purchase|attn|route)\b", re.I)


def v_invoice_number(txt):
    t = _norm(txt).strip(" :#")
    t = re.sub(r"^(?:no\.?|number|#)\s*[:#]?\s*", "", t, flags=re.I)
    if not t:
        return None
    tok = t.split(" ")[0].strip(" ,;:")
    if len(t.split(" ")) > 1 and re.fullmatch(r"[A-Za-z]{1,4}", tok):
        # 'INV 12345' style -> join
        nxt = t.split(" ")[1]
        if re.search(r"\d", nxt):
            tok = tok + nxt
    tok = tok.strip(" ,;:")
    if not re.search(r"\d", tok) or len(tok) < 3 or len(tok) > 25:
        return None
    if DATE_RE.fullmatch(tok) or tok.startswith("$"):
        return None
    if re.fullmatch(r"\d{1,3}(,\d{3})*\.\d{2}", tok):
        return None
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-/._#]*", tok):
        return None
    return tok


def v_date(txt):
    return parse_date(txt)


def v_money(txt):
    t = _norm(txt)
    if not t:
        return None
    first = t.split("   ")[0]
    first = re.sub(r"^\$\s+", "$", first)
    v = parse_money(first)
    if v is None:
        ms = money_in(first)
        v = ms[0] if len(ms) == 1 else None
    return v


def v_token(txt, min_len=2, max_len=30):
    t = _norm(txt).strip(" :#")
    if not t or _LABEL_WORDS.match(t) and not re.search(r"\d", t):
        return None
    if len(t) > max_len:
        t = t.split(" ")[0]
    if len(t) < min_len or len(t) > max_len:
        return None
    if DATE_RE.fullmatch(t):
        return None
    return t


def v_account(txt):
    t = v_token(txt, 3, 25)
    if not t or not re.search(r"\d", t):
        return None
    if " " in t and not re.fullmatch(r"[A-Z]{1,3} \d+", t):
        t = t.split(" ")[0]
        if not re.search(r"\d", t):
            return None
    return t


def v_po(txt):
    t = v_token(txt, 2, 30)
    if not t:
        return None
    if re.match(r"^(remit|contract|terms|ship|sales|quote|net|due|reference|"
                r"order|release|description|bill|customer|start|remarks|field|"
                r"technician|item|shipvia|via|date|type|job|salesman|occupancy|ordered|unit|service|rate|qty|amount|"
                r"invoice|page|account|po\b|p\.o\.)", t, re.I):
        return None
    if t.upper() in ("N/A", "NONE", "NA"):
        return None
    if len(t.split()) > 3 or _PROP_WORDS.search(t) or re.search(r"\s-\s*R1\b|office|suite", t, re.I) \
            or _STREET_RE.match(t) and len(t.split()) > 1:
        return None
    if not re.search(r"\d", t):
        # words only: accept a single ALL-CAPS token such as SHOP / TRAP / CREDIT
        if not re.fullmatch(r"[A-Z][A-Z&\-]{2,15}", t):
            return None
    return t


def _first(hits):
    return hits[0][0] if hits else None


# --------------------------------------------------------------------------
# 4. CUSTOMER / ADDRESS BLOCKS (Ship To / Bill To / Service At ...)
# --------------------------------------------------------------------------
_CUSTOMER_RE = re.compile(
    r"residential\s*one|\bres\s?1\b|res1\.net|c/o residential|\bR1\b|8975\s+guilford", re.I)
_STREET_RE = re.compile(
    r"^\s*(?:\d+[A-Za-z]?(?:-\d+)?\s+[A-Za-z0-9]|p\.?\s*o\.?\s*box)", re.I)
_CITY_RE = re.compile(
    r"\b[A-Za-z][A-Za-z .]+,?\s+(?:[A-Z]{2}|Maryland|Virginia|District of Columbia)\s*,?\s*\d{5}(?:-\d{4})?\b")
_STOP_BLOCK = re.compile(
    r"^(item|qty|quantity|description|date|invoice|terms|payment|po\b|p\.o\.|"
    r"ship point|ship date|sales|order|account|work|service date|activity|billing date|"
    r"phone|tel\b|fax|email|e-mail|customer|ordered|stock|ln#|code|"
    r"special|instructions|remarks|contract|period|memo|amount|authorized|"
    r"route|total|subtotal|balance|current|bill to|ship to|sold to)", re.I)
BLOCK_LABEL_START = re.compile(
    r"^(ship|bill|sold|service|location|job|invoice|customer|remit|deliver|"
    r"property|compactor|date|page|payment|account|order|terms|due|to:)", re.I)


def address_block(pages, label_re, max_lines=6, page_limit=1):
    """-> (lines_of_block, label_text). Lines under/after a label like 'Ship To'."""
    rx = re.compile(r"^\s*(?:" + label_re + r")\s*[:#]?\s*", re.I)
    fallback = None
    for pi, page in enumerate(pages[:page_limit]):
        for li, ln in enumerate(page):
            for si, seg in enumerate(ln["segs"]):
                m = rx.match(seg["text"])
                if not m:
                    continue
                rest = seg["text"][m.end():].strip()
                rest = re.sub(r"(?i)^address\s*:?\s*", "", rest)
                block = []
                col_x0 = seg["x0"]
                right_lim = None
                nxt = ln["segs"][si + 1] if si + 1 < len(ln["segs"]) else None
                if rest:
                    block.append(rest)
                    if nxt is not None:
                        right_lim = nxt["x0"] - 2
                elif nxt is not None:
                    if BLOCK_LABEL_START.match(nxt["text"]) or (nxt["x0"] - seg["x1"] > 120 and not _STREET_RE.match(nxt["text"])) \
                            or re.match(r"(?i)please|remit|residential one", nxt["text"]):
                        right_lim = nxt["x0"] - 2
                    else:
                        block.append(nxt["text"])
                        col_x0 = nxt["x0"]
                        if si + 2 < len(ln["segs"]):
                            right_lim = ln["segs"][si + 2]["x0"] - 2
                last_y = ln["y1"]
                for lj in range(li + 1, min(len(page), li + 12)):
                    nl = page[lj]
                    if nl["y0"] - last_y > (50 if not block else 26):
                        break
                    first_txt = nl["segs"][0]["text"].strip()
                    if nl["segs"][0]["x0"] <= col_x0 + 5 and _STOP_BLOCK.match(first_txt) \
                            and not _STREET_RE.match(first_txt):
                        break
                    picked = None
                    for k2, s2 in enumerate(nl["segs"]):
                        is_last = k2 == len(nl["segs"]) - 1
                        if s2["x1"] < col_x0 - 3 and not (is_last and s2["x1"] >= col_x0 - 30):
                            continue
                        if s2["x0"] < col_x0 - 20 and not is_last:
                            continue
                        if right_lim is not None and s2["x0"] >= right_lim:
                            continue
                        if s2["x0"] > col_x0 + 120:
                            continue
                        picked = s2
                        break
                    if picked is None:
                        continue
                    t = picked["text"].strip()
                    if (_STOP_BLOCK.match(t) and not _STREET_RE.match(t)) or t.startswith("("):
                        break
                    block.append(t)
                    last_y = nl["y1"]
                    if _CITY_RE.search(t) or len(block) >= max_lines:
                        break
                block = [b for b in block if b]
                if block:
                    if any(_STREET_RE.match(b) or _CITY_RE.search(b) for b in block):
                        return block, m.group(0).strip()
                    if fallback is None:
                        fallback = (block, m.group(0).strip())
    return fallback if fallback else ([], None)


_PROP_WORDS = re.compile(
    r"\b(apartments?|apts?\.?|manor|house|towers?|commons|village|courts?|square|"
    r"heights|crossing|place|flats|terrace|gardens?|building|condo(minium)?|park|"
    r"senior|landing|summit|haven|station|plaza|estates|homes|residences?|lofts?|"
    r"hills|glen|woods|point|pointe|ridge|view|green|mar|maples|westerly)\b", re.I)


_STREET_SUFFIX = re.compile(
    r"\b(st|street|ave|avenue|rd|road|dr|drive|blvd|boulevard|ln|lane|way|ct|pl|pkwy|parkway|"
    r"hwy|highway|cir|circle|ter|pike|suite|ste|unit|nw|ne|sw|se)\b\.?", re.I)


def _person_like(t):
    t = t.strip()
    return bool(re.fullmatch(r"(mr|mrs|ms|dr)?\.?\s*[A-Z][a-z]+(\s+[A-Z]\.?)?\s+[A-Z][a-z]+", t)
                or re.fullmatch(r"[A-Z]{2,12}\s+[A-Z]{2,14}", t)) \
        and not _PROP_WORDS.search(t) and not _COMPANY_HINT.search(t)


def split_block(block):
    """-> (name, street_address_text). Drops Attn/person/phone lines."""
    if not block:
        return None, None
    name_lines, addr_lines = [], []
    for b in block:
        t = b.strip(" ,")
        if re.match(r"^(attn|attention)\b", t, re.I):
            continue
        if re.fullmatch(r"[\d()\-./ ]{7,}", t) or re.match(r"^\(?\d{3}\)?[\s.-]?\d{3}[\s.-]\d{4}", t):
            continue
        if re.fullmatch(r"(usa|us|united states)", t, re.I):
            continue
        numbered_name = (not addr_lines and not name_lines and _PROP_WORDS.search(t)
                         and not _STREET_SUFFIX.search(t) and not re.match(r"(?i)\s*p\.?\s*o\.?\s*box", t))
        if numbered_name:
            name_lines.append(t)
            continue
        if _STREET_RE.match(t) or _CITY_RE.search(t) or (addr_lines and not _PROP_WORDS.search(t)):
            addr_lines.append(t)
        else:
            name_lines.append(t)
    name = None
    name_lines = [n for n in name_lines if clean_property(n)]
    ranked = sorted(name_lines, key=lambda n: (
        0 if _PROP_WORDS.search(clean_property(n)) else
        1 if not _CUSTOMER_RE.search(n) and not _person_like(n) else
        2 if _CUSTOMER_RE.search(n) else 3))
    if ranked:
        name = ranked[0]
        if _person_like(name):
            name = None
    addr = ", ".join(a for a in addr_lines[:3]) or None
    return name, addr


def clean_property(name):
    if not name:
        return None
    n = re.sub(r"(?i)\s*[-/]?\s*(c/o\s*)?residential\s*one(\s*(llc|management|apartments))?\s*[-/]?\s*", " ", name)
    n = re.sub(r"(?i)\s+-\s*R1\b", "", n)
    n = re.sub(r"\s+", " ", n).strip(" -/,")
    if _PROP_WORDS.search(n):
        n = re.sub(r"(?i),?\s+(llc|l\.l\.c\.?|inc\.?)$", "", n).strip(" ,")
    if len(n) < 3 or len(n) > 70 or n.endswith(":") or not re.search(r"[A-Za-z]{3}", n) \
            or re.fullmatch(r"[ADFT]{12,}", n) or re.search(r"(?i)\bname\b", n) or re.match(
            r"(?i)^(counter pick up|leasing office.*|management office|accounts payable|ap dept|attn.*|"
            r"united states|usa|bill to|ship to|sold to|terms|statement date|in store purchase.*|"
            r"please.*|number|no|copy|customer|page.*|date.*|invoice.*)$", n):
        return None
    if _STOP_BLOCK.match(n) and not _PROP_WORDS.search(n):
        return None
    return n


# --------------------------------------------------------------------------
# 5. VENDOR IDENTIFICATION  (printed name + masters/vendors_data.csv)
# --------------------------------------------------------------------------
_SUFFIX_ALT = (r"l\.?l\.?c\.?|inc\.?|incorporated|corp\.?|corporation|company|co\.|"
               r"ltd\.?|limited|l\.?p\.?|pllc|pc")
_SUFFIX_RE = re.compile(r"\b(" + _SUFFIX_ALT + r")(?=\W|$)", re.I)
_COMPANY_HINT = re.compile(
    r"\b(llc|l\.l\.c|inc|incorporated|corp|corporation|company|ltd|services|"
    r"solutions|supply|systems|plumbing|contractors|staffing|refinishing|"
    r"cleaning|software|mechanical|products|enterprises|industries|"
    r"international|electric|elevator|landscap\w*)\b\.?", re.I)
_GENERIC_TOP = re.compile(
    r"^(invoice|sales invoice|contract invoice|maintenance invoice|statement|"
    r"monthly statement|credit memo|bill to|ship to|page|remit|date|"
    r"customer|original invoice|open invoice detail|sign up|visit|"
    r"login token|terms|statement of account|shipping receipt|attn|contact|"
    r"please|thank|for questions|questions|payment|amount|total)\b", re.I)
_GENERIC_CORES = {
    "apartments", "apartment", "services", "service", "management", "supply",
    "maintenance", "invoice", "the", "office", "leasing office", "residential",
    "residential one", "property", "properties", "cleaning", "plumbing",
    "electric", "construction", "painting", "landscaping", "contractors",
    "staffing", "security", "solutions", "systems", "home", "depot", "general",
    "washington", "maryland", "baltimore", "columbia", "silver spring",
}
_GENERIC_DOMAINS = {
    "gmail", "yahoo", "outlook", "hotmail", "aol", "icloud", "google", "drive",
    "billtrust", "ceojuice", "payments", "res1", "www", "com", "net", "org",
    "us", "co", "info", "mail", "versapay", "paypal", "intuit", "quickbooks",
}


def _core(name):
    """'Chadwell Supply, Inc. (ATLANTA/30348-5172)' -> 'chadwell supply'."""
    n = re.sub(r"\(.*?\)", " ", str(name or ""))
    n = re.sub(r"\s-\s.*$", " ", n)              # 'Ferguson - Pittsburgh'
    n = _SUFFIX_RE.sub(" ", n)
    n = n.lower().replace("&", " and ").replace("'", "")
    n = re.sub(r"[^a-z0-9 ]+", " ", n)
    return re.sub(r"\s+", " ", n).strip()


def _compact(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def _norm_addr(s):
    s = str(s or "").lower()
    s = re.sub(r"p\.?\s*o\.?\s*box", "po box", s)
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    for a, b in (("drive", "dr"), ("street", "st"), ("avenue", "ave"),
                 ("road", "rd"), ("boulevard", "blvd"), ("place", "pl"),
                 ("highway", "hwy"), ("parkway", "pkwy"), ("suite", "ste"),
                 ("court", "ct"), ("lane", "ln"), ("circle", "cir")):
        s = re.sub(r"\b%s\b" % a, b, s)
    return re.sub(r"\s+", " ", s).strip()


class VendorIndex:
    """Index of masters/vendors_data.csv for evidence scoring."""

    def __init__(self, rows, name_col="vendor", address_col="vendorAddress",
                 zip_col="zipcode"):
        self.entries = []
        seen = set()
        for r in rows or []:
            name = (r.get(name_col) or "").strip()
            if not name:
                continue
            core = _core(name)
            if len(core) < 3 or core in _GENERIC_CORES:
                continue
            if name.rstrip().endswith("(-/-)") and _person_like(re.sub(r"\s*\(.*?\)\s*$", "", name)):
                continue
            addr = _norm_addr((r.get(address_col) or "").split(",")[0])
            if _CUSTOMER_RE.search(addr) or "8975 guilford" in addr:
                addr = ""
            zp = re.sub(r"[^0-9]", "", str(r.get(zip_col) or ""))[:5]
            key = (core, addr, zp)
            if key in seen:
                continue
            seen.add(key)
            display = re.sub(r"\s*\(.*?\)\s*$", "", name).strip()
            self.entries.append({"name": name, "display": display, "core": core,
                                 "addr": addr, "zip": zp, "compact": _compact(core),
                                 "first": core.split(" ")[0]})

    def score(self, zone_core, body_core, zone_addr, domains, file_compact, file_core=None,
              logo_core=None):
        out = []
        for e in self.entries:
            sc, why = 0, []
            pat = r"(?<![a-z0-9])" + re.escape(e["core"]) + r"(?![a-z0-9])"
            if re.search(pat, zone_core):
                sc += 3
                why.append("name")
            elif re.search(pat, body_core):
                sc += 1
                why.append("name(body)")
            if len(e["compact"]) >= 5 and e["compact"] in file_compact:
                sc += 2
                why.append("filename")
            if len(e["compact"]) >= 6 and e["compact"] in domains:
                sc += 3
                why.append("web/email domain")
            elif len(e["first"]) >= 4 and e["first"] not in _GENERIC_CORES and any(
                    e["first"] in d for d in domains):
                sc += 2
                why.append("web/email domain")
                if re.search(r"(?<![a-z0-9])" + re.escape(e["first"]) + r"(?![a-z0-9])", zone_core):
                    sc += 1
            if file_core and "filename" not in why and (file_core == e["core"] or file_core.startswith(e["core"] + " ")):
                sc += 2
                why.append("filename vendor")
            if logo_core and len(e["first"]) >= 4 and e["first"] not in _GENERIC_CORES and \
                    re.search(r"(?<![a-z0-9])" + re.escape(e["first"]) + r"(?![a-z0-9])", logo_core) and "name" not in why:
                sc += 2
                why.append("logo word")
            if e["addr"] and len(e["addr"]) >= 6 and re.search(
                    r"(?<![a-z0-9])" + re.escape(e["addr"]) + r"(?![a-z0-9])", zone_addr):
                sc += 2
                why.append("remit address")
                if e["zip"] and e["zip"] in zone_addr:
                    sc += 1
                    why.append("zip")
            if sc:
                out.append((sc, e, why))
        out.sort(key=lambda t: (-t[0], -("name" in t[2]), -len(t[1]["core"])))
        return out


def _filename_vendor(stem):
    """'..._Invoice_6716_from_PRECISION_CARPET_SERVICES' -> 'PRECISION CARPET SERVICES'."""
    m = re.search(r"_from_(.+)$", stem, re.I)
    if not m:
        return None
    v = m.group(1)
    v = re.sub(r"(_[\d.][\w.]*)+$", "", v)
    v = re.sub(r"_MHIC\d+$", "", v, flags=re.I)
    v = re.sub(r"(?<=[A-Za-z])\d+$", "", v)
    v = v.replace("__", " & ").replace("_", " ")
    return _norm(v) or None


_LEAD_LABEL = re.compile(
    r"^(remit(\s*payment)?\s*to|mail\s*(check\s*)?to|send\s*payment\s*to|account\s*name|"
    r"payable\s*to|make\s*checks?\s*payable\s*to|attn|attention|please\s*remit\s*to|"
    r"lockbox\s*check\s*payment|payment\s*for)\s*:?\s*", re.I)
_TRAIL_JUNK = re.compile(
    r"\s+(federal\s*id|fed\s*id|tax\s*id|fein|ein|invoice|phone|tel|fax|www\.|"
    r"accounting\s*department|billing\s*team|\||-\s|\u2022).*$", re.I)


def _clean_vendor_text(t):
    t = _LEAD_LABEL.sub("", t.strip())
    t = _TRAIL_JUNK.sub("", t)
    t = re.split(r"\s{2,}|\s\|\s|\s\u2022\s", t)[0]
    return t.strip(" ,:-")


def _find_printed(lines, core):
    """Best printed form of a vendor whose core name we know."""
    pat = re.compile(r"(?<![a-z0-9])" + re.escape(core) + r"(?![a-z0-9])")
    best, best_rank = None, 9
    for ln in lines:
        for s in ln["segs"]:
            if len(s["text"]) > 90:
                continue
            c = _core(s["text"])
            if not pat.search(c):
                continue
            t = _clean_vendor_text(s["text"])
            ct = _core(t)
            if ct == core:
                rank = 0 if _SUFFIX_RE.search(t) else 1
            elif ct.startswith(core):
                rank = 2
                # cut after core + suffix
                m = re.search(r"(?i)(" + r"\W+".join(map(re.escape, core.split())) +
                              r")(\W+(?:" + _SUFFIX_ALT + r"))?", t.replace("'", ""))
                if m:
                    t = t[:m.end()] if m.end() <= len(t) else t
            else:
                rank = 4
            if rank < best_rank and not _CUSTOMER_RE.search(t):
                best, best_rank = t.strip(" ,:-"), rank
    return best if best_rank <= 2 else None


def vendor_zone_lines(pages, customer_texts):
    """Lines where a vendor identifies itself: top 30% / bottom 20% of page 1,
    remit/payment blocks, and letter sign-offs -- minus customer blocks."""
    if not pages or not pages[0]:
        return []
    out = []
    p = pages[0]
    h = p[0]["page_h"] or 792

    def ok(t):
        if _CUSTOMER_RE.search(t) or re.match(r"(?i)\s*(contact|attn|attention)\b", t):
            return False
        tl = t.strip().lower()
        return not any(tl and tl in ct for ct in customer_texts)

    for l in p:
        if l["y0"] < 0.30 * h or l["y1"] > 0.80 * h:
            segs = [s for s in l["segs"] if ok(s["text"])]
            if segs:
                out.append(dict(l, segs=segs, text="   ".join(s["text"] for s in segs)))
    for page in pages:
        for i, l in enumerate(page):
            if re.search(r"(remit|mail to|send payment|payable to|lockbox|payment mailing|"
                         r"sincerely|account name|check remittance)", l["text"], re.I):
                for x in page[i:i + 5]:
                    segs = [s for s in x["segs"] if ok(s["text"])]
                    if segs:
                        out.append(dict(x, segs=segs, text="   ".join(s["text"] for s in segs)))
    return out


def identify_vendor(pages, lines, text_all, stem, vendor_index, customer_texts,
                    pdf_path=None, notes=None):
    """-> (printed vendor, master row name or None, how)."""
    notes = notes if notes is not None else []
    zl = vendor_zone_lines(pages, customer_texts)
    zone_text = "\n".join(l["text"] for l in zl)
    zone_core = _core(zone_text.replace("\n", " \n "))
    body_lines = [l["text"] for l in lines if not _CUSTOMER_RE.search(l["text"])
                  and not any(l["text"].strip().lower() in ct for ct in customer_texts)]
    body_core = _core("\n".join(body_lines))
    zone_addr = _norm_addr(zone_text)
    domains = set()
    for d in re.findall(r"(?:@|www\.|https?://)([A-Za-z0-9.\-]+)\.(?:com|net|org|us|co|ai|biz)", text_all, re.I):
        for part in d.lower().split("."):
            part = part.replace("-", "")
            if part and part not in _GENERIC_DOMAINS and len(part) >= 4:
                domains.add(part)
    file_compact = _compact(stem)
    file_v = _filename_vendor(stem)

    best = None
    if vendor_index is not None:
        scored = vendor_index.score(zone_core, body_core, zone_addr, domains, file_compact,
                                    file_core=_core(file_v) if file_v else None)
        if scored and scored[0][0] >= 3:
            best = scored[0]

    # printed company names in the vendor zone
    printed_cands = []
    for l in zl:
        for s in l["segs"]:
            t = _clean_vendor_text(s["text"])
            if not t or len(t) > 60 or _GENERIC_TOP.match(t) or DATE_RE.search(t):
                continue
            if re.search(r"@|www\.|http|\d{3}[-.)]\s?\d{3}[-.]\d{4}|^\d", t):
                continue
            if _COMPANY_HINT.search(t) and not _STREET_RE.match(t) and not _CITY_RE.search(t):
                printed_cands.append(t)

    if best is not None and ("name" in best[2] or "filename" in best[2] or "web/email domain" in best[2]):
        sc, e, why = best
        printed = _find_printed(zl, e["core"]) or _find_printed(lines, e["core"])
        if not printed and "name" not in why:
            same = [t for t in printed_cands if _core(t).split(" ")[0] == e["first"]]
            if same:
                notes.append(f"printed name '{same[0]}' differs from master vendor '{e['display']}' "
                             f"(matched by {'+'.join(why)})")
                return same[0], e["name"], "+".join(why) + "+printed"
        return printed or e["display"], e["name"], "+".join(why)
    if file_v:
        printed = _find_printed(lines, _core(file_v))
        master = best[1]["name"] if best else None
        return printed or file_v, master, "filename" + ("+printed" if printed else "") + \
            ("+" + "+".join(best[2]) if best else "")
    if printed_cands:
        t = printed_cands[0]
        master = None
        if best is not None:
            master = best[1]["name"]
            notes.append(f"remit address matches master vendor '{best[1]['display']}' but the printed name is '{t}'")
        return t, master, "printed company line"
    if best is not None:
        sc, e, why = best
        return e["display"], e["name"], "+".join(why)
    # logo text: OCR the top of page 1
    if pdf_path is not None:
        try:
            hdr = ocr_page_header(pdf_path)
            hdr_text = "\n".join(l["text"] for l in hdr)
            if hdr_text.strip():
                if vendor_index is not None:
                    hc = _core(hdr_text.replace("\n", " \n "))
                    scored = vendor_index.score(hc + " \n " + zone_core, body_core, zone_addr,
                                                domains, file_compact, logo_core=hc)
                    if scored and scored[0][0] >= 3:
                        sc, e, why = scored[0]
                        return (_find_printed(hdr, e["core"]) or e["display"], e["name"],
                                "logo OCR+" + "+".join(why))
                for l in hdr:
                    t = _clean_vendor_text(l["segs"][0]["text"])
                    cm = re.search(r"\s+[A-Z][a-z]+,\s*[A-Z]{2}\s*\d{5}.*$", t)
                    if cm:
                        t = t[:cm.start()].strip()
                    if 3 <= len(t) <= 50 and not _GENERIC_TOP.match(t) and not _CUSTOMER_RE.search(t) \
                            and re.search(r"[A-Za-z]{3}", t) and not DATE_RE.search(t):
                        notes.append("vendor read from the logo by OCR -- check spelling")
                        return t, None, "logo OCR (raw)"
        except Exception as exc:
            notes.append(f"logo OCR failed: {exc}")
    return None, None, None


def vendor_address(pages, vendor):
    for page in pages[:3]:
        for i, l in enumerate(page):
            for seg in l["segs"]:
                m = re.search(r"(remit(?:\s*payment)?\s*to|mail\s*(?:check\s*)?to|send\s*payment\s*to|"
                              r"payment\s*mailing\s*address|lockbox\s*check\s*payment|please\s*remit\s*to|"
                              r"check\s*remittance)\s*:?", seg["text"], re.I)
                if not m:
                    continue
                rest = seg["text"][m.end():].strip(" :")
                parts = [rest] if rest else []
                if not (rest and _CITY_RE.search(rest)):
                    x0 = seg["x0"]
                    for nl in page[i + 1:i + 6]:
                        s2 = min(nl["segs"], key=lambda s: abs(s["x0"] - x0))
                        if abs(s2["x0"] - x0) > 60:
                            continue
                        t = s2["text"]
                        if re.match(r"(?i)(making payment|please|account|bank|aba|ach)", t):
                            continue
                        parts.append(t)
                        if _CITY_RE.search(t):
                            break
                joined = ", ".join(p for p in parts if p)
                joined = re.sub(r"\s*\u2022\s*", ", ", joined)
                joined = ", ".join(p for p in joined.split(", ") if not re.fullmatch(r"[ADFT]{15,}", p.strip()))
                if joined and (_CITY_RE.search(joined) or re.search(r"po box|p\.o\. box", joined, re.I)):
                    return joined[:150]
    if vendor:
        key = vendor.lower()[:12]
        for i, l in enumerate(pages[0]):
            seg = next((s for s in l["segs"] if key in s["text"].lower()), None)
            if seg is None:
                continue
            parts = []
            for nl in pages[0][i + 1:i + 4]:
                s2 = min(nl["segs"], key=lambda s: abs(s["x0"] - seg["x0"]))
                if abs(s2["x0"] - seg["x0"]) > 40:
                    break
                parts.append(s2["text"])
                if _CITY_RE.search(s2["text"]):
                    break
            j = ", ".join(parts)
            if _CITY_RE.search(j) and not _CUSTOMER_RE.search(j):
                return j[:150]
    return None


# --------------------------------------------------------------------------
# 6. LINE ITEMS
# --------------------------------------------------------------------------
_HDR_WORDS = {
    "desc": r"description|item description|product or service|service|activity|memo|"
            r"description/item ordered|particulars|item desc|description/libc",
    "qty": r"qty|quantity|ordered|shipped|hours|units|qty\.|freq exch qty",
    "price": r"rate|price|unit price|price each|unit cost|net price",
    "amount": r"amount|total|extension|ext\.? price|extended price|line total|"
              r"extended amount|amount \(usd\)|net charges",
    "item": r"item|item number|item no|stock number|product|sku|material|item code",
}
_END_ITEMS = re.compile(
    r"^(sub[\s-]?total|subtotal|total|tax\b|sales tax|freight|balance|invoice total|"
    r"invoice balance|amount due|ways to pay|payments?/credits|shipping & handling|"
    r"discount|thank you|please|terms|lines total|product category summary|"
    r"invoice sub-total|current\s+period|signature|we accept returns|total due|"
    r"special programs|credit amount|grand total|page \d|nonstock|description\s+total)", re.I)
_END_ANY = re.compile(r"\b(lines total|invoice total|sub[\s-]?total\s*:|total due|"
                      r"product category summary)\b", re.I)


def _header_line(page):
    best = None
    for li, ln in enumerate(page):
        texts = [s["text"].lower() for s in ln["segs"]]
        found = {}
        for key, pat in _HDR_WORDS.items():
            rx = re.compile(r"^(?:" + pat + r")\b", re.I)
            for s, t in zip(ln["segs"], texts):
                if rx.match(t) and key not in found and all(s is not v for v in found.values()):
                    found[key] = s
        ok = ("amount" in found and ("desc" in found or "item" in found)) or (
            "price" in found and ("desc" in found or "item" in found) and len(found) >= 3)
        if ok and len(ln["segs"]) >= 2:
            if best is None or len(found) > len(best[1]):
                best = (li, found)
                if len(found) >= 4:
                    return best
    return best


def _fee_lines_follow(page, ln):
    try:
        k = page.index(ln)
    except ValueError:
        return False
    for nl in page[k + 1:k + 3]:
        segs = nl["segs"]
        if len(segs) >= 3 and re.fullmatch(r"[A-Z]", segs[-1]["text"].strip()) and \
                parse_money(segs[-2]["text"]) is not None and \
                re.fullmatch(r"[A-Z][A-Z &/.-]{3,40}", segs[0]["text"].strip()) and \
                not re.match(r"(?i)sub|total|tax", segs[0]["text"].strip()):
            return True
    return False


_NOT_ITEM = re.compile(
    r"(?i)^(sub[\s-]?total|total|tax|sales tax|balance|amount|invoice|payment|terms|due|"
    r"current|past due|previous|credit|discount|paid|please|remit|customer|account|date|"
    r"billing date|net|page|contract amount|deposit|late)\b")


def _fallback_items(pages, subtotal, amount, tax):
    """No item table found: take 'description ... $amount' lines that add up
    to the subtotal/total; else one summary line with the invoice's main
    description so every invoice has at least one line."""
    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    sub, amt, tx = num(subtotal), num(amount), num(tax) or 0.0
    targets = [x for x in (sub, amt, (amt - tx) if amt is not None else None) if x is not None]
    cands = []
    for pi, page in enumerate(pages):
        h = page[0]["page_h"] if page else 792
        for ln in page:
            if pi == 0 and ln["y0"] < 0.22 * h:
                continue
            segs = ln["segs"]
            if len(segs) < 2:
                continue
            last = segs[-1]["text"].strip().rstrip("*").strip()
            if not re.fullmatch(r"\(?-?\$?\s?[\d,]*\.\d{2}\)?", last):
                continue
            desc_segs = [sg for sg in segs[:-1] if re.search(r"[A-Za-z]{3}", sg["text"])]
            if not desc_segs:
                continue
            desc = " ".join(sg["text"] for sg in desc_segs).strip()
            if _NOT_ITEM.match(desc) or re.search(r"(?i)total|balance due|amount due|see overage|\btax\b", desc) \
                    or _CUSTOMER_RE.search(desc):
                continue
            nums = [parse_money(sg["text"]) for sg in segs if parse_money(sg["text"]) is not None
                    and re.fullmatch(r"\(?-?\$?\s?[\d,]*\.?\d+\)?", sg["text"].strip())]
            qty = None
            if len(nums) >= 3 and nums[-2] and abs(nums[-3] * nums[-2] - nums[-1]) < 0.02:
                qty = nums[-3]
            elif len(nums) >= 2 and nums[0] and abs(nums[0] * (nums[-2] or 0) - nums[-1]) < 0.02:
                qty = nums[0]
            cands.append({"page": pi + 1, "item_code": None, "description": desc[:200], "qty": qty,
                          "unit_price": nums[-2] if len(nums) >= 2 else None,
                          "amount": parse_money(last)})
    if cands and any(abs(sum(c["amount"] or 0 for c in cands) - t) < 0.05 for t in targets):
        return cands
    # one summary line
    base = sub if sub is not None else (amt - tx if amt is not None else None)
    if base is None:
        return []
    desc = None
    for page in pages[:1]:
        h = page[0]["page_h"] if page else 792
        for ln in page:
            t = ln["text"].strip()
            if ln["y0"] < 0.25 * h or len(t) < 12 or _NOT_ITEM.match(t) or money_in(t) or DATE_RE.search(t):
                continue
            if re.search(r"(?i)service|maintenance|agreement|contract|repair|labor|inspection|monthly|"
                         r"rental|cleaning|install|replace|subscription|fee", t):
                desc = t[:200]
                break
    return [{"page": 1, "item_code": None,
             "description": "(no itemised table on the invoice) " + (desc or "invoice charges"),
             "qty": None, "unit_price": None, "amount": base}]


def extract_line_items(pages):
    items = []
    prev_cols = None
    finished = False
    for pi, page in enumerate(pages):
        if finished:
            break                         # the item table already ended at a total line
        hdr = _header_line(page)
        if not hdr:
            if prev_cols is None or not items:
                continue
            hdr = (-1, prev_cols)          # continuation page without a header row
        hi, cols = hdr
        prev_cols = cols
        has_qty = "qty" in cols or "price" in cols
        cur = None
        started = False
        for ln in page[hi + 1:]:
            segs = list(ln["segs"])
            flagged = False
            while len(segs) > 1 and re.fullmatch(r"[A-Z]", segs[-1]["text"].strip()):
                segs = segs[:-1]          # trailing tax flag (Y/N/T)
                flagged = True
            t = " ".join(s["text"] for s in segs).strip()
            first = segs[0]["text"].strip()
            if re.match(r"(?i)tax\s*\+", t):
                continue                  # per-line tax note inside the table (Chadwell)
            if started and (re.match(r"(?i)gl code\s*/\s*category", first) or (any(
                    re.fullmatch(r"(?i)sub[\s-]?total:?", s["text"].strip()) for s in segs)
                    and not _fee_lines_follow(page, ln))):
                finished = True
                break
            if _END_ITEMS.match(first) or _END_ITEMS.match(t) or _END_ANY.search(t):
                if started:
                    if re.match(r"(?i)sub[\s-]?total\b", first) and _fee_lines_follow(page, ln):
                        continue          # intermediate subtotal, fee lines follow (Cintas)
                    finished = bool(re.search(r"(?i)total|balance|amount\s*due|credit\s*amount", t)) \
                        and not re.search(r"(?i)continued", t)
                    break
                continue
            nums = []
            lead_code = None
            if "item" in cols and re.fullmatch(r"\d{3,}", segs[0]["text"].strip()) and \
                    _x_overlap(cols["item"]["x0"] - 25, cols["item"]["x1"] + 25, segs[0]["x0"], segs[0]["x1"]) > 0 \
                    and len(segs) >= 3:
                lead_code = segs[0]
            for s in segs:
                if s is lead_code:
                    continue
                st = s["text"].strip()
                v = parse_money(st)
                if v is not None and re.search(r"\d", st) and not DATE_RE.fullmatch(st) and \
                        re.fullmatch(r"\(?-?\$?\s?[\d,]*\.?\d+\)?-?(\s?(CR|EA|ea|[TNY*]))?", st):
                    nums.append((s, v))
            last_txt = segs[-1]["text"].rstrip("TNY *")
            money_last = bool(nums) and nums[-1][0] is segs[-1] and re.search(r"\.\d{2}\)?$", last_txt)
            text_segs = [s for s in segs if all(s is not n[0] for n in nums)]
            desc_txt = " ".join(s["text"] for s in text_segs).strip()
            is_item = bool(money_last) and (len(nums) >= 2 or not has_qty) and bool(desc_txt or len(nums) >= 3)
            if not is_item and started and flagged and money_last and len(nums) == 1 and desc_txt \
                    and re.fullmatch(r"[A-Z][A-Z &/.-]{3,40}", desc_txt):
                is_item = True            # fee line such as 'SERVICE CHARGE 6.93 N'
            if is_item and re.search(r"(?i)subtotal|^tax\b", desc_txt):
                is_item = False
            if is_item:
                started = True
                amount = nums[-1][1]
                price = nums[-2][1] if len(nums) >= 2 else None
                qty = None
                if "qty" in cols:
                    q = cols["qty"]
                    for s, v in nums[:-1]:
                        if len(nums) >= 2 and s is nums[-2][0] and "price" in cols:
                            continue
                        if _x_overlap(q["x0"] - 15, q["x1"] + 15, s["x0"], s["x1"]) > 0:
                            qty = v
                            break
                if len(nums) >= 3 and price not in (None, 0) and amount and \
                        abs(nums[-3][1] * price - amount) <= max(0.02, 0.01 * abs(amount)) and \
                        (qty is None or abs(qty * price - amount) > max(0.02, 0.01 * abs(amount))):
                    qty = nums[-3][1]
                if qty is None and len(nums) >= 3:
                    qty = nums[-3][1]
                if qty is None and price not in (None, 0) and amount:
                    r = amount / price
                    if abs(r - round(r)) < 0.02 and 0 < round(r) < 100000:
                        qty = float(round(r))
                if "price" not in cols and qty is not None and len(nums) == 2 and price == qty:
                    price = None          # 'qty  amount' layout: the 2nd-last number IS the qty
                if qty is None and len(nums) == 2 and "qty" in cols:
                    # 'qty  amount' layouts with no price column
                    qty, price = price, None
                code = lead_code["text"].strip() if lead_code is not None else None
                if code is None and "item" in cols and "desc" in cols and cols["item"] is not cols["desc"]:
                    it = cols["item"]
                    for s in text_segs:
                        if _x_overlap(it["x0"] - 8, it["x1"] + 30, s["x0"], s["x1"]) > 0 and s["x1"] < cols["desc"]["x0"] + 5:
                            code = s["text"]
                            break
                desc = " ".join(s["text"] for s in text_segs if s["text"] != code
                                and not re.fullmatch(r"[A-Z]", s["text"].strip())).strip()
                cur = {"page": pi + 1, "item_code": code, "description": desc,
                       "qty": qty, "unit_price": price, "amount": amount}
                items.append(cur)
            elif cur is not None and not nums and t and len(t) < 120:
                if re.match(r"(?i)^(gl code|receipt sku|contract:|equipment:|serial|model:|location:)", t):
                    cur["description"] = (cur["description"] + " | " + t).strip(" |")[:200]
                elif not re.match(r"(?i)^(hourly labor|fee for)", t) or True:
                    cur["description"] = (cur["description"] + " " + t).strip()
    return items


# --------------------------------------------------------------------------
# 7. DOCUMENT CLASS
# --------------------------------------------------------------------------
def classify(text_all, pages):
    tl = text_all.lower()
    head = "\n".join(l["text"] for l in pages[0][:14]).lower() if pages and pages[0] else tl[:600]
    if "check request form" in tl:
        return "check_request", "Check Request"
    if re.search(r"\bcredit memo\b", head) or "credit - do not pay" in tl:
        return "credit_memo", "Credit Memo"
    if re.search(r"(shipping receipt|sales receipt|\breceipt\b)", head) and not re.search(r"\binvoice\b", head):
        return "receipt", "Receipt"
    if re.search(r"(is \d+ days past due|reminder that your invoice|past due notice|"
                 r"overdue notice|final notice)", tl):
        return "past_due_notice", "Overdue / Late Notice"
    if re.search(r"disconnect(ion)? notice", tl):
        return "notice", "Disconnection Notice"
    if re.search(r"^(monthly statement|statement of account|statement|open invoice detail)\b", head, re.M) \
            or re.search(r"\bstmt date\b", head):
        return "statement", "Statement of Account" if "statement of account" in head else "Statement"
    if re.search(r"dear (valued )?(client|customer)", tl) and not re.search(
            r"(total due|amount due|balance due|invoice total)", tl):
        return "notice", "Notice"
    return "invoice", "Standard Vendor Invoice"


_UTILITY_WORDS = re.compile(
    r"\b(kwh|therms?|ccf|hcf|meter (number|reading)|electric (service|supply|delivery)|"
    r"gas (service|supply|delivery)|water (and|&) sewer|sewer charge|usage|"
    r"supply charges|delivery charges|distribution charge|service period|billing period)\b", re.I)


# --------------------------------------------------------------------------
# 8. FIELD EXTRACTION FOR ONE INVOICE (a group of pages)
# --------------------------------------------------------------------------
INV_NO_LABELS = (
    r"invoice\s*(?:no\.?|number|num\.?|#)|inv\s*(?:no\.?|#|number)|invoice(?=\s*:)|"
    r"credit memo\s*#|invoice(?=\s+[A-Z]{0,4}-?\d{3,})|reference\s*#|"
    r"document\s*(?:no\.?|number)|bill\s*(?:no\.?|number)")
DATE_LABELS_STRONG = (r"invoice\s*date|inv\.?\s*date|date\s*of\s*invoice|billing\s*date|"
                      r"bill\s*date|document\s*date|statement\s*date")
DATE_LABELS_WEAK = r"date(?!\s*(?:due|ordered|requested|shipped|range|of service))"
DUE_LABELS = (r"due\s*date|payment\s*due(?:\s*date)?|date\s*due|please\s*pay\s*before|"
              r"pay\s*by|due\s*by|due(?=\s*:)")
TOTAL_LABELS = [
    (r"total\s*amount\s*due", 1), (r"amount\s*due", 1), (r"total\s*due", 1),
    (r"balance\s*due", 2), (r"invoice\s*total", 2), (r"invoice\s*amount", 3),
    (r"total\s*usd", 2), (r"current\s*charges\s*due", 3),
    (r"please\s*pay", 3), (r"invoice\s*balance", 4), (r"grand\s*total", 3),
    (r"total(?!\s*(?:balance|merchandise|lease|rental|current|qty|lines|amount\s*remitted))", 5),
    (r"balance(?!\s*(?:forward|owed|brought))", 6),
]
_BAD_TOTAL_CTX = re.compile(r"(past\s*due|account\s*balance|previous|balance\s*forward|"
                            r"sub[\s-]?total|aging|prior|ytd|spend)", re.I)
_AGING_LINE = re.compile(r"(\d+\s*(-|to|–)\s*\d+\s*days|over\s*9\d|\bcurrent\b.*\b30\b|overdue)", re.I)


def _total_candidates(pages, statement=False):
    cands = []
    for rx, prio in TOTAL_LABELS:
        for v, info in find_values(pages, rx, v_money, max_below=3):
            lt = info["line_text"]
            if _BAD_TOTAL_CTX.search(info["label"]) or _BAD_TOTAL_CTX.search(info["seg"][:30]):
                continue
            if not statement and info["how"] == "below" and _AGING_LINE.search(lt):
                continue
            if not statement:
                prev = pages[info["page"]][info["line"] - 1]["text"] if info["line"] else ""
                if _AGING_LINE.search(lt) or _AGING_LINE.search(prev):
                    continue
            cands.append((prio, info["page"], info["y"], v, info))
    return cands


def _pick_total(cands, prefer_last=True):
    if not cands:
        return None, None
    best_prio = min(c[0] for c in cands)
    top = [c for c in cands if c[0] == best_prio]
    top.sort(key=lambda c: (c[1], c[2]))
    c = top[-1] if prefer_last else top[0]
    return c[3], c[4]


def extract_fields(pages, stem, vendor_index=None, pdf_path=None, doc_class=None,
                   doc_type=None):
    notes = []
    lines = _all_lines(pages)
    text_all = "\n".join(l["text"] for l in lines)
    if doc_class is None:
        doc_class, doc_type = classify(text_all, pages)
    f = {"doc_class": doc_class, "doc_type": doc_type}

    # ---- invoice number -------------------------------------------------
    inv_hits = find_values(pages, INV_NO_LABELS, v_invoice_number, mid=True)
    inv_hits = [h for h in inv_hits if not re.match(
        r"(?i)invoice\s*(date|total|amount|balance|summary|details|due|remarks)", h[1]["label"] + " " + h[1]["seg"][len(h[1]["label"]):][:10])]
    stem_nums = set(re.findall(r"[A-Z]{0,5}-?\d{4,}(?:-\d+)?", stem, re.I))
    inv = inv_hits[0][0] if inv_hits else None
    if inv_hits and stem_nums:
        for v, info in inv_hits:
            if any(_compact(v).lstrip("0") == _compact(s).lstrip("0") or
                   _compact(s).lstrip("0").endswith(_compact(v).lstrip("0")) and len(v) >= 5
                   for s in stem_nums):
                inv = v
                break
    if inv is None:
        oh = find_values(pages, r"order\s*#|order\s*number", v_invoice_number)
        if oh:
            inv = oh[0][0]
            notes.append("no 'Invoice #' label; used the Order # as invoice number")
    f["invoice_number"] = inv
    f["_inv_hits"] = [v for v, _ in inv_hits]

    # ---- dates ------------------------------------------------------------
    dh = find_values(pages, DATE_LABELS_STRONG, v_date, page_limit=2, mid=True)
    inv_date = dh[0][0] if dh else None
    if inv_date is None:
        dw = find_values(pages, DATE_LABELS_WEAK, v_date, page_limit=2)
        inv_date = dw[0][0] if dw else None
    if inv_date is None:
        for l in pages[0][:25]:
            d = parse_date(l["text"])
            if d:
                inv_date = d
                notes.append("invoice date has no label; took the first date printed near the top")
                break
    f["invoice_date"] = inv_date
    due_hits = find_values(pages, DUE_LABELS, v_date, page_limit=2, mid=True)
    f["due_date"] = due_hits[0][0] if due_hits else None
    if not f["due_date"] and re.search(r"due\s*upon\s*receipt|due\s*on\s*receipt", text_all, re.I):
        f["terms_note"] = "Due upon receipt"

    # ---- totals ---------------------------------------------------------------
    cands = _total_candidates(pages, statement=(doc_class == "statement"))
    if doc_class == "credit_memo":
        ca = find_values(pages, r"credit\s*amount", v_money, max_below=1)
        if ca:
            cands = [(0, ca[0][1]["page"], ca[0][1]["y"], -abs(ca[0][0]), ca[0][1])] + cands
    amount, amount_info = _pick_total(cands)
    sub_hits = find_values(pages, r"(?:invoice\s*)?sub[\s-]?total|subtotal", v_money, max_below=3)
    sub_hits = [h for h in sub_hits
                if not re.match(r"(?i)\s*(invoice\s*)?sub[\s-]?total\s*(labor|materials?|other|parts)", h[1]["seg"])]
    subtotal = sub_hits[-1][0] if sub_hits else None
    tax_hits = find_values(pages, r"(?:sales|state|total)?\s*tax(?:es)?(?:\s*\([\d.]+%\))?(?:\s*amount)?", v_money, max_below=2)
    tax_hits = [h for h in tax_hits if not re.search(r"tax\s*(id|code)", h[1]["seg"], re.I)
                and not re.match(r"(?i)tax\s*\+", h[1]["seg"])]
    tax = tax_hits[-1][0] if tax_hits else None
    if tax is None:
        th2 = find_values(pages, r"sales\s*tax", v_money, max_below=2, mid=True)
        th2 = [h for h in th2 if not re.search(r"tax\s*(id|code|exempt)", h[1]["seg"], re.I)]
        tax = th2[-1][0] if th2 else None
    fr_hits = find_values(pages, r"freight(?:/misc|/other)?|shipping(?:\s*&\s*handling)?(?=\s*[:$\d])", v_money, max_below=2)
    freight = fr_hits[-1][0] if fr_hits else None
    disc_hits = find_values(pages, r"discount(?:\s*amount)?", v_money, max_below=1)
    discount = disc_hits[-1][0] if disc_hits else None
    f.update({"subtotal": fmt_money(subtotal), "tax": fmt_money(tax), "freight": fmt_money(freight),
              "discount": fmt_money(abs(discount)) if discount else None})

    # ---- PO / account ----------------------------------------------------------
    po_hits = find_values(pages, r"p\.?\s?o\.?\s*(?:number|no\.?|#)|purchase\s*order(?:\s*number)?|"
                                 r"customer\s*p\.?o\.?\s*(?:#|no\.?)?|po\s*number|customer\s*order\s*number",
                          v_po, page_limit=2)
    f["po_number"] = po_hits[0][0] if po_hits else None
    acct_hits = find_values(pages, r"account\s*(?:no\.?|number|#)|customer\s*(?:account|number|no\.?|#|id)|"
                                   r"cust(?:omer)?\.?\s*#|client\s*id|bill\s*to\s*#|sold\s*to\s*#|"
                                   r"master\s*account\s*number|account(?=\s*:)|customer\s*no",
                            v_account, page_limit=2, mid=True)
    acct = None
    for v, info in acct_hits:
        ln = pages[info["page"]][info["line"]]
        prev = pages[info["page"]][info["line"] - 1]["text"] if info["line"] else ""
        if re.search(r"(routing|bank|aba\b|ach\b|wire|swift|bic\b)", ln["text"] + " " + prev, re.I):
            continue
        acct = v
        break
    f["account_number"] = acct

    # ---- addresses / property -------------------------------------------------
    ship, ship_lab = address_block(pages, r"ship[\s-]*to(?![\s-]*(?:id|name)\b)|service\s*(?:at|address|location)|"
                                          r"job\s*(?:address|site|location)|location|compactor\s*location|"
                                          r"deliver\s*to|property\s*address|shipping\s*address")
    bill, bill_lab = address_block(pages, r"bill\s*to\s*address|bill\s*to|sold\s*to|"
                                          r"customer\s*billing\s*address|to:")
    s_name, s_addr = split_block(ship)
    b_name, b_addr = split_block(bill)
    prop = None
    pm = find_values(pages, r"property\s*name|property\s*to\s*charge",
                     lambda t: _norm(t) if len(_norm(t)) > 3 and not re.match(r"(?i)property", t) else None,
                     page_limit=1)
    if pm:
        prop = clean_property(pm[0][0])
    for cand in (s_name, b_name):
        if not prop:
            prop = clean_property(cand)
    if not prop and s_addr:
        # unlabeled bill-to block: the name printed just above the ship-to street
        street = s_addr.split(",")[0].strip().lower()
        for page in pages[:1]:
            for i, l in enumerate(page):
                if i and l["segs"][0]["text"].strip().lower() == street:
                    cand = page[i - 1]["segs"][0]["text"].strip()
                    if not _STREET_RE.match(cand) and not _CITY_RE.search(cand) \
                            and not BLOCK_LABEL_START.match(cand) and not cand.endswith(":") \
                            and re.search(r"[A-Za-z]{3}", cand) and not _person_like(cand):
                        prop = clean_property(cand)
                        break
    if not prop:
        # unlabeled customer block right after the vendor's remit city line
        for page in pages[:1]:
            for i, l in enumerate(page):
                if re.match(r"(?i)remit\s*to", l["segs"][0]["text"]):
                    seen_city = False
                    for nl in page[i + 1:i + 8]:
                        t0 = nl["segs"][0]["text"].strip()
                        if _CITY_RE.search(t0):
                            seen_city = True
                            continue
                        if seen_city and not _STREET_RE.match(t0) and len(t0) > 3:
                            prop = clean_property(t0)
                            break
                    break
    if not prop:
        cn = re.search(r"Card\s*Nickname:\s*([A-Za-z][A-Za-z0-9 &'.-]{2,40}?)(?:\s{2,}|\s+Card#|$)", text_all, re.M)
        if cn and not _person_like(cn.group(1)):
            prop = clean_property(cn.group(1))
    if not prop:
        p0 = pages[0]
        for i, l in enumerate(p0[:60]):
            for sg in l["segs"]:
                if re.fullmatch(r"(?i)\s*(c/o\s*)?residential\s*one(\s*(llc|management))?\s*", sg["text"]):
                    for j in range(i - 1, max(-1, i - 5), -1):
                        above = [s2 for s2 in p0[j]["segs"] if abs(s2["x0"] - sg["x0"]) < 25]
                        if not above:
                            break
                        t2 = above[0]["text"].strip()
                        if _STREET_RE.match(t2) and _STREET_SUFFIX.search(t2) or _CITY_RE.search(t2):
                            continue
                        prop = clean_property(t2)
                        break
                if prop:
                    break
            if prop:
                break
    if not prop:
        for l in pages[0][:40]:
            for sg in l["segs"]:
                m = re.match(r"(?i)^\s*residential\s*one\s*[-\u2013]\s*(.{3,50})$", sg["text"]) or \
                    re.match(r"(?i)^\s*(.{3,50}?)\s*[-\u2013,]?\s*c/o\s*residential\s*one\s*$", sg["text"])
                if m and _PROP_WORDS.search(m.group(1)):
                    prop = clean_property(m.group(1))
                    break
            if prop:
                break
    f["property"] = prop
    f["service_address"] = s_addr or b_addr
    f["bill_to"] = "; ".join(bill) if bill else None
    f["ship_to"] = "; ".join(ship) if ship else None
    customer_texts = [b.lower() for b in (ship + bill)]
    if prop:
        customer_texts.append(prop.lower())

    # ---- vendor -------------------------------------------------------------------
    if doc_class == "check_request":
        payee = find_values(pages, r"name", lambda t: _norm(t) if re.search(r"[A-Za-z]{2,}", t) else None, page_limit=1)
        f["vendor"] = payee[0][0] if payee else None
        f["vendor_master_match"], f["vendor_how"] = None, "check-request payee"
    else:
        v, vm, how = identify_vendor(pages, lines, text_all, stem, vendor_index,
                                     customer_texts, pdf_path=pdf_path, notes=notes)
        f["vendor"], f["vendor_master_match"], f["vendor_how"] = v, vm, how
    f["vendor_address"] = vendor_address(pages, f.get("vendor"))

    # ---- line items -------------------------------------------------------------
    f["line_items"] = extract_line_items(pages)
    sub_docs = []

    # ---- document-type specifics ----------------------------------------------------
    flat = re.sub(r"\s+", " ", text_all)
    if doc_class == "past_due_notice":
        m = re.search(r"invoice\s+([A-Z0-9-]{4,}),?\s+dated\s+([\d/]+)\s+is past due with a balance of\s+\$?([\d,]+\.\d{2})", flat, re.I)
        if m:
            f["invoice_number"] = m.group(1)
            f["invoice_date"] = parse_date(m.group(2))
            amount = parse_money(m.group(3))
            notes.append("past-due reminder letter (not an invoice); invoice #, date and balance are those of the invoice it refers to")
        for i, l in enumerate(pages[0]):
            if re.match(r"(?i)c/o residential one", l["text"]) and i:
                f["property"] = clean_property(pages[0][i - 1]["text"])
    elif doc_class == "check_request":
        req = find_values(pages, r"date", v_date, page_limit=1)
        if req:
            f["invoice_date"] = req[0][0]
        tot = find_values(pages[:1], r"total", lambda t: parse_money(t), page_limit=1)
        if tot:
            amount = tot[0][0]
        for i, l in enumerate(pages[0]):
            if re.match(r"(?i)property to charge", l["text"]):
                cand = ([pages[0][i - 1]["text"]] if i else []) + [s["text"] for s in l["segs"][1:]]
                for c in cand:
                    if not re.search(r"requester|date|phone", c, re.I):
                        f["property"] = re.split(r"\s{3,}", c)[0]
                        break
        f["invoice_number"] = None
        items = []
        hdr_i = next((i for i, l in enumerate(pages[0]) if re.match(r"(?i)description\s+expense code", l["text"])), None)
        if hdr_i is not None:
            for l in pages[0][hdr_i + 1:]:
                if re.match(r"(?i)total", l["text"]):
                    break
                segs = l["segs"]
                if len(segs) >= 2 and parse_money(segs[-1]["text"]) is not None:
                    items.append({"page": 1, "item_code": segs[-2]["text"] if len(segs) >= 3 else None,
                                  "description": segs[0]["text"], "qty": None,
                                  "unit_price": None, "amount": parse_money(segs[-1]["text"])})
        f["line_items"] = items
        f["subtotal"] = f["tax"] = f["freight"] = None
        if len(pages) > 1:
            notes.append(f"check request form + {len(pages) - 1} page(s) of attached receipts (receipt totals not used)")
    elif doc_class == "statement":
        tb = find_values(pages, r"total\s*balance|total\s*amount\s*due|total\s*due|amount\s*due", v_money, max_below=3)
        if not tb:
            tb = find_values(pages, r"balance", v_money, max_below=3)
        tb = [h for h in tb if not re.search(r"past\s*due", h[1]["label"], re.I)]
        if tb:
            amount = tb[0][0]
        sd = find_values(pages, r"statement\s*date|stmt\s*date|date", v_date, page_limit=1)
        if sd:
            f["invoice_date"] = sd[0][0]
        seen = set()
        for l in lines:
            t = l["text"]
            m2 = re.search(r"INV\s*#\s*(\d+)\.\s*Due\s*([\d/]+)\.\s*Orig\.\s*Amount\s*\$([\d,]+\.\d{2})", t)
            m3 = re.match(r"^INV\s+(\d{6,})\s+\d+\s+.*?(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+\$?([\d,]+\.\d{2})", t)
            m4 = re.match(r"^(\d{1,2}/\d{1,2}/\d{2})\s+(\d{1,2}/\d{1,2}/\d{2})\s+\w+\s+\w+\s+([\w-]+)\s+(\(?[\d,]+\.\d{2}\)?)", t)
            m = re.match(r"^(\d{1,2}/\d{1,2}/\d{2,4})\s+(?:INV\s*#?\s*)?([A-Z]{0,4}-?\d{5,}(?:-\d+)?)\b", t)
            d = None
            if m2:
                d = (m2.group(1), parse_date(t), parse_date(m2.group(2)), parse_money(m2.group(3)))
            elif m3:
                d = (m3.group(1), parse_date(m3.group(2)), parse_date(m3.group(3)), parse_money(m3.group(4)))
            elif m4:
                d = (m4.group(3), parse_date(m4.group(1)), parse_date(m4.group(2)), parse_money(m4.group(4)))
            elif m:
                ms = money_in(t)
                d = (m.group(2), parse_date(m.group(1)), None, ms[-1] if ms else None)
            if d and d[0] not in seen and d[3] is not None:
                seen.add(d[0])
                sub_docs.append({"invoice_number": d[0], "invoice_date": d[1], "due_date": d[2], "amount": d[3]})
        f["invoice_number"] = None
        f["due_date"] = None
        notes.append(f"statement lists {len(sub_docs)} invoice(s); a statement has no single invoice number")
        f["line_items"] = [{"page": None, "item_code": d["invoice_number"],
                            "description": f"statement line: invoice {d['invoice_number']} dated {d['invoice_date']}"
                                           + (f", due {d['due_date']}" if d.get('due_date') else ""),
                            "qty": None, "unit_price": None, "amount": d["amount"]} for d in sub_docs]
        f["subtotal"] = f["tax"] = f["freight"] = None
    elif doc_class == "receipt":
        m = re.search(r"([\d,]+\.\d{2})\s*(?:AMOUNT\s*DU\s?E|TOTAL\b)", "\n".join(l["text"] for l in pages[0]))
        tl = [l["text"] for l in pages[0]]
        tots = [money_in(x)[0] for x in tl if re.search(r"amount\s*du\s?e", x, re.I) and money_in(x)]
        if tots:
            amount = tots[0]
        elif m:
            amount = parse_money(m.group(1))
        om = re.search(r"Ord#\s*:?\s*([\d/]+)", text_all)
        if om:
            f["invoice_number"] = om.group(1)
        f["subtotal"] = f["tax"] = f["freight"] = None
        f["po_number"] = None
        f["property"] = None
        if len(pages) > 1:
            notes.append(f"{len(pages)}-page receipt bundle; values shown are from the first receipt only")

    f["amount"] = fmt_money(amount)
    if not f["line_items"] and doc_class in ("invoice", "credit_memo"):
        f["line_items"] = _fallback_items(pages, f.get("subtotal"), f.get("amount"), f.get("tax"))

    # ---- utility-bill fields (only when the content really is a utility bill) --
    is_util = bool(len(_UTILITY_WORDS.findall(text_all)) >= 2 and re.search(
        r"\b(kwh|therms?|ccf|hcf|gallons|meter)\b", text_all, re.I))
    f["is_utility_bill"] = is_util
    sp = re.search(r"(?:maintenance|service|billing)\s*period\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:-|to|through)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_all, re.I) \
        or re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(\d{1,2}/\d{1,2}/\d{4})\s*Monthly Billing", text_all, re.I)
    if sp:
        f["service_period_start"] = parse_date(sp.group(1))
        f["service_period_end"] = parse_date(sp.group(2))
    if is_util:
        cc = find_values(pages, r"total\s*current\s*charges|current\s*charges|service\s*period\s*total", v_money)
        f["current_charges"] = fmt_money(cc[0][0]) if cc else f["amount"]
        lf = find_values(pages, r"late\s*(?:payment\s*)?(?:fee|charge)", v_money)
        f["late_fee"] = fmt_money(lf[0][0]) if lf else None
        tl = text_all.lower()
        f["utility_type"] = ("Electric" if "kwh" in tl else "Gas" if "therm" in tl
                             else "Water" if ("water" in tl or "sewer" in tl) else "Other")
    f["sub_documents"] = sub_docs
    f["notes"] = notes
    return f


# --------------------------------------------------------------------------
# 9. WHOLE PDF: split into invoices when one PDF holds several
# --------------------------------------------------------------------------
def _page_invoice_numbers(page):
    hits = find_values([page], INV_NO_LABELS, v_invoice_number, mid=True)
    hits = [h for h in hits if not re.match(r"(?i)invoice\s*(date|total|amount|balance|summary|details|due)", h[1]["label"])]
    return hits[0][0] if hits else None


def _group_pages(pages):
    groups, cur, cur_no = [], [], None
    for i, p in enumerate(pages):
        no = _page_invoice_numbers(p)
        if cur and no and cur_no and _compact(no).lstrip("0") != _compact(cur_no).lstrip("0"):
            groups.append((cur_no, cur))
            cur, cur_no = [], None
        cur.append(i)
        if no and not cur_no:
            cur_no = no
    if cur:
        groups.append((cur_no, cur))
    return groups


def extract_invoice_local(pdf_path, vendor_rows=None, vendor_index=None, max_pages=None):
    pdf_path = Path(pdf_path)
    if vendor_index is None and vendor_rows:
        vendor_index = VendorIndex(vendor_rows)
    pages, ocr_pages = read_pdf_lines(pdf_path, max_pages=max_pages)
    lines = _all_lines(pages)
    text_all = "\n".join(l["text"] for l in lines)
    if len(re.sub(r"\s", "", text_all)) < 20:
        return {"doc_class": "unreadable", "doc_type": None, "line_items": [],
                "notes": ["no readable text, even after OCR"], "ocr_used": bool(ocr_pages),
                "pages": len(pages), "sub_documents": [], "doubts": ["unreadable PDF"]}
    doc_class, doc_type = classify(text_all, pages)
    stem = pdf_path.stem
    groups = _group_pages(pages) if doc_class in ("invoice", "credit_memo") and len(pages) > 1 else [(None, list(range(len(pages))))]
    if len(groups) > 1:
        results = []
        for no, idx in groups:
            sub_pages = [pages[i] for i in idx]
            sub_text = "\n".join(l["text"] for p in sub_pages for l in p)
            c, t = classify(sub_text, sub_pages)
            r = extract_fields(sub_pages, stem, vendor_index, pdf_path, c, t)
            r["page_range"] = f"{idx[0] + 1}-{idx[-1] + 1}" if len(idx) > 1 else f"{idx[0] + 1}"
            for it in r["line_items"]:
                if it.get("page"):
                    it["page"] = idx[0] + it["page"]
            results.append(r)
        stem_c = _compact(stem)
        main = None
        for r in results:
            n = _compact(r.get("invoice_number") or "").lstrip("0")
            if n and len(n) >= 5 and n in stem_c:
                main = r
                break
        main = main or results[0]
        others = [r for r in results if r is not main]
        f = dict(main)
        f["sub_documents"] = [{"invoice_number": r.get("invoice_number"), "invoice_date": r.get("invoice_date"),
                               "due_date": r.get("due_date"), "amount": float(r["amount"]) if r.get("amount") else None,
                               "pages": r.get("page_range"), "line_items": r.get("line_items")} for r in others]
        f["notes"] = list(main["notes"]) + [
            f"PDF holds {len(results)} separate invoices: " + "; ".join(
                f"#{r.get('invoice_number')} ${r.get('amount')} (p.{r.get('page_range')})" for r in results)
            + f" -- main row uses #{main.get('invoice_number')}; the others need their own entry"]
        f["multi_invoice"] = True
    else:
        f = extract_fields(pages, stem, vendor_index, pdf_path, doc_class, doc_type)
        f["multi_invoice"] = False
    if ocr_pages:
        f["notes"] = ["OCR used on page(s) " + ",".join(map(str, ocr_pages))] + f["notes"]
    f["ocr_used"] = bool(ocr_pages)
    f["pages"] = len(pages)
    f["doubts"] = check_fields(f)
    f.pop("_inv_hits", None)
    return f


def check_fields(f):
    doubts = []
    cls = f.get("doc_class")
    need = ["vendor", "invoice_number", "invoice_date", "amount"]
    if cls in ("statement", "check_request", "receipt"):
        need = ["vendor", "invoice_date", "amount"]
    elif cls == "notice":
        need = ["vendor"]
    for k in need:
        if not f.get(k):
            doubts.append(f"missing {k.replace('_', ' ')}")
    def num(k):
        try:
            return float(f[k]) if f.get(k) not in (None, "") else None
        except (TypeError, ValueError):
            return None
    amt, sub, tax, fr = num("amount"), num("subtotal"), num("tax") or 0.0, num("freight") or 0.0
    if amt is not None and sub is not None and cls in ("invoice", "credit_memo"):
        disc = 0.0
        dsc = num("discount") or 0.0
        if not any(abs(x - abs(amt)) < 0.02 for x in (abs(sub + tax), abs(sub + tax + fr), abs(sub), abs(sub + fr),
                                                     abs(sub + tax + fr - dsc))):
            doubts.append(f"total {amt:.2f} differs from subtotal {sub:.2f} + tax {tax:.2f} + freight {fr:.2f} (discount/fee/credit on invoice?)")
    items = f.get("line_items") or []
    if items and cls in ("invoice", "credit_memo"):
        s = round(sum(i.get("amount") or 0 for i in items), 2)
        targets = [v for v in (sub, amt, (amt or 0) - tax, (amt or 0) - tax - fr) if v is not None]
        if not any(abs(abs(s) - abs(x)) < 0.05 for x in targets):
            doubts.append(f"line items add up to {s:.2f}, not the subtotal/total (some lines may be missed or are category subtotals)")
    if f.get("vendor") and not f.get("vendor_master_match") and cls != "check_request":
        doubts.append("vendor not found in masters/vendors_data.csv")
    if not f.get("property") and cls not in ("notice",):
        doubts.append("no property / ship-to name found")
    if f.get("ocr_used"):
        doubts.append("scanned page(s) read by OCR -- verify values")
    if f.get("multi_invoice"):
        doubts.append("PDF holds more than one invoice")
    return doubts
