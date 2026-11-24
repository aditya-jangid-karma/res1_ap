"""Match a raw vendor name (as read off an invoice) against a master list of
correct vendor names, so we never write a typo or an inconsistent spelling
into the spreadsheet.

PLAIN ENGLISH — how a name gets matched (in this order, first hit wins):

  1. ALIAS   -- Someone already told us "when you see X, it really means Y"
                (this lives in masters\\aliases.csv). If the raw name matches
                a known alias, we trust it completely and use the correction.
  2. EXACT   -- The raw name (ignoring case/extra spaces) is already spelled
                exactly like a name in the vendor master. Safe to use as-is.
  3. CLOSE   -- The raw name is *similar* to one or more master names (a
                likely OCR/typing slip) but not identical. This is NOT
                trusted automatically -- a human must look at it. We still
                write down the RAW value we read (never a guess), but we
                also hand back our best guess as a "suggestion" so the
                reviewer doesn't have to search for it themselves. If more
                than one master name is close enough and we also know the
                invoice's address, we use the address to pick the more
                likely one (e.g. two vendors named almost the same thing but
                at different addresses).
  4. NONE    -- Nothing close enough was found. Also needs human review.

Only case 1 and case 2 are allowed to change what gets written to the
spreadsheet. Case 3 and 4 always keep the raw text and simply flag the row
for review -- this is a deliberate safety rule (see the design spec, item
D4: "non-exact = review, never auto-fill"). We would rather ask a human to
glance at ten flagged rows than silently write the wrong vendor once.

This file only depends on the Python standard library:
  - csv       to read the master and alias spreadsheets (as plain CSV files)
  - difflib   to measure how "similar" two names are (0.0 = nothing alike,
              1.0 = identical), which is how we detect likely typos
  - dataclasses  to define the small Match result object below
"""

import csv
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Optional

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False



def _normalise(name):
    """Lower-case, trim, and collapse inner spaces so that "A  B "  ==  "a b".
    This is the ONE normalisation rule used everywhere in this file, and it
    matches the rule already used elsewhere in this project."""
    return " ".join(str(name).lower().split())


def load_master(csv_path, name_col="Vendor Name", address_col="Address"):
    """Read a flattened master list (vendors, properties, etc.) into a plain
    list of dict rows -- one dict per CSV row, keyed by column header.

    `name_col` / `address_col` just tell us which columns hold the name and
    the address, so the SAME loader can later read a property master (or any
    other master) that uses different column names.

    If the file doesn't exist or can't be read, we return an empty list
    instead of crashing -- callers can treat "no master loaded yet" the same
    way they'd treat "nothing matched"."""
    rows = []
    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                rows.append(row)
    except (OSError, csv.Error):
        return []
    return rows


def load_aliases(csv_path):
    """Read masters\\aliases.csv (columns: raw,correct,count,last_seen) into
    a dict of {normalised raw name -> correct vendor name}.

    This is the "we've seen this exact mistake before and already know the
    fix" list. `count` and `last_seen` are just bookkeeping for a human
    reviewing the file later -- this loader doesn't need them.

    Same rule as load_master: a missing/unreadable file returns {} rather
    than raising, so an empty or not-yet-created aliases file is a normal,
    harmless starting state."""
    aliases = {}
    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                raw = row.get("raw")
                correct = row.get("correct")
                if raw and correct:
                    aliases[_normalise(raw)] = correct
    except (OSError, csv.Error):
        return {}
    return aliases


def load_alias_examples(csv_path, limit=5):
    """Read masters\\aliases.csv and hand back up to `limit` (raw, correct)
    pairs to use as "here's a mistake you've made before" examples in the
    Qwen prompt (see reader_qwen.build_prompt).

    Unlike load_aliases(), this keeps the ORIGINAL raw text exactly as typed
    (not lower-cased/collapsed) -- we want Qwen to see the same messy text a
    human actually saw, not our normalised lookup key.

    Which rows are "most useful" to show the model: we sort by `count`
    (how many times a human has had to fix this exact misread) DESCENDING,
    then by `last_seen` DESCENDING as a tiebreak, so the most frequent and
    most recent corrections are the ones we spend prompt space on. A blank
    or missing `count` is treated as 0 rather than crashing the sort.

    Rows with a blank raw or correct are skipped -- there's nothing useful
    to show the model. Same forgiving behaviour as the other loaders here: a
    missing/unreadable file returns [] instead of raising."""
    rows = []
    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                raw = row.get("raw")
                correct = row.get("correct")
                if not raw or not correct:
                    continue
                count_text = (row.get("count") or "").strip()
                try:
                    count = int(count_text)
                except ValueError:
                    count = 0
                last_seen = row.get("last_seen") or ""
                rows.append((count, last_seen, raw, correct))
    except (OSError, csv.Error):
        return []

    rows.sort(key=lambda r: (r[0], r[1]), reverse=True)
    return [(raw, correct) for _, _, raw, correct in rows[:limit]]


@dataclass
class Match:
    """The result of looking up one raw vendor name.

    value       -- what should actually be WRITTEN into the Vendor cell.
    status      -- "READ_OK" (safe to trust) or "NEEDS_REVIEW" (a human
                   should check this row before it's relied on).
    suggestion  -- our best guess at the correct name, shown to a human for
                   convenience. Only set when status is NEEDS_REVIEW because
                   of a close-but-not-exact match; otherwise None.
    how         -- which rule produced this result: "alias", "exact",
                   "close", or "none". Handy for debugging/logging.
    """
    value: str
    status: str
    suggestion: Optional[str]
    how: str
    code: Optional[str] = None


class Lookup:
    """Looks up raw vendor names against a master list, using the 4-step
    order described at the top of this file (alias -> exact -> close ->
    none)."""

    def __init__(self, master_rows, aliases=None, name_col="Vendor Name",
                 address_col="Address", code_col="ERP Vendor ID", threshold=0.90,
                 property_vendor_rows=None):
        """
        master_rows -- list of dict rows, e.g. straight from load_master().
        aliases     -- dict from load_aliases(), or None if there are no
                       known aliases yet.
        name_col / address_col -- which keys in each row dict hold the
                       vendor name and the address (lets this same class
                       serve a differently-shaped master later).
        code_col    -- the key for the ID/Code to return.
        threshold   -- how similar (0.0-1.0) a name has to be to count as a
                       "close" match. 0.90 is deliberately strict: this is a
                       safety net for typos, not a general fuzzy search.
        property_vendor_rows -- optional list of dict rows from property_vendor_map.csv
                       holding property-specific Intacct vs Yardi vendor codes.
        """
        self.aliases = aliases or {}
        self.threshold = threshold

        # Fast path for an EXACT match: normalised name -> (canonical name, code).
        self._exact = {}

        # Slower path used only for CLOSE matches / the address tiebreak.
        # One entry per master row: (normalised name, canonical name,
        # normalised address, code).
        self._candidates = []

        for row in master_rows:
            canonical = row.get(name_col)
            code = row.get(code_col)
            if not canonical:
                continue  # a blank/missing name in the master is unusable
            normalised_name = _normalise(canonical)
            self._exact.setdefault(normalised_name, (canonical, code))
            normalised_address = _normalise(row.get(address_col) or "")
            self._candidates.append(
                (normalised_name, canonical, normalised_address, code))

        # Secondary index: (normalised_vendor, normalised_property) or normalised_vendor -> {intacct, yardi}
        self._prop_vendor_map = {}
        if property_vendor_rows:
            for r in property_vendor_rows:
                v_norm = _normalise(r.get("Vendor Name") or "")
                p_norm = _normalise(r.get("Property Name") or "")
                p_code = _normalise(r.get("Property Code") or "")
                intacct_code = r.get("Intacct Code") or ""
                yardi_code = r.get("Yardi Code") or ""
                raw_code = r.get("Raw Vendor Code") or ""

                entry = {
                    "intacct": intacct_code,
                    "yardi": yardi_code,
                    "raw": raw_code
                }
                if v_norm and p_norm:
                    self._prop_vendor_map[(v_norm, p_norm)] = entry
                if v_norm and p_code:
                    self._prop_vendor_map[(v_norm, p_code)] = entry
                if v_norm and v_norm not in self._prop_vendor_map:
                    self._prop_vendor_map[v_norm] = entry

    def resolve_code_for_package(self, vendor_name, package=None, property_name=None, property_code=None, fallback_code=None):
        """Resolves ERP vendor code taking into account INTACCT (numbers) vs YARDI (alphanumeric)."""
        if not self._prop_vendor_map or not vendor_name:
            return fallback_code

        v_norm = _normalise(vendor_name)
        p_norm = _normalise(property_name) if property_name else None
        c_norm = _normalise(property_code) if property_code else None

        entry = None
        if p_norm and (v_norm, p_norm) in self._prop_vendor_map:
            entry = self._prop_vendor_map[(v_norm, p_norm)]
        elif c_norm and (v_norm, c_norm) in self._prop_vendor_map:
            entry = self._prop_vendor_map[(v_norm, c_norm)]
        elif v_norm in self._prop_vendor_map:
            entry = self._prop_vendor_map[v_norm]

        if not entry:
            return fallback_code

        pkg = str(package or "").upper()
        if "INTACCT" in pkg or "SAGE" in pkg:
            return entry.get("intacct") or entry.get("raw") or fallback_code
        elif "YARDI" in pkg:
            return entry.get("yardi") or entry.get("raw") or fallback_code
        return entry.get("raw") or fallback_code

    def match(self, raw_name, address=None, package=None, property_name=None, property_code=None):
        """Look up one raw vendor name (optionally with the invoice's
        address, used only to break ties in step 3) and return a Match."""

        # A blank name is not something we can look up at all.
        if raw_name is None or not str(raw_name).strip():
            return Match(value="", status="NEEDS_REVIEW",
                         suggestion=None, how="none", code=None)

        normalised_raw = _normalise(raw_name)

        # --- 1. ALIAS: someone already corrected this exact mistake before.
        if normalised_raw in self.aliases:
            canon = self.aliases[normalised_raw]
            # Try to get the code for the canonical name from self._exact
            norm_canon = _normalise(canon)
            code = self._exact.get(norm_canon, (None, None))[1]
            final_code = self.resolve_code_for_package(canon, package, property_name, property_code, fallback_code=code)
            return Match(value=canon,
                         status="READ_OK", suggestion=None, how="alias", code=final_code)

        # --- 2. EXACT: already spelled exactly like a master vendor.
        if normalised_raw in self._exact:
            canonical, code = self._exact[normalised_raw]
            final_code = self.resolve_code_for_package(canonical, package, property_name, property_code, fallback_code=code)
            return Match(value=canonical,
                         status="READ_OK", suggestion=None, how="exact", code=final_code)

        # --- 3. CLOSE: similar to one (or more) master names, but not
        # identical. Never trusted automatically -- always flagged.
        close_candidates = []
        for normalised_name, canonical, normalised_address, code in self._candidates:
            if HAS_RAPIDFUZZ:
                name_ratio = fuzz.ratio(normalised_raw, normalised_name) / 100.0
            else:
                name_ratio = SequenceMatcher(
                    None, normalised_raw, normalised_name).ratio()

            if name_ratio >= self.threshold:
                close_candidates.append(
                    (name_ratio, canonical, normalised_address, code))

        if close_candidates:
            if len(close_candidates) > 1 and address:
                normalised_input_address = _normalise(address)
                if HAS_RAPIDFUZZ:
                    best = max(
                        close_candidates,
                        key=lambda c: fuzz.ratio(normalised_input_address, c[2]) / 100.0)
                else:
                    best = max(
                        close_candidates,
                        key=lambda c: SequenceMatcher(
                            None, normalised_input_address, c[2]).ratio())
            else:
                best = max(close_candidates, key=lambda c: c[0])


            _, best_canonical, _, best_code = best
            final_code = self.resolve_code_for_package(best_canonical, package, property_name, property_code, fallback_code=best_code)
            return Match(value=raw_name, status="NEEDS_REVIEW",
                         suggestion=best_canonical, how="close", code=final_code)

        # --- 4. NONE: nothing in the master looks like this at all.
        return Match(value=raw_name, status="NEEDS_REVIEW",
                     suggestion=None, how="none", code=None)
