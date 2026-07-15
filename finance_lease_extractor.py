"""
finance_lease_extractor.py
==========================

Purpose
-------
Read company annual reports (PDF / TXT / Excel) and automatically pull out the
FINANCE LEASE liability amounts, split into:

    * Finance lease - SHORT TERM  (current / due within one year)
    * Finance lease - LONG TERM   (non-current / due after one year)

It searches the text for finance-lease keywords, grabs the number sitting on the
same line, decides whether the line is "current" or "non-current", and writes a
tidy results table. Anything ambiguous is FLAGGED for a human to check.

Why the flags matter (read this!)
---------------------------------
1. IFRS 16 companies often show ONLY "Lease liabilities (current / non-current)"
   without the word "finance". That number may mix operating + finance leases.
   -> we still capture it, but we FLAG it as "IFRS-bare-lease (may include
      operating)".
2. "Capital lease" is the OLD US-GAAP name for a finance lease -> treated the same.
3. "Operating lease" is NOT a finance lease -> such lines are EXCLUDED.
4. Finance leases are frequently NOT a face-of-balance-sheet line. They sit
   *inside* "Other current liabilities" / "Other non-current liabilities" and are
   only broken out in the NOTES. -> we scan the whole document, and if a value is
   found only near an "other liabilities" note we FLAG it.
5. Combined lines ("...long-term debt and finance leases") can't be split
   cleanly -> captured and FLAGGED.

Usage
-----
    python finance_lease_extractor.py <folder-with-reports>  [-o results.csv]

    # or a single file
    python finance_lease_extractor.py report.pdf

Supported inputs: .pdf  .txt  .text  .xlsx  .xls
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from dataclasses import dataclass, field
from typing import List, Optional


# ---------------------------------------------------------------------------
# 1. KEYWORDS  --  edit these lists to teach the script new wording
# ---------------------------------------------------------------------------

# A line must contain one of these to be a finance-lease candidate.
FINANCE_LEASE_TERMS = [
    "finance lease",
    "finance leases",
    "financial lease",          # non-native wording seen in some reports
    "financial leases",
    "capital lease",            # old US-GAAP term = finance lease
    "capital leases",
    "obligations under finance lease",
    "finance lease obligation",
    "finance lease liabilit",   # matches liability / liabilities
]

# Bare lease wording (IFRS 16) -- captured but flagged, because it may also
# contain operating leases.
BARE_LEASE_TERMS = [
    "lease liabilit",           # "lease liability" / "lease liabilities"
    "lease obligation",
]

# If a line mentions any of these it is NOT a finance lease -> skip it.
EXCLUDE_TERMS = [
    "operating lease",
    "right-of-use",
    "right of use",
    "rou asset",
]

# --- Bucket markers, in tiers of confidence ---------------------------------
# TIER 1 (definitive): if present anywhere on the line, they decide the bucket
# outright -- e.g. "Current portion of long-term debt" is CURRENT even though
# "long-term" also appears in it.
DEFINITIVE_CURRENT = [
    "current portion", "current maturit", "current installment",
    "current instalment", "due within one year", "within one year",
    "within 1 year", "less than one year", "repayable within one year",
]
DEFINITIVE_NONCURRENT = [
    "non-current", "noncurrent", "non current", "net of current",
    "after one year", "later than one year", "more than one year",
    "beyond one year", "due after one year", "repayable after one year",
]

# TIER 2 (weak): used only when no definitive phrase is on the line. Compared by
# position, with non-current wording masked so "current" hiding inside
# "non-current" can't win.
WEAK_CURRENT = [
    "short-term", "short term", ", current", "- current", "– current",
    "current:", "current liabilit",
]
WEAK_NONCURRENT = ["long-term", "long term"]

# Words hinting the value is buried inside an "other liabilities" bucket.
OTHER_BUCKET_MARKERS = [
    "other current liabilit",
    "other non-current liabilit",
    "other noncurrent liabilit",
    "other long-term liabilit",
    "other financial liabilit",
    "included in other",
    "within other",
]

# ---------------------------------------------------------------------------
# 2. NUMBER PARSING
# ---------------------------------------------------------------------------

# Matches figures like  1,234   1,234.5   (1,234)   1 234   12.3
NUMBER_RE = re.compile(
    r"""
    \(?                       # optional opening paren (negative)
    (?:USD|US\$|\$|£|€|Rs\.?|INR|EUR|GBP)?\s?   # optional currency symbol
    -?                        # optional minus
    \d{1,3}                   # first group of digits
    (?:[,\s]\d{3})*           # thousands separated by comma or space
    (?:\.\d+)?                # optional decimals
    \)?                       # optional closing paren
    """,
    re.VERBOSE,
)


# If one of these words sits just before a number, the number is a reference
# (accounting standard / note number), NOT a money figure -> skip it.
SKIP_PRECEDING = ("ifrs", "ias ", "asc ", "gaap", "note ", "no. ", "section ", "topic ")


def parse_number(token: str) -> Optional[float]:
    """Turn a messy figure string into a float. Parentheses => negative."""
    t = token.strip()
    if not t:
        return None
    negative = t.startswith("(") and t.endswith(")")
    # strip currency symbols, parens, spaces used as separators
    t = re.sub(r"[()$£€]|USD|US\$|Rs\.?|INR|EUR|GBP", "", t)
    t = t.replace(",", "").replace(" ", "").strip()
    if t in ("", "-", "."):
        return None
    try:
        val = float(t)
    except ValueError:
        return None
    return -val if negative else val


def first_number_on_line(line: str) -> Optional[float]:
    """
    Return the LEFT-MOST money figure on a line.

    Financial statements print the CURRENT reporting year in the first money
    column and the prior year to its right, so the first figure is the one we
    want. We skip:
      * bare 4-digit years (2023 / 2024),
      * numbers that are really references ("IFRS 16", "ASC 842", "Note 18"),
      * numbers that count time ("1 year", "12 months").
    """
    low = line.lower()
    for m in NUMBER_RE.finditer(line):
        raw = m.group(0).strip()
        val = parse_number(raw)
        if val is None:
            continue
        start, end = m.start(), m.end()
        # skip bare 4-digit years like 2023 / 2024
        if raw.isdigit() and 1900 <= val <= 2099 and len(raw) == 4:
            continue
        # skip accounting-standard / note references sitting before the number
        preceding = low[max(0, start - 8):start]
        if any(k in preceding for k in SKIP_PRECEDING):
            continue
        # skip durations like "1 year" / "12 months"
        following = low[end:end + 8].lstrip()
        if following.startswith(("year", "month")):
            continue
        return val
    return None


# ---------------------------------------------------------------------------
# 3. FILE READERS
# ---------------------------------------------------------------------------

def read_txt(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def read_pdf(path: str) -> str:
    try:
        import pdfplumber
    except ImportError:
        raise RuntimeError("pdfplumber not installed. Run: pip install pdfplumber")
    parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def read_excel(path: str) -> str:
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("openpyxl not installed. Run: pip install openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cells = ["" if c is None else str(c) for c in row]
            if any(cells):
                lines.append("\t".join(cells))
    return "\n".join(lines)


def read_any(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return read_pdf(path)
    if ext in (".txt", ".text", ".md"):
        return read_txt(path)
    if ext in (".xlsx", ".xls"):
        return read_excel(path)
    raise ValueError(f"Unsupported file type: {ext}")


# ---------------------------------------------------------------------------
# 4. CORE EXTRACTION
# ---------------------------------------------------------------------------

@dataclass
class Hit:
    """One finance-lease line we matched."""
    line: str
    value: Optional[float]
    bucket: str              # "current" | "non-current" | "unknown"
    is_finance: bool         # True = explicit finance/capital lease
    flags: List[str] = field(default_factory=list)


def _contains_any(text: str, terms) -> bool:
    return any(t in text for t in terms)


def _mask_noncurrent(text: str) -> str:
    """
    Blank out non-current wording so the CURRENT search can't be fooled by the
    substring 'current' hiding inside 'non-current liabilities'.
    """
    for token in ("non-current", "noncurrent", "non current"):
        text = text.replace(token, "§" * len(token))
    return text


# Section headings on a balance sheet. These are strong, unambiguous signals,
# unlike stray marker words that can appear inside an unrelated detail line
# (e.g. "current maturities" sitting inside a long-term debt line).
SECTION_NONCURRENT_HEADINGS = [
    "non-current liabilit", "noncurrent liabilit", "non current liabilit",
    "long-term liabilit", "long term liabilit",
]
SECTION_CURRENT_HEADINGS = ["current liabilit"]


def _heading_bucket(line_low: str) -> Optional[str]:
    """If this line is a liabilities section heading, return its bucket."""
    if any(h in line_low for h in SECTION_NONCURRENT_HEADINGS):
        return "non-current"
    masked = _mask_noncurrent(line_low)          # avoid 'current' in 'non-current'
    if any(h in masked for h in SECTION_CURRENT_HEADINGS):
        return "current"
    return None


def classify_bucket(line: str, context_before: str) -> str:
    """Decide current vs non-current: definitive phrase, then weak marker, then
    the nearest section heading above the line."""
    low = line.lower()

    # Tier 1 -- definitive phrases on the line win outright.
    if any(m in low for m in DEFINITIVE_CURRENT):
        return "current"
    if any(m in low for m in DEFINITIVE_NONCURRENT):
        return "non-current"

    # Tier 2 -- weaker markers, compared by position (non-current masked).
    masked = _mask_noncurrent(low)
    cur_pos = max((masked.rfind(m) for m in WEAK_CURRENT), default=-1)
    non_pos = max((low.rfind(m) for m in WEAK_NONCURRENT), default=-1)
    if cur_pos != -1 or non_pos != -1:
        return "current" if cur_pos > non_pos else "non-current"

    # Tier 3 -- nearest section heading above the line.
    for prev in reversed(context_before.splitlines()):
        bucket = _heading_bucket(prev.lower())
        if bucket:
            return bucket
    return "unknown"


def scan_text(text: str) -> List[Hit]:
    """Walk every line and collect finance-lease hits."""
    lines = text.splitlines()
    hits: List[Hit] = []
    for i, raw in enumerate(lines):
        low = raw.lower()

        is_finance = _contains_any(low, FINANCE_LEASE_TERMS)
        is_bare = _contains_any(low, BARE_LEASE_TERMS)
        if not (is_finance or is_bare):
            continue

        # never count operating leases / ROU asset lines
        if _contains_any(low, EXCLUDE_TERMS):
            continue

        context_before = "\n".join(lines[max(0, i - 10):i])
        bucket = classify_bucket(raw, context_before)
        value = first_number_on_line(raw)

        # if the matched line has no number, peek at the next line
        # (reports sometimes wrap the label and value onto two lines)
        if value is None and i + 1 < len(lines):
            value = first_number_on_line(lines[i + 1])

        flags: List[str] = []
        if is_bare and not is_finance:
            flags.append("IFRS-bare-lease (may include operating lease)")
        if _contains_any(low, OTHER_BUCKET_MARKERS) or _contains_any(
            context_before.lower(), OTHER_BUCKET_MARKERS
        ):
            flags.append("buried in 'other liabilities' - verify in notes")
        if bucket == "unknown":
            flags.append("could not tell current vs non-current")
        if value is None:
            flags.append("no number found on/after line")
        if ("debt" in low or "borrowing" in low) and ("lease" in low):
            flags.append("combined debt+lease line - cannot split cleanly")

        hits.append(Hit(raw.strip(), value, bucket, is_finance, flags))
    return hits


@dataclass
class CompanyResult:
    name: str
    short_term: Optional[float] = None
    long_term: Optional[float] = None
    flags: List[str] = field(default_factory=list)
    hits: List[Hit] = field(default_factory=list)


def summarise(name: str, hits: List[Hit]) -> CompanyResult:
    """
    Reduce many hits to one short-term and one long-term number.

    Preference order for each bucket:
      1. explicit finance/capital lease hit with a value
      2. otherwise a bare 'lease liabilities' hit with a value (flagged)
    We take the FIRST such hit per bucket (balance-sheet face value usually
    comes before note detail) and record every flag we saw.
    """
    res = CompanyResult(name=name, hits=hits)

    def pick(bucket: str):
        finance = [h for h in hits if h.bucket == bucket and h.is_finance and h.value is not None]
        if finance:
            return finance[0]
        bare = [h for h in hits if h.bucket == bucket and not h.is_finance and h.value is not None]
        if bare:
            return bare[0]
        return None

    st = pick("current")
    lt = pick("non-current")
    if st:
        res.short_term = st.value
        res.flags.extend(st.flags)
    if lt:
        res.long_term = lt.value
        res.flags.extend(lt.flags)

    # surface unknown-bucket finance hits so nothing is silently dropped
    for h in hits:
        if h.bucket == "unknown" and (h.is_finance or h.value is not None):
            res.flags.append(f"unclassified finance-lease line: '{h.line[:60]}'")

    if not hits:
        res.flags.append("NO finance/lease line found - check manually")

    # de-duplicate flags but keep order
    seen = set()
    res.flags = [f for f in res.flags if not (f in seen or seen.add(f))]
    return res


# ---------------------------------------------------------------------------
# 5. DRIVER
# ---------------------------------------------------------------------------

def process_path(path: str) -> CompanyResult:
    name = os.path.splitext(os.path.basename(path))[0]
    try:
        text = read_any(path)
    except Exception as exc:  # noqa: BLE001
        r = CompanyResult(name=name)
        r.flags.append(f"could not read file: {exc}")
        return r
    hits = scan_text(text)
    return summarise(name, hits)


def gather_files(target: str) -> List[str]:
    if os.path.isfile(target):
        return [target]
    exts = (".pdf", ".txt", ".text", ".md", ".xlsx", ".xls")
    out = []
    for root, _dirs, files in os.walk(target):
        for f in sorted(files):
            if f.lower().endswith(exts):
                out.append(os.path.join(root, f))
    return out


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Extract finance-lease short/long-term values from annual reports.")
    ap.add_argument("target", help="A report file OR a folder of reports")
    ap.add_argument("-o", "--output", default="finance_lease_results.csv", help="CSV output path")
    ap.add_argument("-v", "--verbose", action="store_true", help="Print every matched line")
    args = ap.parse_args(argv)

    files = gather_files(args.target)
    if not files:
        print(f"No supported reports found in: {args.target}", file=sys.stderr)
        return 1

    results = [process_path(p) for p in files]

    # write CSV
    with open(args.output, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Company", "Finance Lease - Short Term", "Finance Lease - Long Term", "Flags"])
        for r in results:
            w.writerow([
                r.name,
                "" if r.short_term is None else r.short_term,
                "" if r.long_term is None else r.long_term,
                " | ".join(r.flags),
            ])

    # print to screen
    print(f"\n{'Company':<28} {'Short-term':>14} {'Long-term':>14}   Flags")
    print("-" * 100)
    for r in results:
        st = "-" if r.short_term is None else f"{r.short_term:,.1f}"
        lt = "-" if r.long_term is None else f"{r.long_term:,.1f}"
        print(f"{r.name:<28} {st:>14} {lt:>14}   {'; '.join(r.flags) if r.flags else 'ok'}")
        if args.verbose:
            for h in r.hits:
                print(f"      [{h.bucket:<11}] val={h.value} :: {h.line[:80]}")
    print(f"\nSaved -> {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
