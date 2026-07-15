#!/usr/bin/env python3
"""Extract current and non-current finance lease liabilities from annual reports.

The extractor is intentionally conservative:
- It searches lease-note / debt-note wording, not only the balance-sheet face.
- It excludes operating leases, ROU assets, lease cost, cash-flow items and
  undiscounted maturity payments.
- It returns REVIEW / NOT DISCLOSED rather than guessing.

Primary text engine: Poppler pdftotext -layout (when installed).
Fallback text engine: PyMuPDF.

Outputs: CSV, JSON and an evidence text file. Excel can be built from the CSV
using the companion build_finance_lease_excel.py script.
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, Optional
import xml.etree.ElementTree as ET

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
MONEY_RE = re.compile(r"(?<![A-Za-z0-9])(?:\(?\s*[-+]?\s*\$?\s*\d[\d,]*(?:\.\d+)?\s*\)?)")

CURRENT_LABELS = (
    "other current liabilities",
    "accrued expenses and other current liabilities",
    "accrued expenses and other liabilities",
    "accounts payable and other accrued liabilities",
    "short-term debt",
)
LONG_LABELS = (
    "other long-term liabilities",
    "other non-current liabilities",
    "other noncurrent liabilities",
    "long-term debt",
)

EXCLUDE_LINE_TERMS = (
    "right-of-use asset",
    "right of use asset",
    "rou asset",
    "lease cost",
    "lease expense",
    "principal payment",
    "cash flow",
    "cash paid",
    "assets obtained",
    "acquired under finance leases",
    "including interest",
    "total finance lease payments",
    "finance lease revenue",
    "finance receivable",
    "sales-type",
    "lessor",
)

@dataclass
class Candidate:
    kind: str  # current | long_term | total
    amount_mm: float
    page: int
    rule: str
    score: int
    line: str
    evidence: str
    parent_line_item: str
    unit_basis: str
    derived: bool = False
    order: int = 0

@dataclass
class Result:
    company: str
    file_name: str
    finance_lease_short_term_usd_mm: Optional[float]
    finance_lease_long_term_usd_mm: Optional[float]
    short_term_page: Optional[int]
    long_term_page: Optional[int]
    short_term_parent_line_item: str
    long_term_parent_line_item: str
    short_term_exact_source_line: str
    long_term_exact_source_line: str
    short_term_rule: str
    long_term_rule: str
    short_term_confidence: str
    long_term_confidence: str
    overall_status: str
    review_flag: str
    verification_1: str
    verification_2: str


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def company_from_filename(path: Path) -> str:
    return re.sub(r"_10K$", "", path.stem, flags=re.I)


def extract_pdf_pages_pdftotext(path: Path) -> list[str]:
    exe = shutil.which("pdftotext")
    if not exe:
        raise RuntimeError("pdftotext is not installed")
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
        out = Path(tmp.name)
    try:
        subprocess.run(
            [exe, "-layout", str(path), str(out)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        text = out.read_text(encoding="utf-8", errors="replace")
        return text.split("\f")
    finally:
        out.unlink(missing_ok=True)


def extract_pdf_pages_pymupdf(path: Path) -> list[str]:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed")
    pages: list[str] = []
    with fitz.open(path) as doc:
        for page in doc:
            # sort=True is more stable for tables than plain extraction.
            pages.append(page.get_text("text", sort=True))
    return pages


def extract_pdf_pages(path: Path, engine: str) -> list[str]:
    if engine == "pdftotext":
        return extract_pdf_pages_pdftotext(path)
    if engine == "pymupdf":
        return extract_pdf_pages_pymupdf(path)
    if engine == "auto":
        try:
            return extract_pdf_pages_pdftotext(path)
        except Exception:
            return extract_pdf_pages_pymupdf(path)
    raise ValueError(f"Unknown engine: {engine}")


def parse_xlsx_text(path: Path) -> list[str]:
    """Read visible cell text from XLSX without modifying the workbook."""
    lines: list[str] = []
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", ns):
                shared.append("".join(t.text or "" for t in si.iterfind(".//a:t", ns)))
        for name in sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n)):
            root = ET.fromstring(z.read(name))
            for row in root.findall(".//a:row", ns):
                vals: list[str] = []
                for c in row.findall("a:c", ns):
                    t = c.attrib.get("t")
                    v = c.find("a:v", ns)
                    if v is None:
                        inline = c.find("a:is/a:t", ns)
                        vals.append(inline.text if inline is not None and inline.text else "")
                    elif t == "s":
                        idx = int(v.text or 0)
                        vals.append(shared[idx] if idx < len(shared) else "")
                    else:
                        vals.append(v.text or "")
                if any(v.strip() for v in vals):
                    lines.append(" | ".join(vals))
    return lines


def clean_line_for_numbers(line: str) -> str:
    # Remove footnote markers attached to labels, e.g. liabilities(2).
    s = re.sub(r"(?<=[A-Za-z])\(\d{1,2}\)", "", line)
    # Remove maturity years embedded in labels, e.g. due through 2044.
    s = re.sub(r"\b(?:due|through|until|maturing|matures?)\s+(?:in\s+)?(?:19|20)\d{2}\b", "", s, flags=re.I)
    return s


def parse_numbers(line: str) -> list[float]:
    s = clean_line_for_numbers(line)
    values: list[float] = []
    for m in MONEY_RE.finditer(s):
        raw = m.group(0).strip()
        compact = raw.replace("$", "").replace(",", "").replace(" ", "")
        negative = compact.startswith("(") and compact.endswith(")")
        compact = compact.strip("()")
        if compact.startswith("+"):
            compact = compact[1:]
        try:
            val = float(compact)
        except ValueError:
            continue
        values.append(-val if negative else val)
    return values


def find_year_header(lines: list[str], idx: int, lookback: int = 45) -> list[int]:
    best: list[int] = []
    for j in range(idx - 1, max(-1, idx - lookback) - 1, -1):
        years = [int(y) for y in YEAR_RE.findall(lines[j])]
        if len(years) >= 2:
            return years
        if len(years) == 1 and not best:
            best = years
    return best


def has_finance_columns(lines: list[str], idx: int, lookback: int = 18) -> bool:
    for j in range(idx - 1, max(-1, idx - lookback) - 1, -1):
        low = lines[j].lower()
        if "operating" in low and "finance" in low and "lease" in low:
            return True
        # Header can be split across adjacent lines, for example:
        #   Operating          Finance
        #   Leases             Leases
        if "operating" in low and "finance" in low and j + 1 < len(lines) and "lease" in lines[j + 1].lower():
            return True
        if "operating" in low and j + 1 < len(lines) and "finance" in lines[j + 1].lower():
            return True
    return False


def nearest_lease_heading(lines: list[str], idx: int, lookback: int = 18) -> str:
    for j in range(idx - 1, max(-1, idx - lookback) - 1, -1):
        low = normalize_space(lines[j]).lower()
        if low in {"finance leases", "finance leases:"} or (low.startswith("finance leases") and len(low) < 35):
            return "finance"
        if low in {"operating leases", "operating leases:"} or (low.startswith("operating leases") and len(low) < 35):
            return "operating"
    return ""


def select_latest_year_value(line: str, lines: list[str], idx: int) -> Optional[float]:
    vals = parse_numbers(line)
    if not vals:
        return None
    years = find_year_header(lines, idx)
    if len(years) >= 2 and len(vals) >= len(years):
        vals = vals[-len(years):]
        return vals[years.index(max(years))]
    return vals[0]


def select_finance_column_value(line: str) -> Optional[float]:
    vals = parse_numbers(line)
    if len(vals) >= 2:
        # Standard tables use Operating, Finance, [Total].
        return vals[1]
    return vals[0] if vals else None


def evidence_window(lines: list[str], idx: int, radius: int = 3) -> str:
    start = max(0, idx - radius)
    end = min(len(lines), idx + radius + 1)
    return "\n".join(normalize_space(x) for x in lines[start:end] if normalize_space(x))


def detect_unit(page_text: str, evidence: str) -> tuple[float, str]:
    search = (evidence + "\n" + page_text).lower()
    if re.search(r"(?:\$\s*)?in thousands|amounts? in thousands", search):
        return 0.001, "USD thousands converted to USD millions"
    if re.search(r"(?:\$\s*)?in billions|amounts? in billions", search):
        return 1000.0, "USD billions converted to USD millions"
    return 1.0, "USD millions"


def good_balance_sheet_line(line: str) -> bool:
    low = line.lower()
    return not any(term in low for term in EXCLUDE_LINE_TERMS)


def add_candidate(
    out: list[Candidate], *, kind: str, raw_amount: Optional[float], page_num: int,
    rule: str, score: int, line: str, evidence: str, parent: str,
    page_text: str, derived: bool = False,
) -> None:
    if raw_amount is None or math.isnan(raw_amount):
        return
    factor, unit = detect_unit(page_text, evidence)
    amount = abs(raw_amount) * factor
    # Reject implausible negative/zero table artefacts; zero is allowed but not useful here.
    if amount < 0:
        return
    out.append(Candidate(kind, amount, page_num, rule, score, normalize_space(line), evidence, parent, unit, derived, len(out)))


def analyze_pages(pages: list[str]) -> tuple[list[Candidate], str]:
    candidates: list[Candidate] = []
    full_lower = "\n".join(pages).lower()
    disclosure_status = ""
    if re.search(r"finance leases? (?:were|are|is) (?:not material|immaterial|not significant)", full_lower):
        disclosure_status = "Finance leases are stated to be immaterial/not significant, but no segregated amount is provided."

    for pnum, page in enumerate(pages, start=1):
        lines = page.splitlines()
        for i, line in enumerate(lines):
            low = normalize_space(line).lower()
            if not low:
                continue

            # R1: Explicit current / non-current labels on the same line.
            if good_balance_sheet_line(line):
                if re.search(r"finance lease (?:liabilities|obligations).*(?:due within one year|current)", low):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R1 explicit finance-lease current label", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent=normalize_space(line), page_text=page)
                if re.search(r"(?:long-term|non-current|noncurrent) finance lease (?:liabilities|obligations)", low) or re.search(r"finance lease (?:liabilities|obligations).*(?:long-term|non-current|noncurrent)", low):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                  rule="R1 explicit finance-lease non-current label", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent=normalize_space(line), page_text=page)

            # R2: Apple-style row where finance lease and current parent line are together,
            # followed by a non-current parent line that does not repeat 'finance leases'.
            if "finance leases" in low and any(lbl in low for lbl in CURRENT_LABELS) and good_balance_sheet_line(line):
                val = select_latest_year_value(line, lines, i)
                add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                              rule="R2 finance-lease row mapped to current parent line", score=98,
                              line=line, evidence=evidence_window(lines, i), parent=next((x for x in CURRENT_LABELS if x in low), "current liabilities"), page_text=page)
                for k in range(i + 1, min(len(lines), i + 4)):
                    nxt = normalize_space(lines[k]).lower()
                    if any(lbl in nxt for lbl in LONG_LABELS) and good_balance_sheet_line(lines[k]):
                        val2 = select_latest_year_value(lines[k], lines, k)
                        add_candidate(candidates, kind="long_term", raw_amount=val2, page_num=pnum,
                                      rule="R2 continuation row mapped to non-current parent line", score=98,
                                      line=lines[k], evidence=evidence_window(lines, k), parent=next((x for x in LONG_LABELS if x in nxt), "non-current liabilities"), page_text=page)
                        break

            # R3: Rows inside an explicit Finance Leases balance-sheet section.
            heading = nearest_lease_heading(lines, i)
            if heading == "finance" and good_balance_sheet_line(line):
                if any(lbl in low for lbl in CURRENT_LABELS):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R3 current parent line inside Finance Leases section", score=98,
                                  line=line, evidence=evidence_window(lines, i), parent=next((x for x in CURRENT_LABELS if x in low), "current liabilities"), page_text=page)
                if any(lbl in low for lbl in LONG_LABELS):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                  rule="R3 non-current parent line inside Finance Leases section", score=98,
                                  line=line, evidence=evidence_window(lines, i), parent=next((x for x in LONG_LABELS if x in low), "non-current liabilities"), page_text=page)

            # R4: Section heading 'Short-term lease liabilities' / 'Long-term lease liabilities',
            # followed by a 'Finance leases' row.
            if "short-term lease liabilities" in low or "current lease liabilities" in low:
                for k in range(i + 1, min(len(lines), i + 6)):
                    nxt = normalize_space(lines[k]).lower()
                    if nxt.startswith("finance leases") and good_balance_sheet_line(lines[k]):
                        val = select_latest_year_value(lines[k], lines, k)
                        add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                      rule="R4 finance row under short-term lease-liability heading", score=99,
                                      line=lines[k], evidence=evidence_window(lines, k), parent=normalize_space(line), page_text=page)
                        break
            if "long-term lease liabilities" in low or "non-current lease liabilities" in low or "noncurrent lease liabilities" in low:
                for k in range(i + 1, min(len(lines), i + 6)):
                    nxt = normalize_space(lines[k]).lower()
                    if nxt.startswith("finance leases") and good_balance_sheet_line(lines[k]):
                        val = select_latest_year_value(lines[k], lines, k)
                        add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                      rule="R4 finance row under long-term lease-liability heading", score=99,
                                      line=lines[k], evidence=evidence_window(lines, k), parent=normalize_space(line), page_text=page)
                        break

            # R5: A table with Operating Leases / Finance Leases columns.
            if has_finance_columns(lines, i):
                if re.search(r"lease liabilities,?\s*current", low):
                    val = select_finance_column_value(line)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R5 Finance column - current lease liabilities", score=99,
                                  line=line, evidence=evidence_window(lines, i), parent="Current lease liabilities", page_text=page)
                if re.search(r"lease liabilities,?\s*(?:non-current|noncurrent)", low):
                    val = select_finance_column_value(line)
                    add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                  rule="R5 Finance column - non-current lease liabilities", score=99,
                                  line=line, evidence=evidence_window(lines, i), parent="Non-current lease liabilities", page_text=page)
                if "less: current portion of lease liabilities" in low:
                    val = select_finance_column_value(line)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R5 Finance column - current portion of lease liabilities", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent="Current portion of lease liabilities", page_text=page)
                if "total long-term lease liabilities" in low:
                    val = select_finance_column_value(line)
                    add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                  rule="R5 Finance column - total long-term lease liabilities", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent="Long-term lease liabilities", page_text=page)

            # R6: IBM-style table introduced as finance leases recognized in the balance sheet.
            context = " ".join(normalize_space(x).lower() for x in lines[max(0, i - 12):i])
            if "finance leases recognized in the consolidated balance sheet" in context and good_balance_sheet_line(line):
                if low.startswith("short-term debt"):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R6 finance-lease table - short-term debt", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent="Short-term debt", page_text=page)
                if low.startswith("long-term debt"):
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="long_term", raw_amount=val, page_num=pnum,
                                  rule="R6 finance-lease table - long-term debt", score=100,
                                  line=line, evidence=evidence_window(lines, i), parent="Long-term debt", page_text=page)

            # R7: Boeing-style current row and total finance lease obligations.
            if normalize_space(low).startswith("finance lease obligations") and good_balance_sheet_line(line):
                # Current if section heading explicitly says short-term/current debt.
                prior = " ".join(normalize_space(x).lower() for x in lines[max(0, i - 18):i])
                if "short-term debt and current portion of long-term debt" in prior:
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="current", raw_amount=val, page_num=pnum,
                                  rule="R7 finance lease obligations in current-debt table", score=98,
                                  line=line, evidence=evidence_window(lines, i), parent="Short-term debt and current portion of long-term debt", page_text=page)
                if "due through" in low or "total finance lease obligations" in low:
                    val = select_latest_year_value(line, lines, i)
                    add_candidate(candidates, kind="total", raw_amount=val, page_num=pnum,
                                  rule="R7 total finance lease obligations", score=95,
                                  line=line, evidence=evidence_window(lines, i), parent="Total finance lease obligations", page_text=page)

    # Generic derivation: if total and current are available but long-term is not,
    # derive non-current = total - current.
    current = best_candidate(candidates, "current")
    long_term = best_candidate(candidates, "long_term")
    total = best_candidate(candidates, "total")
    if current and total and not long_term and total.amount_mm >= current.amount_mm:
        candidates.append(Candidate(
            kind="long_term",
            amount_mm=round(total.amount_mm - current.amount_mm, 6),
            page=total.page,
            rule="R8 derived: total finance lease obligations less current portion",
            score=92,
            line=f"Derived from total {total.amount_mm:g} less current {current.amount_mm:g}",
            evidence=total.evidence + "\n--- CURRENT EVIDENCE ---\n" + current.evidence,
            parent_line_item="Non-current portion derived from total finance lease obligations",
            unit_basis="USD millions",
            derived=True,
            order=len(candidates),
        ))
    return candidates, disclosure_status


def best_candidate(candidates: list[Candidate], kind: str) -> Optional[Candidate]:
    relevant = [c for c in candidates if c.kind == kind]
    if not relevant:
        return None
    relevant.sort(key=lambda c: (c.score, c.page, c.order, not c.derived), reverse=True)
    return relevant[0]


def conflict_note(candidates: list[Candidate], kind: str) -> str:
    relevant = sorted([c for c in candidates if c.kind == kind], key=lambda c: c.score, reverse=True)
    if len(relevant) < 2:
        return ""
    top = relevant[0]
    close = [
        c for c in relevant[1:]
        if c.score >= top.score - 3
        and abs(c.amount_mm - top.amount_mm) > 0.001
        # Repeated prior-year and latest-year tables often appear on the same
        # page with the same rule/parent line (Amazon is a common example).
        # The later occurrence is selected; the earlier occurrence is not a
        # true classification conflict.
        and not (c.page == top.page and c.rule == top.rule and c.parent_line_item == top.parent_line_item)
    ]
    if close:
        vals = ", ".join(f"{c.amount_mm:g} (p.{c.page}, {c.rule})" for c in [top] + close[:3])
        return f"Conflicting high-score {kind} candidates: {vals}"
    return ""


def confidence(candidate: Optional[Candidate], conflict: str) -> str:
    if candidate is None:
        return "Not available"
    if conflict:
        return "Needs review"
    if candidate.score >= 98 and not candidate.derived:
        return "High"
    if candidate.score >= 92:
        return "Medium"
    return "Low"


def verify_evidence(candidate: Optional[Candidate]) -> str:
    if candidate is None:
        return "N/A - no value extracted"
    # Verification 1: the chosen line/evidence must contain finance lease terminology,
    # or be an explicitly documented derivation.
    evidence_low = candidate.evidence.lower()
    if candidate.derived:
        return "PASS - arithmetic derivation from total and current evidence"
    # R3/R5/R6 are only created after the parser has confirmed an explicit
    # Finance Leases heading/column in the nearby table, even when the chosen
    # row itself says only 'Other long-term liabilities' or 'Lease liabilities'.
    context_confirmed = candidate.rule.startswith(("R3", "R5", "R6"))
    if "finance lease" not in evidence_low and "finance leases" not in evidence_low and not context_confirmed:
        return "FAIL - evidence window lacks finance lease terminology"
    return "PASS - finance-lease table context and liability row confirmed"


def verify_arithmetic(current: Optional[Candidate], long_term: Optional[Candidate], total: Optional[Candidate]) -> str:
    # Verification 2: where a total is disclosed, current + long-term must reconcile.
    if current and long_term and total:
        diff = abs((current.amount_mm + long_term.amount_mm) - total.amount_mm)
        return "PASS - current + long-term reconciles to total" if diff <= 0.01 else f"FAIL - total reconciliation difference {diff:.3f} USD mm"
    if current and long_term:
        return "PASS - both segregated balance-sheet liabilities were extracted"
    return "REVIEW - one or both segregated amounts are unavailable"


def result_from_pdf(path: Path, engine: str) -> tuple[Result, list[Candidate]]:
    pages = extract_pdf_pages(path, engine)
    candidates, disclosure_status = analyze_pages(pages)
    cur = best_candidate(candidates, "current")
    lt = best_candidate(candidates, "long_term")
    total = best_candidate(candidates, "total")
    cur_conflict = conflict_note(candidates, "current")
    lt_conflict = conflict_note(candidates, "long_term")
    flags = "; ".join(x for x in [cur_conflict, lt_conflict] if x)
    if not cur and not lt:
        if disclosure_status:
            status = "Immaterial/unquantified"
            flags = (flags + "; " if flags else "") + disclosure_status
        else:
            status = "Not separately disclosed"
            flags = (flags + "; " if flags else "") + "No reliable segregated finance-lease liability amounts found; no zero was assumed."
    elif cur and lt:
        status = "Extracted - both amounts" if not (cur.derived or lt.derived) else "Extracted - one amount derived"
    else:
        status = "Partial - one amount unavailable"
        flags = (flags + "; " if flags else "") + "Only one of the two segregated amounts was found."

    v1_parts = [verify_evidence(cur), verify_evidence(lt)]
    v1 = " | ".join(v1_parts)
    v2 = verify_arithmetic(cur, lt, total)
    return Result(
        company=company_from_filename(path),
        file_name=path.name,
        finance_lease_short_term_usd_mm=cur.amount_mm if cur else None,
        finance_lease_long_term_usd_mm=lt.amount_mm if lt else None,
        short_term_page=cur.page if cur else None,
        long_term_page=lt.page if lt else None,
        short_term_parent_line_item=cur.parent_line_item if cur else "",
        long_term_parent_line_item=lt.parent_line_item if lt else "",
        short_term_exact_source_line=cur.line if cur else "",
        long_term_exact_source_line=lt.line if lt else "",
        short_term_rule=cur.rule if cur else "",
        long_term_rule=lt.rule if lt else "",
        short_term_confidence=confidence(cur, cur_conflict),
        long_term_confidence=confidence(lt, lt_conflict),
        overall_status=status,
        review_flag=flags,
        verification_1=v1,
        verification_2=v2,
    ), candidates


def result_from_xlsx(path: Path) -> tuple[Result, list[Candidate]]:
    lines = parse_xlsx_text(path)
    joined = "\n".join(lines)
    finance_lines = [normalize_space(x) for x in lines if "finance lease" in x.lower()]
    combined = any("debt and finance leases" in x.lower() for x in finance_lines)
    if combined:
        status = "Combined/not segregated"
        flag = "The workbook contains combined debt-and-finance-lease lines but no lease note that separates current and non-current finance lease liabilities."
    else:
        status = "Not separately disclosed"
        flag = "No segregated finance-lease liability lines were found in the workbook."
    evidence = " || ".join(finance_lines[:6])
    return Result(
        company=company_from_filename(path), file_name=path.name,
        finance_lease_short_term_usd_mm=None, finance_lease_long_term_usd_mm=None,
        short_term_page=None, long_term_page=None,
        short_term_parent_line_item="", long_term_parent_line_item="",
        short_term_exact_source_line=evidence, long_term_exact_source_line=evidence,
        short_term_rule="XLSX text scan", long_term_rule="XLSX text scan",
        short_term_confidence="Not available", long_term_confidence="Not available",
        overall_status=status, review_flag=flag,
        verification_1="PASS - combined wording found" if combined else "PASS - no segregated wording found",
        verification_2="REVIEW - exact split cannot be computed from the supplied workbook",
    ), []


def collect_files(input_path: Path, work_dir: Path) -> list[Path]:
    if input_path.is_dir():
        return sorted([p for p in input_path.iterdir() if p.suffix.lower() in {".pdf", ".xlsx"}])
    if input_path.suffix.lower() == ".zip":
        extract_dir = work_dir / "reports"
        extract_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(input_path) as z:
            z.extractall(extract_dir)
        return sorted([p for p in extract_dir.rglob("*") if p.suffix.lower() in {".pdf", ".xlsx"}])
    if input_path.suffix.lower() in {".pdf", ".xlsx"}:
        return [input_path]
    raise ValueError("Input must be a directory, ZIP, PDF or XLSX")


def write_outputs(results: list[Result], all_candidates: dict[str, list[Candidate]], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = [asdict(r) for r in results]
    csv_path = out_dir / "finance_lease_results.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)
    (out_dir / "finance_lease_results.json").write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")

    with (out_dir / "finance_lease_evidence.txt").open("w", encoding="utf-8") as f:
        for r in results:
            f.write(f"\n{'='*100}\n{r.company} | {r.file_name}\nSTATUS: {r.overall_status}\nFLAG: {r.review_flag}\n")
            f.write(f"SHORT-TERM: {r.finance_lease_short_term_usd_mm} | page {r.short_term_page} | {r.short_term_rule}\n{r.short_term_exact_source_line}\n")
            f.write(f"LONG-TERM: {r.finance_lease_long_term_usd_mm} | page {r.long_term_page} | {r.long_term_rule}\n{r.long_term_exact_source_line}\n")
            f.write("\nALL CANDIDATES:\n")
            for c in sorted(all_candidates.get(r.file_name, []), key=lambda x: (x.kind, -x.score)):
                f.write(f"- {c.kind}: {c.amount_mm:g} USD mm | p.{c.page} | score {c.score} | {c.rule} | {c.line}\n")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="Directory, ZIP, PDF or XLSX containing annual reports")
    ap.add_argument("--output-dir", default="finance_lease_output")
    ap.add_argument("--engine", choices=["auto", "pdftotext", "pymupdf"], default="auto")
    args = ap.parse_args()

    input_path = Path(args.input).resolve()
    out_dir = Path(args.output_dir).resolve()
    with tempfile.TemporaryDirectory(prefix="finance_lease_extract_") as td:
        files = collect_files(input_path, Path(td))
        if not files:
            raise SystemExit("No PDF/XLSX reports found")
        results: list[Result] = []
        candidates_by_file: dict[str, list[Candidate]] = {}
        for path in files:
            try:
                if path.suffix.lower() == ".pdf":
                    result, cands = result_from_pdf(path, args.engine)
                else:
                    result, cands = result_from_xlsx(path)
            except Exception as exc:
                result = Result(
                    company=company_from_filename(path), file_name=path.name,
                    finance_lease_short_term_usd_mm=None, finance_lease_long_term_usd_mm=None,
                    short_term_page=None, long_term_page=None,
                    short_term_parent_line_item="", long_term_parent_line_item="",
                    short_term_exact_source_line="", long_term_exact_source_line="",
                    short_term_rule="", long_term_rule="",
                    short_term_confidence="Error", long_term_confidence="Error",
                    overall_status="Extraction error", review_flag=str(exc),
                    verification_1="FAIL", verification_2="FAIL",
                )
                cands = []
            results.append(result)
            candidates_by_file[path.name] = cands
            print(f"{result.company:45s} ST={result.finance_lease_short_term_usd_mm!s:>10s} LT={result.finance_lease_long_term_usd_mm!s:>10s} | {result.overall_status}")
        write_outputs(results, candidates_by_file, out_dir)
        print(f"\nSaved outputs to: {out_dir}")

if __name__ == "__main__":
    main()
