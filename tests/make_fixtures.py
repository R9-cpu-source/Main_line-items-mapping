#!/usr/bin/env python3
"""Generate 10-K-style fixtures (PDF + XLSX) with KNOWN finance-lease answers,
so finance_lease_extractor.py can be verified end-to-end.

Each fixture targets one extraction rule. Expected answers are in EXPECTED.
Run:  python tests/make_fixtures.py   ->  writes PDFs/XLSX into tests/fixtures/
"""
from __future__ import annotations

from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

FIX = Path(__file__).resolve().parent / "fixtures"

# company file -> (expected short-term, expected long-term)  (None = not disclosed)
EXPECTED = {
    "R1_Explicit_10K.pdf": (34.5, 128.7),
    "R5_FinanceColumn_10K.pdf": (34.5, 128.7),
    "R6_DebtTable_10K.pdf": (500.0, 1200.0),
    "R7_TotalDerived_10K.pdf": (40.0, 260.0),   # LT derived = 300 total - 40 current
    "Immaterial_10K.pdf": (None, None),
    "Combined_10K.xlsx": (None, None),
}


def _pdf(name: str, lines: list[str]) -> None:
    """Render lines to a PDF. A line equal to '<<PAGEBREAK>>' starts a new page."""
    FIX.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(FIX / name), pagesize=letter)
    c.setFont("Courier", 9)
    y = 740
    for ln in lines:
        if ln == "<<PAGEBREAK>>":
            c.showPage(); c.setFont("Courier", 9); y = 740
            continue
        c.drawString(48, y, ln)
        y -= 13
        if y < 60:
            c.showPage(); c.setFont("Courier", 9); y = 740
    c.save()


def r1_explicit() -> None:
    _pdf("R1_Explicit_10K.pdf", [
        "SAMPLE CORP - FORM 10-K",
        "Note 10. Leases  ($ in millions)                    2024        2023",
        "",
        "Finance lease liabilities, current                  34.5        30.1",
        "Operating lease liabilities, current                60.2        58.0",
        "Non-current finance lease liabilities              128.7       140.9",
        "Operating lease liabilities, non-current           410.0       420.0",
    ])


def r5_finance_column() -> None:
    _pdf("R5_FinanceColumn_10K.pdf", [
        "COLUMNAR CORP - FORM 10-K",
        "Note 12. Leases  ($ in millions)",
        "                                   Operating Leases   Finance Leases",
        "Lease liabilities, current                  120.0             34.5",
        "Lease liabilities, non-current              410.0            128.7",
    ])


def r6_debt_table() -> None:
    _pdf("R6_DebtTable_10K.pdf", [
        "BIGBLUE CORP - FORM 10-K",
        "The amounts of finance leases recognized in the consolidated balance sheet",
        "were as follows ($ in millions):                    2024        2023",
        "",
        "Short-term debt                                    500.0       480.0",
        "Long-term debt                                   1,200.0     1,150.0",
    ])


def r7_total_derived() -> None:
    # Realistic layout: the current-debt table and the maturity total sit in
    # separate notes / pages, as they do in real filings.
    _pdf("R7_TotalDerived_10K.pdf", [
        "PLANEMAKER CORP - FORM 10-K",
        "Note 14. Debt  ($ in millions)",
        "Short-term debt and current portion of long-term debt",
        "Finance lease obligations                           40.0        38.0",
        "<<PAGEBREAK>>",
        "Note 15. Lease Commitments  ($ in millions)",
        "Future minimum finance lease payments:",
        "Finance lease obligations due through 2044         300.0       290.0",
    ])


def immaterial() -> None:
    _pdf("Immaterial_10K.pdf", [
        "SMALLCO CORP - FORM 10-K",
        "Note 9. Leases",
        "The Company's finance leases are not material to the consolidated",
        "financial statements and are therefore not presented separately.",
    ])


def combined_xlsx() -> None:
    FIX.mkdir(parents=True, exist_ok=True)
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["COMBINED CORP - Balance Sheet", "2024", "2023"])
    ws.append(["Current portion of long-term debt and finance leases", 145, 132])
    ws.append(["Long-term debt and finance leases", 4560, 4300])
    wb.save(str(FIX / "Combined_10K.xlsx"))


def main() -> None:
    r1_explicit()
    r5_finance_column()
    r6_debt_table()
    r7_total_derived()
    immaterial()
    combined_xlsx()
    print(f"Wrote {len(EXPECTED)} fixtures to {FIX}")


if __name__ == "__main__":
    main()
