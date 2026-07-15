# Balance-Sheet Line-Item Mapping — Finance Lease Extractor

Extracts the **Finance Lease** liability from company annual reports, split into
**Short Term (current)** and **Long Term (non-current)**, with evidence, a
confidence level, and two self-verification checks. Designed for real 10-K /
annual-report PDFs (and XLSX), not just clean text.

---

## 1. The 5-category mapping (context)

The wider task maps balance-sheet items into 5 buckets. This tool focuses on
the two lease buckets (#4 and #5):

| # | Category | Typical line items / keywords |
|---|----------|-------------------------------|
| 1 | Long-Term Debt | Long-term borrowings; bonds/notes payable (non-current); term loans; senior notes; debentures |
| 2 | Short-Term Debt | Short-term borrowings; bank overdraft; commercial paper; current borrowings |
| 3 | Current Portion of Long-Term Debt | Current portion / current maturities of long-term debt |
| 4 | **Finance Lease — Long Term** | Finance lease liabilities (non-current); obligations under finance leases (non-current) |
| 5 | **Finance Lease — Short Term** | Finance lease liabilities (current); current portion of finance lease |

**Where the finance-lease number actually lives** (the tool searches all of these):
the face of the balance sheet; *inside* “Other current / non-current
liabilities” with the split only in the notes; a debt-note table (“finance
leases recognized in the balance sheet”); or a maturity table (“due through
YYYY”), from which the split is derived.

---

## 2. How it works (simple words)

`finance_lease_extractor.py`:

1. **Reads** each report to text — Poppler `pdftotext -layout` if installed,
   else **PyMuPDF**; XLSX via the standard library.
2. **Scans** every line with 8 rules (R1–R8) covering the common report layouts:
   - **R1** explicit “Finance lease liabilities, current / non-current”.
   - **R2** finance-lease row mapped to a current parent line, with the next
     non-current parent line (Apple-style).
   - **R3** rows inside an explicit “Finance Leases” section.
   - **R4** “Finance leases” row under a short-/long-term lease-liability heading.
   - **R5** table with **Operating / Finance** columns (picks the Finance column).
   - **R6** debt table “finance leases recognized in the balance sheet”
     (short-term debt / long-term debt).
   - **R7** current row + total finance-lease obligations (maturity table).
   - **R8** derive non-current = **total − current** when only those are given.
3. **Excludes** operating-lease-only noise: ROU assets, lease cost/expense,
   cash-flow and undiscounted-payment lines (see `EXCLUDE_LINE_TERMS`).
4. **Picks numbers** using the year header (latest reporting year) or the
   Finance column, converting thousands/billions to USD millions.
5. **Scores & ranks** candidates, notes **conflicts**, assigns a **confidence**
   (High / Medium / Needs review / Not available), and returns
   **NOT DISCLOSED / REVIEW rather than guessing**.
6. **Two verifications** per company: (V1) the chosen evidence really is
   finance-lease context; (V2) current + long-term reconciles to the disclosed
   total.

**Outputs** (in `--output-dir`, default `finance_lease_output/`):
`finance_lease_results.csv`, `finance_lease_results.json`, and
`finance_lease_evidence.txt` (every candidate + evidence window).

### Run it

```bash
pip install -r requirements.txt        # PyMuPDF fallback engine
# (optional, better tables) install Poppler: apt-get install poppler-utils

python finance_lease_extractor.py path/to/reports/      # folder
python finance_lease_extractor.py report.pdf            # single file
python finance_lease_extractor.py reports.zip           # a zip of reports
python finance_lease_extractor.py reports/ --engine pymupdf   # force engine
```

Inputs: `.pdf`, `.xlsx`, a folder, or a `.zip`.

---

## 3. Verification (reproducible)

`tests/make_fixtures.py` builds 10-K-style fixtures with **known** answers and
`tests/verify.py` asserts the extractor reproduces them:

```bash
pip install -r tests/requirements-dev.txt
python tests/verify.py        # regenerates fixtures, runs extractor, asserts
```

| Fixture (rule) | Expected ST | Expected LT | Result |
|---|---|---|---|
| R1 explicit finance-lease labels | 34.5 | 128.7 | ✅ |
| R5 Operating/Finance columns | 34.5 | 128.7 | ✅ |
| R6 debt table (finance leases in B/S) | 500 | 1,200 | ✅ |
| R7 total + current → R8 derived LT | 40 | 260 | ✅ |
| Immaterial disclosure | — | — | ✅ (Immaterial/unquantified) |
| Combined debt+lease XLSX | — | — | ✅ (Combined/not segregated) |

All pass, deterministically, on two consecutive runs.

---

## 4. Known limitations found during verification (review these on real data)

1. **R5 can fire on operating-lease rows.** `has_finance_columns()` treats an
   “operating” line followed by a “finance” line as a Finance column, and R5’s
   row regex also matches *“operating lease liabilities, non-current”*; plain
   `operating lease` is **not** in `EXCLUDE_LINE_TERMS`. This creates spurious
   candidates / false **“Needs review”** flags, and on a filing with no
   higher-scored rule it could return the operating figure. → treat R5-only,
   “Needs review” results as manual-check.
2. **“Capital lease” (pre-2019 term) is not matched** — only “finance lease”.
   Old filings that still say *capital lease* will read as *Not disclosed*.
3. **R7 proximity assumption.** If a maturity **total** line sits within ~18
   lines of the “short-term debt and current portion…” heading, it is also
   counted as current. Real filings separate them; crammed layouts can miscount.
4. **Unit detection is page-wide** — an unrelated “in thousands” note on the
   same page can rescale a value.

Items 1–2 are the highest-impact for a 20-company IFRS/US-GAAP mix. Say the word
and I’ll add an `operating lease` exclusion to R5 and a `capital lease` alias.
