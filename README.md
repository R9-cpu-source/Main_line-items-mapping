# Balance-Sheet Line-Item Mapping — Finance Lease Extractor

This project does two things:

1. **Explains the mapping** of balance-sheet (B&S) items into 5 debt/lease
   categories, and flags the confusing cases.
2. **Extracts, in code**, the **Finance Lease** amount from company annual
   reports, split into **Short Term** and **Long Term**, by searching for
   keywords and reading the number next to them.

---

## 1. The 5 categories and which line items fall under each

| # | Category | Typical line items / keywords in reports |
|---|----------|------------------------------------------|
| 1 | **Long-Term Debt** | Long-term borrowings; Bonds / Notes payable (non-current); Term loans (non-current); Senior notes; Debentures; Long-term loans & borrowings |
| 2 | **Short-Term Debt** | Short-term borrowings; Bank overdraft; Commercial paper; Revolving credit (current); Working capital loans; Current borrowings |
| 3 | **Current Portion of Long-Term Debt (CPLTD)** | Current portion / current maturities of long-term debt; Long-term debt due within one year; Current installments of long-term debt |
| 4 | **Finance Lease — Long Term** | Finance lease liabilities (non-current); Obligations under finance leases (non-current); **Capital lease** obligations (non-current — old US term) |
| 5 | **Finance Lease — Short Term** | Finance lease liabilities (current); Current portion of finance lease; Obligations under finance leases (current); Capital lease (current) |

### Where the finance-lease numbers actually live

They are **not always a separate line** on the face of the balance sheet. Look
in this order:

1. **Face of the balance sheet** — a clear line such as *"Finance lease
   liabilities, current / non-current"* (mostly US GAAP / ASC 842).
2. **Rolled into "Other liabilities"** — the balance sheet only shows *"Other
   current liabilities"* / *"Other non-current liabilities"*, and the finance
   lease amount is broken out **in the notes** ("Other liabilities include
   finance lease obligations of X"). The code scans the notes for exactly this.
3. **Lease maturity note** — a table splitting lease liabilities into
   "within one year" vs "after one year".

### ⚠️ Confusions the tool flags for you

- **IFRS 16 "bare lease" (biggest one).** Many companies show only *"Lease
  liabilities (current / non-current)"* with no word "finance." That figure can
  **mix operating + finance** leases and cannot be split from the face of the
  statement. → flagged `IFRS-bare-lease (may include operating lease)`.
- **"Capital lease" = "finance lease"** (old US-GAAP term). Treated the same.
- **"Operating lease" is NOT a finance lease** → those lines are excluded.
- **Combined lines** like *"long-term debt and finance leases"* can't be split
  cleanly → flagged `combined debt+lease line`.
- **Buried in "other liabilities"** → flagged `verify in notes`.

---

## 2. How the code works (in simple words)

`finance_lease_extractor.py`:

1. **Reads** each report — PDF, TXT, or Excel — and turns it into plain text.
2. **Searches** every line for finance-lease keywords
   (`finance lease`, `capital lease`, `obligations under finance lease`, …).
3. **Excludes** any line that says `operating lease` / `right-of-use`.
4. **Decides current vs non-current** for each line using a 3-tier rule:
   - Tier 1: a definitive phrase on the line (`current portion`, `net of
     current`, `non-current`, `within one year`…) wins outright.
   - Tier 2: weaker words (`long-term`, `, current`) compared by position, with
     "non-current" masked so it can't be misread as "current".
   - Tier 3: the nearest **section heading** above the line
     (`CURRENT LIABILITIES` / `NON-CURRENT LIABILITIES`).
5. **Reads the number** — the left-most money figure (the current reporting
   year), skipping years (2024), standard references (`IFRS 16`, `ASC 842`), and
   durations (`1 year`).
6. **Outputs** a table + CSV: `Company | Short Term | Long Term | Flags`, with
   the tricky cases flagged for manual review.

### Run it

```bash
pip install -r requirements.txt          # only what your file types need

# a folder of reports
python finance_lease_extractor.py path/to/reports/ -o results.csv

# a single file
python finance_lease_extractor.py report.pdf

# add -v to see every matched line and its bucket
python finance_lease_extractor.py samples/ -v
```

Supported inputs: `.pdf`, `.txt`, `.xlsx`, `.xls`.

### Teaching it new wording

Company reports word things differently. To add a phrase, edit the keyword lists
at the top of `finance_lease_extractor.py`:
`FINANCE_LEASE_TERMS`, `EXCLUDE_TERMS`, `DEFINITIVE_CURRENT`,
`DEFINITIVE_NONCURRENT`, etc.

---

## 3. Verification

The `samples/` folder contains 5 realistic report snippets covering the tricky
cases. Running the tool on them reproduces the hand-checked answers in
`finance_lease_results.csv`:

| Sample | Short Term | Long Term | Flag |
|--------|-----------|-----------|------|
| US GAAP explicit | 34.5 | 128.7 | ok |
| IFRS bare lease | 47.3 | 210.6 | may include operating lease |
| Capital lease (old term) | 3,200 | 18,700 | ok |
| Buried in "other liabilities" | 9.4 | 41.2 | verify in notes |
| Combined debt+lease line | 145 | 4,560 | cannot split cleanly |

> Point the tool at your 20 real reports the same way:
> `python finance_lease_extractor.py your_reports_folder/ -o results.csv`
