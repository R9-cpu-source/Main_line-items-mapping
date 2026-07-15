#!/usr/bin/env python3
"""Repeatable verification for finance_lease_extractor.py.

Regenerates the fixtures, runs the extractor over them (PyMuPDF engine so it
works without Poppler), and asserts the extracted short/long-term finance-lease
values equal the known answers in make_fixtures.EXPECTED.

Run:  python tests/verify.py
Exit code 0 = all pass.
"""
from __future__ import annotations

import csv
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

import make_fixtures  # noqa: E402

TOL = 0.01


def _num(s: str):
    return None if s in ("", None) else float(s)


def main() -> int:
    make_fixtures.main()
    with tempfile.TemporaryDirectory() as out:
        subprocess.run(
            [sys.executable, str(ROOT / "finance_lease_extractor.py"),
             str(make_fixtures.FIX), "--engine", "pymupdf", "--output-dir", out],
            check=True, cwd=ROOT,
        )
        rows = {r["file_name"]: r
                for r in csv.DictReader(open(Path(out) / "finance_lease_results.csv", encoding="utf-8-sig"))}

    ok = True
    for fname, (exp_st, exp_lt) in make_fixtures.EXPECTED.items():
        row = rows.get(fname)
        if row is None:
            print(f"FAIL {fname}: not in results"); ok = False; continue
        got_st = _num(row["finance_lease_short_term_usd_mm"])
        got_lt = _num(row["finance_lease_long_term_usd_mm"])

        def eq(a, b):
            if a is None or b is None:
                return a is None and b is None
            return abs(a - b) <= TOL

        status = "PASS" if eq(got_st, exp_st) and eq(got_lt, exp_lt) else "FAIL"
        if status == "FAIL":
            ok = False
        print(f"{status} {fname:28} expected ST={exp_st} LT={exp_lt} | got ST={got_st} LT={got_lt}")

    print("\nALL PASS" if ok else "\nSOME FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
