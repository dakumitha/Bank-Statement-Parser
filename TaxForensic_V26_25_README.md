# TaxForensic V26.25

## Purpose

V26.25 is an updated full Python application source based on V26.23. It targets statements whose transaction date/narration and financial cells are printed on different vertical baselines, including multi-line transaction layouts. It is not a compiled Windows EXE.

## Changes

1. Numeric tokens retain their calibrated money-column role in the parser evidence candidates.
2. A numeric amount cell below the transaction-date baseline is not automatically discarded into narration if its right edge is inside a calibrated debit, credit, or balance column.
3. A missing movement amount can be restored only when exactly one non-baseline amount token exists in a calibrated debit/credit column for that transaction. Multiple candidate cells are recorded as a conflict and left for review.
4. A misassigned movement can be corrected only when a visible amount token in the correct calibrated money column exactly equals the difference between consecutive printed running balances. The app does not manufacture an amount from the balance equation.
5. Page B/F-to-C/F checkpoint logic and the strict extraction gate remain enabled.
6. The implementation uses column geometry and balance evidence, not an SBI-specific bank-name exception.

## Install

Keep previous source and EXE files as backups. Copy `TaxForensic_V26_25.py` into `C:\\TaxForensic`, then run:

```bat
cd /d C:\\TaxForensic
venv\\Scripts\\activate
python TaxForensic_V26_25.py --version
python TaxForensic_V26_25.py --corpus
python TaxForensic_V26_25.py --test
python TaxForensic_V26_25.py --analyze "C:\\TaxForensic\\statement-31016477497.pdf"
```

## Important limitations

- The actual SBI PDF was not available in the development environment, so V26.25 has not been accuracy-tested against that real statement.
- Syntax compilation and isolated deterministic tests passed for multiline column recovery, exact running-balance token correction, checkpoint recovery, and refusal of mismatched checkpoints.
- The complete PySide6 GUI could not be launched in the development environment because PySide6 was unavailable there.
- Ambiguous cells are deliberately not guessed. Keep the original EXE until your real SBI report and GUI have been checked.

## Please report

After running `--analyze`, share the final summary and status lines beginning with `Multi-line column recovery:`. Compare rows failing the extraction gate and balance discrepancies with the previous V26.23 output.
