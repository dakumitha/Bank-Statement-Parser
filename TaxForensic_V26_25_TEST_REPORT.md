# TaxForensic V26.25 test report

## Checks run in the development environment

- `python -m py_compile TaxForensic_V26_25.py`: PASS
- Existing isolated checkpoint recovery regression: PASS
- Existing mismatched-checkpoint refusal regression: PASS
- Existing zero-amount transaction rejection regression: PASS
- Existing calibrated right-edge column-routing tests: PASS
- Synthetic non-baseline deposit candidate recovery: PASS
- Synthetic exact printed-balance delta correction: PASS

## Synthetic examples

- A `100.00` candidate positioned below the date baseline but inside the calibrated credit column was recovered as a credit when it was the only candidate.
- A row with printed balance changing from `7003.85` to `911.85` was corrected from an incorrectly assigned `123.60` debit to the visible `6092.00` debit token only because that token was explicitly classified in the debit column and exactly matched the printed balance delta.
- Recovery functions refuse ambiguous multi-token candidates rather than selecting one arbitrarily.

## Not tested

- The real `statement-31016477497.pdf` was not present in the development container. No claim is made that its 977 extraction errors or ₹5,968.40 discrepancy have been resolved.
- The Qt GUI was not launched because PySide6 is unavailable in the development environment.

## Verify locally

Run `--test` and then `--analyze` against the real SBI PDF from the existing Windows virtual environment. Compare rows failing the extraction gate and balance discrepancies with the previous V26.23 output. Inspect the original printed row before treating any discrepancy as a parser error.
