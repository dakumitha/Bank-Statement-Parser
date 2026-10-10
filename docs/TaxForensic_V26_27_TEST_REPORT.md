# TaxForensic V26.27 Test Report

## Local checks

- `python -m py_compile TaxForensic_V26_27.py TaxForensic_Universal_Financial_Engine.py test_universal_financial_engine.py`: **PASS**
- `python -m unittest -v test_universal_financial_engine.py`: **PASS, 9 tests**
- Synthetic bank CSV with explicit debit/credit/running balance: **PASS**
- Ambiguous bank amount with no direction evidence: **PASS, flagged rather than guessed**
- Credit-card CSV detection and unmarked-amount caution: **PASS**
- Broker trade mapping including symbol, ISIN, quantity, price, gross and mismatch flag: **PASS**
- VDA export mapping including token quantity, network and transaction hash: **PASS**
- VDA asset movement kept separate from fiat cash flow when currency equals asset: **PASS**
- Possible duplicate detection retains both rows and flags them: **PASS**
- Metadata before headers preserves source row number: **PASS**
- JSON broker export normalization: **PASS**
- Locale amount parsing for Indian grouping, European decimal punctuation and parenthesized negatives: **PASS**

## Not run / not verified

- Full `TaxForensic_V26_27.py --test`: **NOT RUN** because PySide6 is unavailable in this environment.
- GUI launch and workbook review: **NOT RUN**.
- Real SBI 128-page PDF: **NOT AVAILABLE; NOT VERIFIED**.
- Real Bank of India 56-page PDF: **NOT AVAILABLE; NOT VERIFIED**.
- Broker contract-note PDFs and VDA exchange PDFs: **NOT YET ROUTED THROUGH THE NEW STRUCTURED ENGINE**.
- Scanned/OCR and international-language layout coverage: **NOT CERTIFIED**.
- GitHub Actions: check results must be read from the actual workflow runs; no CI success is implied by these local checks.

## Interpretation

These tests verify the new deterministic structured-data mapping and safety flags against synthetic fixtures. They do not prove accuracy on arbitrary real-world documents. The next critical step is to test the actual SBI and Bank of India PDFs and add representative contract-note, card and VDA fixtures, including visually verified expected values.
