# TaxForensic V26.27: Universal Structured Financial Ingestion (Milestone 1)

## What changed

V26.27 preserves the existing V26.26 PySide6 desktop application and adds a sibling, dependency-light structured-data engine for transaction exports. The GUI's CSV/TSV and supported Excel/JSON paths now try this engine first when a recognized structured financial ledger is found. Unknown layouts fall back to the previous parser.

The engine currently normalizes structured CSV, TSV, TXT, JSON, XLSX and XLSM exports into a richer common ledger. It recognizes bank, credit-card, broker/contract-note and VDA/crypto transaction families from column semantics rather than hard-coded bank names.

### New normalized evidence fields

- Original source filename, worksheet and row number
- Original header names and row values retained as JSON evidence
- Transaction date, reference, symbol/ISIN, exchange, side, quantity and unit price
- Gross value, fees, taxes, currency, asset quantity, blockchain/network, transaction hash and wallet endpoints when supplied
- Printed balance kept separate from calculated balance
- Per-row review flags and stable fingerprints for possible duplicates

### Conservative rules

- An amount in a narration/reference is never used as a transaction amount.
- An unmarked amount is not automatically assigned to debit or credit for card, broker or VDA rows.
- Crypto-token movements are not treated as fiat cash flow when the supplied currency matches the asset.
- Quantity × price is compared with the supplied gross value where all three exist; discrepancies are flagged, not silently corrected.
- Duplicate-looking rows are flagged but never automatically deleted.
- Missing values remain missing. This update does not weaken the existing PDF extraction gate or balance-reconciliation rules.

## Install side by side

Keep V26.23, V26.26 and any existing executable unchanged. Copy these files into the same folder:

- `TaxForensic_V26_27.py`
- `TaxForensic_Universal_Financial_Engine.py`

The full release ZIP also contains V26.26 for reference and the deterministic tests.

Run in Windows CMD from the installation folder:

```bat
python -m py_compile TaxForensic_V26_27.py TaxForensic_Universal_Financial_Engine.py
python -m unittest -v test_universal_financial_engine.py
python TaxForensic_V26_27.py --version
python TaxForensic_V26_27.py --corpus
python TaxForensic_V26_27.py --test
```

The full GUI/application commands require the same supported Python environment and dependencies as the existing application, including PySide6, PyMuPDF and openpyxl.

## Scope and limitations

This is the first universal-ingestion milestone, not a claim of universal extraction accuracy. The new structured-data engine does not yet replace the PDF coordinate parser for broker contract notes or VDA exchange PDFs. Existing PDF bank/credit-card extraction remains in place. Scanned statements, every international layout, all historic bank formats and all broker/exchange exports still require a representative test corpus and layout-specific validation through reusable grammar/layout logic.

No real SBI or Bank of India PDF was available in this development workspace. V26.27 has not been certified against the actual 128-page SBI statement or the known 56-page Bank of India statement. Do not use the synthetic test results as proof that either real statement is fixed. PySide6 is also unavailable in this workspace, so the full Qt application test suite and GUI could not be run here.
