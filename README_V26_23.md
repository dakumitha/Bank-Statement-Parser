# TaxForensic V26.23

This branch preserves the V26.23 Python desktop application source for review and testing.

## Source
- `src/TaxForensic_V26_23.py`

## Run locally
Use the existing TaxForensic Windows environment:

```bat
cd /d C:\TaxForensic
venv\Scripts\activate
python TaxForensic_V26_23.py --version
python TaxForensic_V26_23.py --corpus
python TaxForensic_V26_23.py --test
python TaxForensic_V26_23.py --analyze "C:\TaxForensic\5183349284.pdf"
python TaxForensic_V26_23.py --analyze "C:\TaxForensic\statement-31016477497.pdf"
```

The source should be copied from `src/TaxForensic_V26_23.py` to the application directory before running those commands.

## Important status
V26.23 worked well on the supplied Bank of India statement, where a demonstrated carried-forward discrepancy was present in the statement itself. It did not reliably extract the supplied SBI statement: the reported V26.23 analysis had 1,643 entries, 977 rows failing the extraction gate, and one balance discrepancy. Therefore, this branch is a preserved baseline, not a claim that SBI extraction is fixed.

Always retain printed balances as evidence, calculate running balances independently, and flag unresolved rows rather than guessing values.
