# TaxForensic Forensic Platform: V26.28 Foundation

## Product rule

TaxForensic is a deterministic financial-forensics application. AI is optional and must never be the authority for extracted amounts, arithmetic, matching, tax computations or issue outcomes.

## What this increment adds

The module src/taxforensic_forensic_core.py provides a standard-library-only first layer:

- Canonical financial event records with source-evidence references.
- Decimal-based money parsing.
- Reconciliation of opening balance plus verified transactions against the printed closing balance.
- Exact duplicate candidates, without deleting or merging records.
- Inter-account transfer candidates, kept in a review state.
- An issue register with an explicit investigation lifecycle.
- Reported-versus-observed amount comparison that does not make a tax or legal conclusion.

Existing desktop application and previous versioned source files are not modified by this increment.

## Evidence rules

1. Preserve original evidence and retain SHA-256, filename, source identifier, page/row locator and extraction metadata.
2. Printed closing balances and calculated closing balances remain separate.
3. Only verified events enter balance arithmetic. If any event is excluded for review, the reconciliation is marked REVIEW_REQUIRED even if partial arithmetic matches.
4. Duplicate and transfer detection produce candidates, never automatic deletions, merges or declarations of common ownership.
5. A difference is an investigative lead, not proof of income or concealment.
6. Never sum transactions across currencies without an explicit, evidenced conversion record.
7. Every candidate retains its source and status; uncertainty must not silently disappear from totals.

## Run tests

From the repository root with Python 3.10 or later:

    python -m unittest discover -s tests -p "test_*.py" -v

These are deterministic unit tests. They do not establish that all PDF layouts, scans, banks or financial instruments are supported. Real-world extraction needs a labelled benchmark corpus and independent reconciliation.

## Planned build sequence

1. Evidence intake, immutable original storage, hashing and extraction audit.
2. Candidate transaction ledger and human review queue with source locators.
3. Reconciliation and duplicate/transfer candidate workflows.
4. Multi-document event matching across banks, cards, brokers, books, ITR, AIS, TIS, Form 26AS and supporting records.
5. Fund-flow graph and time-ordered source/use-of-funds analysis.
6. Case issue register, explanation tracking and evidence-linked export.
7. Assessment-year-versioned tax rules and calculations, separate from factual forensic analysis.
8. GUI integration, packaging, offline verification and regression corpus.

## Limitations in this increment

This is a core library, not yet an integrated end-user GUI workflow. It does not parse arbitrary PDFs, establish account ownership, identify beneficial ownership, determine taxable income, or replace an officer's evaluation of evidence. Those features require separately tested modules and explicit source evidence.
