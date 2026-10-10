# Bank Statement Parser / TaxForensic

A Windows desktop application for financial-statement extraction and forensic reconciliation.

## Current development release

The current working source package is TaxForensic V26.25. The source is a Python/PySide6 desktop application; it is not a standalone executable. See the release README and test report before use.

### Core goals

- Extract bank and credit-card statement transactions from PDFs and supported structured formats.
- Keep printed statement balances separate from formula-calculated balances.
- Reconcile transactions against available statement checkpoints.
- Flag missing, ambiguous, or inconsistent values for review instead of guessing.
- Preserve source evidence and traceability.

## Important status

This repository has just been initialized. The complete V26.25 source package and its accompanying test report still need to be committed. Do not treat this README-only commit as the application upload or as a validated release.

## Local development

Use a supported Python environment and install dependencies from a reviewed dependency list. The application currently relies on PySide6, PyMuPDF, and openpyxl. Optional OCR components are not required for the core parser.

Before using the application for financial or tax work, run the included deterministic corpus and headless analysis against the specific statements being processed. Passing synthetic tests alone does not establish correctness on every bank's statement format.
