# TaxForensic Extraction Benchmark Protocol

## Purpose

A universal parser must be evaluated against labelled source rows, not only against hand-picked examples or a list of supported banks. The benchmark module is deterministic and independent of AI.

## Source-row identity

Each labelled and extracted record must carry a stable row key derived from evidence provenance, such as a source-document hash plus page and row locator. The key must not be generated from the amount/date fields being measured. If source locators are missing or duplicated, the benchmark exposes that ambiguity rather than guessing a pairing.

## Metrics

- Row precision and recall, with F1.
- Date, direction, amount, currency and narration accuracy on paired rows.
- Traceability rate across extracted rows.
- Precision among records marked accepted.
- Count of accepted records without source traceability.
- Duplicate source keys and unmatched expected/extracted keys.

A row can be detected but have a wrong amount. Therefore row detection and field accuracy are reported separately. Narration accuracy uses case- and punctuation-normalised text; amounts are compared at two decimal places.

## Benchmark gates to adopt after a real corpus exists

Do not choose arbitrary numeric thresholds before the baseline corpus is labelled. Establish separate results for digital PDFs, scanned PDFs, CSV/XLSX exports, old passbooks, wrapped narrations, multilingual layouts, multi-page statements and credit-card statements. Report results per format family and overall so a good average cannot conceal a failing category.

At minimum, publish:
- Exact extraction precision and recall.
- Amount/date/direction accuracy.
- Printed-vs-calculated closing-balance reconciliation rate.
- Accepted records lacking evidence locators.
- False acceptance rate and unresolved review rate.
- Runtime and memory on the supported low-resource machine profile.

Synthetic unit tests are necessary but are not proof of universal real-world accuracy.
