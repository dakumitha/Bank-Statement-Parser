"""Deterministic, source-locator-based extraction benchmark metrics."""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
import re
from typing import Iterable, Optional


def _normalise_text(value: str) -> str:
    return " ".join(re.findall(r"[a-z0-9]+", value.casefold()))


def _ratio(numerator: int, denominator: int) -> Decimal:
    if denominator == 0:
        return Decimal("1.00") if numerator == 0 else Decimal("0.00")
    return (Decimal(numerator) / Decimal(denominator)).quantize(Decimal("0.0001"))


@dataclass(frozen=True)
class BenchmarkRow:
    """One labelled or extracted row, identified by a stable source locator.

    row_key should identify the original source row (for example source hash + page +
    row locator), not be derived from the amount/date being evaluated.
    """
    row_key: str
    transaction_date: Optional[str]
    direction: Optional[str]
    amount: Optional[Decimal]
    currency: Optional[str]
    narration: Optional[str]
    accepted: bool = True
    traceable: bool = True

    def __post_init__(self) -> None:
        if not self.row_key.strip():
            raise ValueError("row_key must identify a source row")
        if self.amount is not None and not self.amount.is_finite():
            raise ValueError("Benchmark amounts must be finite")


@dataclass(frozen=True)
class BenchmarkReport:
    expected_rows: int
    extracted_rows: int
    paired_rows: int
    row_precision: Decimal
    row_recall: Decimal
    row_f1: Decimal
    date_accuracy: Decimal
    direction_accuracy: Decimal
    amount_accuracy: Decimal
    currency_accuracy: Decimal
    narration_accuracy: Decimal
    traceability_rate: Decimal
    accepted_row_precision: Decimal
    accepted_without_traceability: int
    duplicate_expected_keys: tuple[str, ...]
    duplicate_extracted_keys: tuple[str, ...]
    unmatched_expected_keys: tuple[str, ...]
    unmatched_extracted_keys: tuple[str, ...]


def evaluate_extraction(expected: Iterable[BenchmarkRow],
                        extracted: Iterable[BenchmarkRow]) -> BenchmarkReport:
    """Compare extraction against labelled rows without using AI or bank-specific rules.

    Rows pair only on unique source-row keys. Amount/date similarity is deliberately
    not used as identity, so wrong extracted values remain measurable field errors.
    """
    gold, actual = list(expected), list(extracted)
    gold_map: dict[str, list[BenchmarkRow]] = {}
    actual_map: dict[str, list[BenchmarkRow]] = {}
    for row in gold:
        gold_map.setdefault(row.row_key, []).append(row)
    for row in actual:
        actual_map.setdefault(row.row_key, []).append(row)

    dup_gold = tuple(sorted(k for k, v in gold_map.items() if len(v) > 1))
    dup_actual = tuple(sorted(k for k, v in actual_map.items() if len(v) > 1))
    paired: list[tuple[BenchmarkRow, BenchmarkRow]] = []
    for key in sorted(set(gold_map) & set(actual_map)):
        if len(gold_map[key]) == 1 and len(actual_map[key]) == 1:
            paired.append((gold_map[key][0], actual_map[key][0]))

    paired_keys = {g.row_key for g, _ in paired}
    ambiguous = set(dup_gold) | set(dup_actual)
    unmatched_gold = tuple(sorted(set(gold_map) - paired_keys | ambiguous))
    unmatched_actual = tuple(sorted(set(actual_map) - paired_keys | ambiguous))
    row_precision = _ratio(len(paired), len(actual))
    row_recall = _ratio(len(paired), len(gold))
    f1 = Decimal("0.00") if row_precision + row_recall == 0 else (
        Decimal(2) * row_precision * row_recall / (row_precision + row_recall)
    ).quantize(Decimal("0.0001"))

    def field_accuracy(field_name: str, normalise=lambda x: x):
        correct = 0
        for g, a in paired:
            gv, av = getattr(g, field_name), getattr(a, field_name)
            if gv is None or av is None:
                continue
            if normalise(gv) == normalise(av):
                correct += 1
        # Missing values count as incorrect; every paired row is in the denominator.
        return _ratio(correct, len(paired))

    accepted_rows = [row for row in actual if row.accepted]
    accepted_matched = sum(1 for _, row in paired if row.accepted)
    accepted_precision = _ratio(accepted_matched, len(accepted_rows))
    no_trace = sum(1 for row in accepted_rows if not row.traceable)
    return BenchmarkReport(
        expected_rows=len(gold),
        extracted_rows=len(actual),
        paired_rows=len(paired),
        row_precision=row_precision,
        row_recall=row_recall,
        row_f1=f1,
        date_accuracy=field_accuracy("transaction_date"),
        direction_accuracy=field_accuracy("direction", lambda x: str(x).casefold()),
        amount_accuracy=field_accuracy("amount", lambda x: Decimal(x).quantize(Decimal("0.01"))),
        currency_accuracy=field_accuracy("currency", lambda x: str(x).upper()),
        narration_accuracy=field_accuracy("narration", lambda x: _normalise_text(str(x))),
        traceability_rate=_ratio(sum(1 for row in actual if row.traceable), len(actual)),
        accepted_row_precision=accepted_precision,
        accepted_without_traceability=no_trace,
        duplicate_expected_keys=dup_gold,
        duplicate_extracted_keys=dup_actual,
        unmatched_expected_keys=unmatched_gold,
        unmatched_extracted_keys=unmatched_actual,
    )
