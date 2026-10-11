"""Tests for source-keyed TaxForensic extraction benchmark metrics."""
import unittest
from decimal import Decimal

from src.taxforensic_benchmark import BenchmarkRow, evaluate_extraction


def row(key, day="2026-04-01", direction="credit", amount="100.00",
        currency="INR", narration="UPI transfer", accepted=True, traceable=True):
    return BenchmarkRow(key, day, direction, Decimal(amount) if amount is not None else None,
                        currency, narration, accepted, traceable)


class BenchmarkTests(unittest.TestCase):
    def test_perfect_extraction_scores_one(self):
        gold = [row("p1"), row("p2", amount="250.00", direction="debit")]
        report = evaluate_extraction(gold, list(gold))
        self.assertEqual(report.row_precision, Decimal("1.0000"))
        self.assertEqual(report.row_recall, Decimal("1.0000"))
        self.assertEqual(report.amount_accuracy, Decimal("1.0000"))
        self.assertEqual(report.traceability_rate, Decimal("1.0000"))

    def test_wrong_amount_is_counted_as_field_error(self):
        gold = [row("p1", amount="100.00")]
        actual = [row("p1", amount="1000.00")]
        report = evaluate_extraction(gold, actual)
        self.assertEqual(report.row_recall, Decimal("1.0000"))
        self.assertEqual(report.amount_accuracy, Decimal("0.0000"))

    def test_missing_and_extra_rows_reduce_recall_and_precision(self):
        gold = [row("p1"), row("p2")]
        actual = [row("p1"), row("extra")]
        report = evaluate_extraction(gold, actual)
        self.assertEqual(report.row_precision, Decimal("0.5000"))
        self.assertEqual(report.row_recall, Decimal("0.5000"))
        self.assertEqual(report.unmatched_expected_keys, ("p2",))
        self.assertEqual(report.unmatched_extracted_keys, ("extra",))

    def test_duplicate_source_keys_are_not_arbitrarily_paired(self):
        gold = [row("p1")]
        actual = [row("p1"), row("p1", amount="200.00")]
        report = evaluate_extraction(gold, actual)
        self.assertEqual(report.paired_rows, 0)
        self.assertEqual(report.duplicate_extracted_keys, ("p1",))

    def test_untraceable_accepted_rows_are_visible(self):
        report = evaluate_extraction([row("p1")], [row("p1", traceable=False)])
        self.assertEqual(report.traceability_rate, Decimal("0.0000"))
        self.assertEqual(report.accepted_without_traceability, 1)

    def test_narration_comparison_ignores_case_and_punctuation(self):
        gold = [row("p1", narration="UPI-REF ABC")]
        actual = [row("p1", narration="upi ref abc")]
        self.assertEqual(evaluate_extraction(gold, actual).narration_accuracy, Decimal("1.0000"))


if __name__ == "__main__":
    unittest.main()
