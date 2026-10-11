"""Unit tests for TaxForensic deterministic forensic primitives."""
import unittest
from datetime import date
from decimal import Decimal

from src.taxforensic_forensic_core import (
    Direction, EvidenceRef, FinancialEvent, RecordStatus, InvestigationIssue,
    IssueStatus, compare_reported_amounts, find_duplicate_candidates,
    find_transfer_candidates, money, reconcile_balance,
)

HASH = "a" * 64


def event(event_id, account, day, amount, direction, narration="TRANSFER",
          status=RecordStatus.VERIFIED, source_id=None, external_reference=None, currency="INR"):
    return FinancialEvent(
        event_id, "CASE-1", account, date.fromisoformat(day), money(amount), Direction(direction),
        currency, narration,
        EvidenceRef(source_id or f"SRC-{event_id}", f"{event_id}.pdf", HASH, page=1, locator="row 1"),
        status=status, external_reference=external_reference,
    )


class MoneyTests(unittest.TestCase):
    def test_decimal_rounding_and_commas(self):
        self.assertEqual(money("1,234.567"), Decimal("1234.57"))

    def test_rejects_nan(self):
        with self.assertRaises(ValueError):
            money("NaN")


class BalanceTests(unittest.TestCase):
    def test_only_verified_records_enter_arithmetic_and_review_is_preserved(self):
        rows = [
            event("1", "A", "2026-04-01", "100", "credit"),
            event("2", "A", "2026-04-02", "25", "debit"),
            event("3", "A", "2026-04-03", "500", "credit", status=RecordStatus.REVIEW_REQUIRED),
        ]
        result = reconcile_balance("1000", rows, "1075")
        self.assertEqual(result.calculated_closing_balance, Decimal("1075.00"))
        self.assertEqual(result.printed_closing_balance, Decimal("1075.00"))
        self.assertEqual(result.status, "REVIEW_REQUIRED")
        self.assertEqual(result.excluded_count, 1)

    def test_mismatch_is_not_silently_fixed(self):
        result = reconcile_balance("1000", [event("1", "A", "2026-04-01", "100", "credit")], "1090")
        self.assertEqual(result.status, "MISMATCH")
        self.assertEqual(result.difference, Decimal("10.00"))

    def test_mixed_currency_refuses_to_sum(self):
        result = reconcile_balance("1000", [
            event("1", "A", "2026-04-01", "100", "credit", currency="INR"),
            event("2", "A", "2026-04-02", "100", "credit", currency="USD"),
        ], "1200")
        self.assertEqual(result.status, "REVIEW_REQUIRED")
        self.assertIsNone(result.calculated_closing_balance)


class MatchingTests(unittest.TestCase):
    def test_duplicates_remain_review_candidates(self):
        rows = [event("1", "A", "2026-04-01", "100", "debit", "UPI abc"),
                event("2", "A", "2026-04-01", "100", "debit", "UPI ABC")]
        groups = find_duplicate_candidates(rows)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].disposition, "REVIEW_REQUIRED")
        self.assertEqual(set(groups[0].event_ids), {"1", "2"})

    def test_rejected_record_not_in_duplicate_scan(self):
        rows = [event("1", "A", "2026-04-01", "100", "debit"),
                event("2", "A", "2026-04-01", "100", "debit", status=RecordStatus.REJECTED)]
        self.assertEqual(find_duplicate_candidates(rows), [])

    def test_transfer_candidate_requires_review_and_can_use_reference(self):
        rows = [
            event("1", "A", "2026-04-01", "1000", "debit", "IMPS ABC123", external_reference="ABC123"),
            event("2", "B", "2026-04-02", "1000", "credit", "IMPS ABC123", external_reference="ABC123"),
        ]
        matches = find_transfer_candidates(rows)
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].status, "CANDIDATE_REVIEW")
        self.assertIn("Matching external reference", matches[0].reasons)

    def test_same_account_is_not_inter_account_transfer(self):
        rows = [event("1", "A", "2026-04-01", "1000", "debit"),
                event("2", "A", "2026-04-02", "1000", "credit")]
        self.assertEqual(find_transfer_candidates(rows), [])

    def test_currency_must_match(self):
        rows = [event("1", "A", "2026-04-01", "1000", "debit", currency="INR"),
                event("2", "B", "2026-04-02", "1000", "credit", currency="USD")]
        self.assertEqual(find_transfer_candidates(rows), [])


class IssueRegisterTests(unittest.TestCase):
    def test_resolution_requires_investigative_step(self):
        issue = InvestigationIssue("I-1", "CASE-1", "Credit mismatch", "Reported vs observed")
        with self.assertRaises(ValueError):
            issue.resolve("unexplained income")
        issue.attach_evidence("SRC-1")
        issue.record_explanation("Loan repayment")
        issue.record_analysis("Verify agreement and prior bank debit.")
        issue.resolve("Resolved after source verification")
        self.assertEqual(issue.status, IssueStatus.RESOLVED)
        self.assertEqual(issue.evidence_ids, ["SRC-1"])

    def test_amount_difference_is_not_tax_conclusion(self):
        result = compare_reported_amounts("100", "150")
        self.assertEqual(result.status, "DIFFERENCE_REQUIRES_INVESTIGATION")
        self.assertIn("before drawing a conclusion", result.note)


if __name__ == "__main__":
    unittest.main()
