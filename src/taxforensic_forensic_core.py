"""Deterministic forensic primitives for TaxForensic. Standard library only."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from enum import Enum
from hashlib import sha256
import re
from typing import Iterable, Optional

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


def money(value: object) -> Decimal:
    """Parse money without binary floating-point arithmetic."""
    try:
        parsed = Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError, AttributeError) as exc:
        raise ValueError(f"Invalid monetary value: {value!r}") from exc
    if not parsed.is_finite():
        raise ValueError("Monetary value must be finite")
    return parsed.quantize(CENT)


def sha256_bytes(content: bytes) -> str:
    """Stable SHA-256 digest for original evidence bytes."""
    return sha256(content).hexdigest()


@dataclass(frozen=True)
class EvidenceRef:
    source_id: str
    file_name: str
    sha256: str
    page: Optional[int] = None
    locator: Optional[str] = None
    raw_excerpt: Optional[str] = None

    def __post_init__(self) -> None:
        if not self.source_id.strip() or not self.file_name.strip():
            raise ValueError("Evidence source_id and file_name are required")
        if not re.fullmatch(r"[0-9a-fA-F]{64}", self.sha256):
            raise ValueError("Evidence SHA-256 must be 64 hexadecimal characters")
        if self.page is not None and self.page < 1:
            raise ValueError("Evidence page numbers are 1-based")


class Direction(str, Enum):
    CREDIT = "credit"
    DEBIT = "debit"


class RecordStatus(str, Enum):
    CANDIDATE = "candidate"
    VERIFIED = "verified"
    REVIEW_REQUIRED = "review_required"
    REJECTED = "rejected"


@dataclass(frozen=True)
class FinancialEvent:
    """Canonical event. Direction carries the sign; amount must stay positive."""
    event_id: str
    case_id: str
    account_id: str
    transaction_date: date
    amount: Decimal
    direction: Direction
    currency: str
    narration: str
    evidence: EvidenceRef
    status: RecordStatus = RecordStatus.CANDIDATE
    confidence: Optional[Decimal] = None
    counterparty: Optional[str] = None
    external_reference: Optional[str] = None
    instrument: str = "bank"

    def __post_init__(self) -> None:
        if not all(v.strip() for v in (self.event_id, self.case_id, self.account_id)):
            raise ValueError("event_id, case_id and account_id are required")
        if self.amount <= ZERO:
            raise ValueError("Amount must be positive; direction carries the sign")
        if not re.fullmatch(r"[A-Z]{3}", self.currency):
            raise ValueError("Currency must be a 3-letter uppercase code")
        if self.confidence is not None and not (Decimal("0") <= self.confidence <= Decimal("1")):
            raise ValueError("Confidence must be between 0 and 1")

    @property
    def signed_amount(self) -> Decimal:
        return self.amount if self.direction == Direction.CREDIT else -self.amount


@dataclass(frozen=True)
class BalanceReconciliation:
    opening_balance: Decimal
    calculated_closing_balance: Optional[Decimal]
    printed_closing_balance: Decimal
    difference: Optional[Decimal]
    status: str
    transaction_count: int
    excluded_count: int
    note: str


def reconcile_balance(opening_balance: object, events: Iterable[FinancialEvent],
                      printed_closing_balance: object, tolerance: object = CENT) -> BalanceReconciliation:
    """Compare opening + VERIFIED events to the printed closing balance."""
    opening, printed, tol = money(opening_balance), money(printed_closing_balance), abs(money(tolerance))
    all_events = list(events)
    eligible = [e for e in all_events if e.status == RecordStatus.VERIFIED]
    excluded = len(all_events) - len(eligible)
    currencies = {e.currency for e in eligible}
    if len(currencies) > 1:
        return BalanceReconciliation(opening, None, printed, None, "REVIEW_REQUIRED", 0, excluded,
                                     "Multiple currencies found; no cross-currency sum was attempted.")
    calculated = (opening + sum((e.signed_amount for e in eligible), ZERO)).quantize(CENT)
    difference = (calculated - printed).quantize(CENT)
    if excluded:
        status = "REVIEW_REQUIRED"
        note = "Unverified records were excluded. Resolve the review queue before relying on this comparison."
    elif abs(difference) <= tol:
        status = "MATCH"
        note = "Calculated closing balance agrees with printed closing balance within tolerance."
    else:
        status = "MISMATCH"
        note = "Investigate missing, duplicated, misread or out-of-period records."
    return BalanceReconciliation(opening, calculated, printed, difference, status, len(eligible), excluded, note)


def _normalise(text: str) -> str:
    return " ".join(re.findall(r"[a-z0-9]+", text.casefold()))


@dataclass(frozen=True)
class DuplicateGroup:
    signature: tuple
    event_ids: tuple[str, ...]
    evidence_ids: tuple[str, ...]
    disposition: str = "REVIEW_REQUIRED"


def find_duplicate_candidates(events: Iterable[FinancialEvent]) -> list[DuplicateGroup]:
    """Find exact duplicate candidates; never delete or merge records."""
    buckets: dict[tuple, list[FinancialEvent]] = {}
    for event in events:
        if event.status == RecordStatus.REJECTED:
            continue
        signature = (event.case_id, event.account_id, event.transaction_date.isoformat(),
                     event.direction.value, event.amount, event.currency, _normalise(event.narration))
        buckets.setdefault(signature, []).append(event)
    return [DuplicateGroup(sig, tuple(e.event_id for e in rows),
                           tuple(e.evidence.source_id for e in rows))
            for sig, rows in sorted(buckets.items(), key=lambda item: str(item[0])) if len(rows) > 1]


@dataclass(frozen=True)
class TransferMatch:
    debit_event_id: str
    credit_event_id: str
    amount: Decimal
    day_gap: int
    score: Decimal
    status: str = "CANDIDATE_REVIEW"
    reasons: tuple[str, ...] = ()


def find_transfer_candidates(events: Iterable[FinancialEvent], max_day_gap: int = 3) -> list[TransferMatch]:
    """Suggest possible transfers between accounts. Results always require human review."""
    if max_day_gap < 0:
        raise ValueError("max_day_gap cannot be negative")
    rows = [e for e in events if e.status == RecordStatus.VERIFIED]
    debits = [e for e in rows if e.direction == Direction.DEBIT]
    credits = [e for e in rows if e.direction == Direction.CREDIT]
    matches = []
    for debit in debits:
        for credit in credits:
            if debit.case_id != credit.case_id or debit.account_id == credit.account_id:
                continue
            if debit.currency != credit.currency or debit.amount != credit.amount:
                continue
            gap = abs((credit.transaction_date - debit.transaction_date).days)
            if gap > max_day_gap:
                continue
            reasons = ["Exact amount and currency", "Different accounts", f"Date gap: {gap} day(s)"]
            score = (Decimal("0.70") + Decimal("0.10") *
                     Decimal(max_day_gap - gap) / Decimal(max(1, max_day_gap))).quantize(Decimal("0.01"))
            if debit.external_reference and debit.external_reference == credit.external_reference:
                reasons.append("Matching external reference")
                score = min(Decimal("0.99"), score + Decimal("0.20"))
            elif set(_normalise(debit.narration).split()) & set(_normalise(credit.narration).split()):
                reasons.append("Narration contains shared token")
            matches.append(TransferMatch(debit.event_id, credit.event_id, debit.amount, gap, score,
                                         "CANDIDATE_REVIEW", tuple(reasons)))
    return sorted(matches, key=lambda m: (-m.score, m.day_gap, m.debit_event_id, m.credit_event_id))


class IssueStatus(str, Enum):
    IDENTIFIED = "identified"
    EVIDENCE_COLLECTED = "evidence_collected"
    EXPLANATION_RECEIVED = "explanation_received"
    ANALYSED = "analysed"
    RESOLVED = "resolved"
    PENDING = "pending"


@dataclass
class InvestigationIssue:
    issue_id: str
    case_id: str
    title: str
    description: str
    status: IssueStatus = IssueStatus.IDENTIFIED
    evidence_ids: list[str] = field(default_factory=list)
    explanation: Optional[str] = None
    analysis_note: Optional[str] = None
    outcome: Optional[str] = None

    def attach_evidence(self, evidence_id: str) -> None:
        if not evidence_id.strip():
            raise ValueError("Evidence ID cannot be empty")
        if evidence_id not in self.evidence_ids:
            self.evidence_ids.append(evidence_id)
        if self.status == IssueStatus.IDENTIFIED:
            self.status = IssueStatus.EVIDENCE_COLLECTED

    def record_explanation(self, explanation: str) -> None:
        if not explanation.strip():
            raise ValueError("Explanation cannot be empty")
        self.explanation = explanation.strip()
        self.status = IssueStatus.EXPLANATION_RECEIVED

    def record_analysis(self, note: str) -> None:
        if not note.strip():
            raise ValueError("Analysis note cannot be empty")
        self.analysis_note = note.strip()
        self.status = IssueStatus.ANALYSED

    def resolve(self, outcome: str) -> None:
        if not outcome.strip():
            raise ValueError("Resolution outcome cannot be empty")
        if self.status not in (IssueStatus.ANALYSED, IssueStatus.EXPLANATION_RECEIVED):
            raise ValueError("Analyse the issue or record an explanation before resolution")
        self.outcome = outcome.strip()
        self.status = IssueStatus.RESOLVED


@dataclass(frozen=True)
class AmountComparison:
    reported_amount: Decimal
    observed_amount: Decimal
    difference: Decimal
    status: str
    note: str


def compare_reported_amounts(reported: object, observed: object, tolerance: object = CENT) -> AmountComparison:
    """Compare figures without inferring taxable character or making a legal conclusion."""
    r, o, tol = money(reported), money(observed), abs(money(tolerance))
    diff = (o - r).quantize(CENT)
    if abs(diff) <= tol:
        return AmountComparison(r, o, diff, "MATCH", "Amounts agree within tolerance.")
    return AmountComparison(r, o, diff, "DIFFERENCE_REQUIRES_INVESTIGATION",
                            "Establish period, source, duplication, ownership and transaction character before drawing a conclusion.")
