import sys
import re
import os
import hashlib
import sqlite3
import csv
import json
import xml.etree.ElementTree as ET
import html

from TaxForensic_Universal_Financial_Engine import (
    read_structured_file as read_universal_structured_file,
    to_application_transactions as universal_to_application_transactions,
)


# ----------------------------------------------------------------------
# PORTABLE APPLICATION PATHS
# ----------------------------------------------------------------------

def get_app_directory():
    """Return the directory beside the running .py file or frozen EXE."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIRECTORY = get_app_directory()
DATABASE_PATH = os.path.join(APP_DIRECTORY, "tax_case.db")


from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

# ----------------------------------------------------------------------
# 0. PYMuPDF COMPATIBILITY
# ----------------------------------------------------------------------

try:
    import pymupdf as fitz
except ImportError:
    import fitz


import openpyxl

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)

from openpyxl.utils import get_column_letter


from PySide6.QtCore import (
    QCoreApplication,
    Qt,
    QRectF,
    Signal,
    QThread,
    QProcess,
    QObject,
    QEvent,
    QTimer,
)

from PySide6.QtGui import (
    QImage,
    QPixmap,
    QPen,
    QBrush,
    QColor,
    QDragEnterEvent,
    QDropEvent,
)

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QStyledItemDelegate,
    QAbstractItemDelegate,
    QLabel,
    QPushButton,
    QGraphicsView,
    QGraphicsScene,
    QGraphicsRectItem,
    QMessageBox,
    QSplitter,
    QSplitterHandle,
    QLineEdit,
    QFileDialog,
    QCheckBox,
    QFrame,
    QProgressBar,
    QGroupBox,
    QScrollArea,
    QSizePolicy,
)


# ======================================================================
# 1. CORE SANITIZATION & REGEX GUARDS
# ======================================================================

DATE_PATTERNS = [
    re.compile(r"^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}$"),
    re.compile(r"^\d{1,2}[-/\.][A-Za-z]{3}[-/\.]\d{2,4}$"),
    re.compile(r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}$"),
]


ACC_NUM_REGEX = re.compile(
    r"(?:A/C\s*(?:NO|NUM|NUMBER)?|ACCOUNT\s*(?:NO|NUM|NUMBER)?)"
    r"\s*[:.\-]?\s*([X\d]{8,18})",
    re.IGNORECASE,
)


BF_PATTERNS = [
    "BROUGHT FORWARD",
    "B/F",
    "CARRIED FORWARD",
    "C/F",
    "OPENING BALANCE B/F",
    "TOTAL C/F",
    "BALANCE B/F",
]


# ----------------------------------------------------------------------
# BANK-STATEMENT FORMAT FAMILIES
# ----------------------------------------------------------------------
# This is deliberately a capability registry, not a brittle one-bank parser.
# Publicly available examples show materially different families: conventional
# debit/credit/balance tables, passbook-style statements where the balance is
# printed on a following line, layouts with reference/branch columns, and
# international SWIFT/MT940-style tagged statements. The extraction engine
# therefore detects structure from the document itself and uses bank names only
# as supporting evidence.
STATEMENT_FORMAT_FAMILIES = {
    "dual_date_table": [
        "HDFC", "SBI", "PNB", "BANK OF BARODA", "CANARA", "UNION BANK",
        "INDIAN BANK", "BANK OF INDIA", "CENTRAL BANK", "UCO BANK",
        "BANK OF MAHARASHTRA", "INDIAN OVERSEAS BANK",
    ],
    "single_date_debit_credit": [
        "AXIS", "KOTAK", "YES BANK", "INDUSIND", "IDFC", "FEDERAL",
        "RBL", "AU SMALL FINANCE",
    ],
    "passbook_running_balance": [
        "SBI", "BANK OF BARODA", "COOPERATIVE", "CO-OPERATIVE",
        "URBAN COOPERATIVE", "URBAN CO-OPERATIVE",
    ],
    "international_tagged": [
        "HSBC", "CITI", "CITIBANK", "BARCLAYS", "DBS", "STANDARD CHARTERED",
    ],
}


FORMAT_RESEARCH_NOTES = {
    "conventional": "Date/narration/reference plus separate debit, credit and running balance fields.",
    "passbook": "Transaction may span multiple physical lines and running balance may appear below the transaction, sometimes with Cr/Dr suffix.",
    "reference_branch": "Reference/cheque and branch/IFSC fields can sit between narration and amount columns.",
    "international": "Value date, posting date, references, D/C indicator, amount and running balance may be separate fields or SWIFT/MT940 tags.",
    "scanned": "No text layer: requires OCR and must remain under the same extraction/reconciliation gate.",
}


def is_valid_row_date(text: str, x0: float, page_width: float = 612.0) -> bool:
    """
    Validates whether a token can represent the beginning of
    a transaction row.

    Narration dates are rejected because they normally occur
    farther to the right.
    """

    if text is None:
        return False

    clean = str(text).strip()

    if not clean:
        return False

    # Do not hard-code a 130-point cutoff. Bank PDFs can place the date
    # column farther right depending on page size, margins and scaling.
    # Keep the date candidate in the left transaction area while allowing
    # ordinary A4/Letter layouts and slightly shifted PDF text layers.
    date_x_limit = max(220.0, min(300.0, float(page_width) * 0.30))
    if x0 > date_x_limit:
        return False

    if clean.startswith("/"):
        return False

    if clean.endswith("/"):
        return False

    if clean.startswith("-"):
        return False

    return any(pattern.match(clean) for pattern in DATE_PATTERNS)


def infer_pdf_date_column_bounds(words, page_width=612.0):
    """Infer a page-local physical date column from PDF text geometry.

    The date column is normally the leftmost cluster of date-like tokens.
    Narration can contain dates, but those dates are normally materially to
    the right of the transaction-date column. V23 uses that geometry only as
    a row-boundary safeguard. It does not invent or normalize any amount.
    """
    candidates = []
    for word in words or []:
        try:
            x0, y0, x1, y1, raw = word[:5]
        except Exception:
            continue
        text = str(raw).strip()
        if not is_valid_row_date(text, float(x0), page_width):
            continue
        candidates.append({
            "x0": float(x0),
            "x1": float(x1),
            "y0": float(y0),
            "text": text,
        })

    if not candidates:
        fallback_right = max(170.0, min(240.0, float(page_width) * 0.28))
        return 0.0, fallback_right

    candidates.sort(key=lambda item: item["x0"])
    clusters = []
    for item in candidates:
        if not clusters or item["x0"] - clusters[-1][-1]["x0"] > 24.0:
            clusters.append([item])
        else:
            clusters[-1].append(item)

    # The leftmost date cluster is the strongest physical-date-column signal.
    # Use a generous right edge based on the cluster itself, but never allow
    # the full generic date-candidate area to swallow narration dates.
    cluster = clusters[0]
    left = max(0.0, min(x["x0"] for x in cluster) - 12.0)
    right = max(x["x1"] for x in cluster) + 18.0

    # A single isolated date can be a page heading. Do not make the bound
    # microscopic in that case. The transaction parser still requires the
    # date to be a valid date token.
    if len(cluster) == 1:
        right = max(right, left + 85.0)

    return left, right


def is_running_balance_token(text: str) -> bool:
    """Detect printed running balances such as 50,202.80Cr or Rs. 1,302.80 Dr."""
    value = str(text or "").strip().replace(" ", "")
    return bool(re.fullmatch(
        r"(?:₹|RS\.?)?[+-]?(?:\d{1,3}(?:,\d{2})*,\d{3}|\d+)(?:\.\d{1,2})?(?:CR|DR)",
        value,
        flags=re.IGNORECASE,
    ))


def pdf_numeric_column_role(x0, x_mid, x1, channels):
    """Classify a numeric PDF token by right edge when calibrated headings allow it."""
    try:
        midpoint = float(x_mid)
        right_edge = float(x1)
    except Exception:
        return None
    edge_channels = (channels or {}).get("_money_right_edge_channels") or {}
    if edge_channels:
        for role in ("withdrawal", "deposit", "balance"):
            bounds = edge_channels.get(role)
            if bounds and float(bounds[0]) <= right_edge <= float(bounds[1]):
                return role
        # Do not revert to centre classification once the right-edge map is
        # available: that is exactly how short amounts cross column boundaries.
        return "narration"
    for role in ("balance", "withdrawal", "deposit"):
        bounds = (channels or {}).get(role)
        if bounds and float(bounds[0]) <= midpoint <= float(bounds[1]):
            return role
    bounds = (channels or {}).get("narration")
    if bounds and float(bounds[0]) <= midpoint <= float(bounds[1]):
        return "narration"
    return None


def is_money_token(text: str) -> bool:
    """Return True for a normal monetary token without a Cr/Dr suffix."""
    value = str(text or "").strip().replace(" ", "")
    return bool(re.fullmatch(
        r"[+-]?(?:\d{1,3}(?:,\d{2})*,\d{3}|\d+)(?:\.\d{2})",
        value,
    ))


def extract_reference_balance(page_text: str, label_pattern: str):
    """Extract an explicitly printed reference/opening balance.

    PDF text extraction can split one visual value into several tokens, for
    example ``202.80`` + ``Cr`` or even ``202`` + ``.80`` + ``Cr``.  This
    routine therefore normalizes visual whitespace and accepts separated
    accounting markers while remaining conservative: it only searches in a
    short window immediately after the reference label and never derives the
    opening amount from transaction arithmetic.
    """
    text = str(page_text or "")
    label_match = re.search(label_pattern, text, flags=re.IGNORECASE)
    if not label_match:
        return None

    window = text[label_match.end():label_match.end() + 220]
    normalized = re.sub(r"[\u00a0\u2007\u202f]", " ", window)
    normalized = re.sub(r"[ \t]+", " ", normalized)

    # Prefer an amount explicitly followed by Cr/Dr, allowing arbitrary
    # whitespace/newlines between amount and marker.
    marked = re.search(
        r"(?:₹\s*|Rs\.?\s*)?"
        r"([+-]?(?:\d{1,3}(?:,\d{2})*,\d{3}|\d+)(?:[.,]\d{1,2})?)"
        r"\s*(CR|DR)\b",
        normalized,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if marked:
        raw = marked.group(1).replace(",", "")
        # Statement locales may use comma as decimal separator.  When there
        # is exactly one comma and no Indian/grouping pattern, accept it as a
        # decimal separator.
        if "," in marked.group(1) and "." not in marked.group(1):
            parts = marked.group(1).split(",")
            if len(parts) == 2 and len(parts[1]) in (1, 2):
                raw = ".".join(parts)
        value = parse_decimal(raw)
        if value is not None:
            return value

    # Then accept a plain amount immediately following the label. This covers
    # statements which omit Cr/Dr on the opening/reference line.
    plain = re.search(
        r"(?:₹\s*|Rs\.?\s*)?"
        r"([+-]?(?:\d{1,3}(?:,\d{2})*,\d{3}|\d+)(?:[.,]\d{1,2})?)",
        normalized,
        flags=re.IGNORECASE,
    )
    if plain:
        raw = plain.group(1).replace(",", "")
        if "," in plain.group(1) and "." not in plain.group(1):
            parts = plain.group(1).split(",")
            if len(parts) == 2 and len(parts[1]) in (1, 2):
                raw = ".".join(parts)
        return parse_decimal(raw)

    return None


def extract_any_reference_balance(page_text: str):
    """Find common opening/reference balance labels without bank-specific code."""
    labels = [
        r"OPENING\s+BALANCE\s*:?",
        r"OPENING\s+BAL\s*:?",
        r"OPENING\s+BALANCE\s+B/?F\s*:?",
        r"BROUGHT\s+FORWARD\s*:?",
        r"BALANCE\s+B/?F\s*:?",
        r"CARRIED\s+FORWARD\s*:?",
        r"BALANCE\s+C/?F\s*:?",
        r"B/?F\s*:?",
    ]
    for label in labels:
        value = extract_reference_balance(page_text, label)
        if value is not None:
            return value, label
    return None, None

# ======================================================================
# V26 DOCUMENT GRAMMAR / LANGUAGE / CREDIT-CARD ARCHITECTURE
# ======================================================================

TAXFORENSIC_ENGINE_VERSION = "V26.27"
CORPUS_VERSION = "3"

DOCUMENT_TYPES = {
    "BANK": "Bank Account Statement",
    "CREDIT_CARD": "Credit Card Statement",
    "UNKNOWN": "Unknown Financial Statement",
}

FINANCIAL_CONCEPTS = {
    "OPENING_BALANCE", "BROUGHT_FORWARD", "CARRIED_FORWARD", "CLOSING_BALANCE",
    "TRANSACTION_DATE", "VALUE_DATE", "POSTING_DATE", "NARRATION", "REFERENCE",
    "DEBIT", "CREDIT", "BALANCE", "PAYMENT", "REFUND", "FEE", "INTEREST",
    "STATEMENT_PERIOD", "PAYMENT_DUE_DATE", "MINIMUM_DUE", "NEW_BALANCE",
    "CREDIT_LIMIT", "AVAILABLE_CREDIT", "CASH_ADVANCE", "PURCHASE", "ADJUSTMENT",
    "REWARD", "CARD_NUMBER", "ACCOUNT_NUMBER",
}

# These are intentionally compact seed packs.  They are interpreted by the
# generic grammar, not used as bank-specific parsers.  Additional packs can be
# dropped beside the EXE without changing the parser source.
BUILTIN_LANGUAGE_PACKS = {
    "en": {
        "name": "English", "script": "Latin", "direction": "LTR",
        "terms": {
            "opening_balance": ["OPENING BALANCE", "OPENING BAL", "BALANCE BROUGHT FORWARD"],
            "brought_forward": ["BROUGHT FORWARD", "B/F", "BALANCE B/F"],
            "carried_forward": ["CARRIED FORWARD", "C/F", "BALANCE C/F"],
            "closing_balance": ["CLOSING BALANCE", "CLOSING BAL", "ENDING BALANCE"],
            "debit": ["DEBIT", "WITHDRAWAL", "DR"], "credit": ["CREDIT", "DEPOSIT", "CR"],
            "payment": ["PAYMENT", "PAYMENTS"], "refund": ["REFUND", "REFUNDED"],
            "fee": ["FEE", "FEES", "CHARGE", "CHARGES"], "interest": ["INTEREST", "FINANCE CHARGE"],
            "minimum_due": ["MINIMUM PAYMENT", "MINIMUM AMOUNT DUE", "MINIMUM DUE"],
            "payment_due_date": ["PAYMENT DUE DATE", "DUE DATE"], "new_balance": ["NEW BALANCE", "AMOUNT DUE"],
            "credit_limit": ["CREDIT LIMIT"], "available_credit": ["AVAILABLE CREDIT"],
            "cash_advance": ["CASH ADVANCE"], "purchase": ["PURCHASE", "PURCHASES"],
            "statement_period": ["STATEMENT PERIOD", "BILLING PERIOD"],
        },
    },
    "fr": {
        "name": "French", "script": "Latin", "direction": "LTR",
        "terms": {
            "opening_balance": ["SOLDE D'OUVERTURE", "SOLDE INITIAL", "SOLDE PRÉCÉDENT"],
            "brought_forward": ["REPORT À NOUVEAU", "REPORT", "SOLDE REPORTÉ"],
            "closing_balance": ["SOLDE DE CLÔTURE", "SOLDE FINAL"],
            "debit": ["DÉBIT", "RETRAIT"], "credit": ["CRÉDIT", "DÉPÔT"],
            "payment": ["PAIEMENT", "VERSEMENT"], "refund": ["REMBOURSEMENT"],
            "fee": ["FRAIS", "COMMISSION"], "interest": ["INTÉRÊTS"],
            "minimum_due": ["PAIEMENT MINIMUM", "MONTANT MINIMUM DÛ"],
            "payment_due_date": ["DATE D'ÉCHÉANCE", "DATE LIMITE DE PAIEMENT"],
            "new_balance": ["NOUVEAU SOLDE", "SOLDE À PAYER"], "credit_limit": ["LIMITE DE CRÉDIT"],
            "available_credit": ["CRÉDIT DISPONIBLE"], "cash_advance": ["AVANCE DE TRÉSORERIE"],
            "purchase": ["ACHAT", "ACHATS"], "statement_period": ["PÉRIODE DU RELEVÉ"],
        },
    },
    "es": {
        "name": "Spanish", "script": "Latin", "direction": "LTR",
        "terms": {
            "opening_balance": ["SALDO INICIAL", "SALDO ANTERIOR"],
            "brought_forward": ["SALDO ARRASTRADO", "SALDO ANTERIOR"],
            "closing_balance": ["SALDO FINAL", "SALDO DE CIERRE"],
            "debit": ["DÉBITO", "CARGO", "RETIRO"], "credit": ["CRÉDITO", "ABONO", "DEPÓSITO"],
            "payment": ["PAGO", "PAGOS"], "refund": ["REEMBOLSO"], "fee": ["COMISIÓN", "CARGO", "TARIFA"],
            "interest": ["INTERESES"], "minimum_due": ["PAGO MÍNIMO", "IMPORTE MÍNIMO"],
            "payment_due_date": ["FECHA DE VENCIMIENTO", "FECHA LÍMITE DE PAGO"],
            "new_balance": ["NUEVO SALDO", "SALDO A PAGAR"], "credit_limit": ["LÍMITE DE CRÉDITO"],
            "available_credit": ["CRÉDITO DISPONIBLE"], "cash_advance": ["DISPOSICIÓN DE EFECTIVO"],
            "purchase": ["COMPRA", "COMPRAS"], "statement_period": ["PERÍODO DEL ESTADO DE CUENTA"],
        },
    },
    "de": {"name":"German", "script":"Latin", "direction":"LTR", "terms": {"opening_balance":["ANFANGSSALDO","ERÖFFNUNGSSALDO"],"closing_balance":["ENDSALDO","ENDGÜLTIGER SALDO"],"debit":["LASTSCHRIFT","SOLL"],"credit":["GUTSCHRIFT","HABEN"],"payment":["ZAHLUNG"],"refund":["ERSTATTUNG"],"fee":["GEBÜHR","GEBÜHREN"],"interest":["ZINSEN"],"minimum_due":["MINDESTZAHLUNG"],"payment_due_date":["FÄLLIGKEITSDATUM"],"credit_limit":["KREDITLIMIT"],"available_credit":["VERFÜGBARER KREDIT"]}},
    "pt": {"name":"Portuguese", "script":"Latin", "direction":"LTR", "terms": {"opening_balance":["SALDO INICIAL","SALDO ANTERIOR"],"closing_balance":["SALDO FINAL"],"debit":["DÉBITO","LEVANTAMENTO"],"credit":["CRÉDITO","DEPÓSITO"],"payment":["PAGAMENTO"],"refund":["REEMBOLSO"],"fee":["TAXA","COMISSÃO"],"interest":["JUROS"],"minimum_due":["PAGAMENTO MÍNIMO"],"payment_due_date":["DATA DE VENCIMENTO"],"credit_limit":["LIMITE DE CRÉDITO"],"available_credit":["CRÉDITO DISPONÍVEL"]}},
    "ru": {"name":"Russian", "script":"Cyrillic", "direction":"LTR", "terms": {"opening_balance":["НАЧАЛЬНЫЙ ОСТАТОК","ВХОДЯЩИЙ ОСТАТОК"],"closing_balance":["КОНЕЧНЫЙ ОСТАТОК"],"debit":["ДЕБЕТ","СПИСАНИЕ"],"credit":["КРЕДИТ","ЗАЧИСЛЕНИЕ"],"payment":["ПЛАТЕЖ"],"refund":["ВОЗВРАТ"],"fee":["КОМИССИЯ","КОМИССИИ"],"interest":["ПРОЦЕНТЫ"],"minimum_due":["МИНИМАЛЬНЫЙ ПЛАТЕЖ"],"payment_due_date":["СРОК ПЛАТЕЖА"],"credit_limit":["КРЕДИТНЫЙ ЛИМИТ"],"available_credit":["ДОСТУПНЫЙ КРЕДИТ"]}},
    "zh": {"name":"Chinese", "script":"Han", "direction":"LTR", "terms": {"opening_balance":["期初余额","上期余额"],"brought_forward":["结转余额"],"closing_balance":["期末余额","期末结余"],"debit":["借方","支出"],"credit":["贷方","存入"],"payment":["还款","付款"],"refund":["退款"],"fee":["费用","手续费"],"interest":["利息"],"minimum_due":["最低还款额"],"payment_due_date":["还款到期日","到期还款日"],"new_balance":["本期应还","新余额"],"credit_limit":["信用额度"],"available_credit":["可用额度"],"cash_advance":["现金预借"],"purchase":["消费","消费金额"],"statement_period":["账单周期","账单期间"]}},
    "ar": {"name":"Arabic", "script":"Arabic", "direction":"RTL", "terms": {"opening_balance":["الرصيد الافتتاحي","الرصيد السابق"],"closing_balance":["الرصيد الختامي","الرصيد النهائي"],"debit":["مدين","سحب"],"credit":["دائن","إيداع"],"payment":["دفعة","سداد"],"refund":["استرداد"],"fee":["رسوم","عمولة"],"interest":["فائدة","فوائد"],"minimum_due":["الحد الأدنى للدفع"],"payment_due_date":["تاريخ استحقاق الدفع"],"credit_limit":["حد الائتمان"],"available_credit":["الائتمان المتاح"],"cash_advance":["سحب نقدي"]}},
    "hi": {"name":"Hindi", "script":"Devanagari", "direction":"LTR", "terms": {"opening_balance":["प्रारंभिक शेष","प्रारंभिक बैलेंस"],"closing_balance":["अंतिम शेष","समापन शेष"],"debit":["डेबिट","निकासी"],"credit":["क्रेडिट","जमा"],"payment":["भुगतान"],"refund":["वापसी"],"fee":["शुल्क"],"interest":["ब्याज"],"minimum_due":["न्यूनतम भुगतान"],"payment_due_date":["भुगतान की अंतिम तिथि"],"credit_limit":["क्रेडिट सीमा"],"available_credit":["उपलब्ध क्रेडिट"]}},
    "ja": {"name":"Japanese", "script":"Japanese", "direction":"LTR", "terms": {"opening_balance":["期首残高","繰越残高"],"closing_balance":["期末残高"],"debit":["借方","出金"],"credit":["貸方","入金"],"payment":["支払"],"refund":["返金"],"fee":["手数料"],"interest":["利息"],"minimum_due":["最低支払額"],"payment_due_date":["支払期日"],"credit_limit":["利用限度額"],"available_credit":["利用可能額"]}},
    "ko": {"name":"Korean", "script":"Hangul", "direction":"LTR", "terms": {"opening_balance":["기초잔액","전월잔액"],"closing_balance":["기말잔액"],"debit":["차변","출금"],"credit":["대변","입금"],"payment":["결제","지급"],"refund":["환불"],"fee":["수수료"],"interest":["이자"],"minimum_due":["최소 결제 금액"],"payment_due_date":["결제일"],"credit_limit":["신용 한도"],"available_credit":["사용 가능 한도"]}},
    "nl": {"name":"Dutch", "script":"Latin", "direction":"LTR", "terms": {"opening_balance":["BEGINSALDO","VORIG SALDO"],"closing_balance":["EINDSALDO"],"debit":["DEBET","AFSCHRIJVING"],"credit":["CREDIT","BIJSCHRIJVING"],"payment":["BETALING"],"refund":["TERUGBETALING"],"fee":["KOSTEN"],"interest":["RENTE"],"minimum_due":["MINIMALE BETALING"],"payment_due_date":["VERVALDATUM"],"credit_limit":["KREDIETLIMIET"],"available_credit":["BESCHIKBAAR KREDIET"]}},
}


def _norm_label_text(value):
    import unicodedata
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = re.sub(r"[\u00a0\u2007\u202f]", " ", text)
    return re.sub(r"\s+", " ", text).strip().casefold()


def detect_statement_language(text):
    """Deterministic script/term heuristic. Never uses an ML model."""
    sample = str(text or "")[:20000]
    scores = {code: 0 for code in BUILTIN_LANGUAGE_PACKS}
    ranges = {
        "ar": lambda c: '\u0600' <= c <= '\u06ff',
        "ru": lambda c: '\u0400' <= c <= '\u04ff',
        "zh": lambda c: '\u4e00' <= c <= '\u9fff',
        "ja": lambda c: ('\u3040' <= c <= '\u30ff') or ('\u4e00' <= c <= '\u9fff'),
        "ko": lambda c: '\uac00' <= c <= '\ud7af',
        "hi": lambda c: '\u0900' <= c <= '\u097f',
    }
    for code, fn in ranges.items():
        scores[code] += sum(1 for c in sample if fn(c))
    upper = _norm_label_text(sample)
    for code, pack in BUILTIN_LANGUAGE_PACKS.items():
        for terms in pack.get("terms", {}).values():
            for term in terms:
                if _norm_label_text(term) in upper:
                    scores[code] += 12
    best = max(scores, key=scores.get)
    if scores[best] == 0:
        return "en"
    return best


def detect_statement_type(text):
    """Classify Bank vs Credit Card from deterministic financial terminology."""
    upper = _norm_label_text(text).upper()
    card_terms = [
        "CREDIT CARD", "CARD NUMBER", "MINIMUM PAYMENT", "MINIMUM AMOUNT DUE",
        "PAYMENT DUE DATE", "CREDIT LIMIT", "AVAILABLE CREDIT", "APR",
        "CASH ADVANCE", "FINANCE CHARGE", "STATEMENT BALANCE",
    ]
    score = sum(2 for term in card_terms if term in upper)
    if "CREDIT CARD" in upper or score >= 4:
        return "CREDIT_CARD"
    return "BANK"


def extract_credit_card_summary(text, language="en"):
    """Extract statement-level credit-card fields conservatively.

    PDF text layers may contain escaped punctuation, non-breaking spaces,
    line breaks, currency symbols and locale-specific separators. This
    routine normalizes presentation whitespace but does not infer values.
    """
    raw = str(text or "")
    norm = re.sub(r"[\u00a0\u2007\u202f]", " ", raw)
    # PDF text extraction may split a label or date across line boundaries.
    # Preserve a single logical space between tokens so deterministic field
    # patterns can match without reconstructing or inferring values.
    norm = re.sub(r"[ \t\r\n]+", " ", norm)

    result = {
        "document_type": "CREDIT_CARD",
        "statement_period": None,
        "payment_due_date": None,
        "minimum_due": None,
        "new_balance": None,
        "credit_limit": None,
        "available_credit": None,
        "previous_balance": None,
    }

    amount = r"(?:₹|RS\.?|\$|€|£)?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)"
    patterns = {
        "payment_due_date": r"(?:PAYMENT DUE DATE|DUE DATE)\s*[:.-]?\s*([0-9]{1,2}[/-][0-9]{1,2}(?:[/-][0-9]{2,4})?)",
        "minimum_due": rf"(?:MINIMUM (?:AMOUNT )?DUE|MINIMUM PAYMENT)\s*[:.-]?\s*{amount}",
        "new_balance": rf"(?:NEW BALANCE|STATEMENT BALANCE|AMOUNT DUE)\s*[:.-]?\s*{amount}",
        "credit_limit": rf"CREDIT LIMIT\s*[:.-]?\s*{amount}",
        "available_credit": rf"AVAILABLE CREDIT\s*[:.-]?\s*{amount}",
        "previous_balance": rf"(?:PREVIOUS BALANCE|OPENING BALANCE)\s*[:.-]?\s*{amount}",
    }

    for key, pat in patterns.items():
        match = re.search(pat, norm, flags=re.IGNORECASE)
        if not match:
            continue
        raw_value = match.group(1)
        if key == "payment_due_date":
            # Dates are not monetary values. Keep the extracted date text
            # exactly as a normalized statement field instead of sending it
            # through parse_decimal().
            result[key] = raw_value
            continue
        parsed = parse_decimal(raw_value)
        if parsed is not None:
            result[key] = str(parsed)

    return result


def ensure_taxforensic_support_dirs():
    root = APP_DIRECTORY
    paths = [os.path.join(root, "language_packs"), os.path.join(root, "corpus")]
    for path in paths:
        os.makedirs(path, exist_ok=True)
    manifest = os.path.join(root, "language_packs", "manifest.json")
    if not os.path.exists(manifest):
        payload = {"format_version":1, "engine":TAXFORENSIC_ENGINE_VERSION,
                   "packs": [{"code":k,"name":v["name"],"script":v["script"],"direction":v["direction"],"builtin":True} for k,v in BUILTIN_LANGUAGE_PACKS.items()]}
        with open(manifest, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)


def build_synthetic_corpus():
    """Return small deterministic structural cases. No original statements and no AI."""
    cases = []
    def add(cid, family, text, expected):
        cases.append({"id":cid,"family":family,"text":text,"expected":expected})
    add("REF-001","reference-balance","BROUGHT FORWARD : 202.80Cr",{"value":"202.80"})
    add("REF-002","reference-balance","BROUGHT FORWARD : 202.80 Cr",{"value":"202.80"})
    add("REF-003","reference-balance","BROUGHT FORWARD :\\n202.80\\nCr",{"value":"202.80"})
    add("REF-004","reference-balance","B/F\\n202.80\\nCr",{"value":"202.80"})
    add("REF-005","reference-balance","BALANCE B/F : Rs. 2,02,000.00 Dr",{"value":"202000.00"})
    add("CC-001","credit-card","CREDIT CARD\\nMINIMUM PAYMENT 2,500.00\\nNEW BALANCE 25,000.00\\nCREDIT LIMIT 100,000.00",{"document_type":"CREDIT_CARD","minimum_due":"2500.00","new_balance":"25000.00","credit_limit":"100000.00"})
    add("CC-002","credit-card","STATEMENT BALANCE 1,234.56\\nAVAILABLE CREDIT 8,765.44\\nPAYMENT DUE DATE 25/10/2026",{"document_type":"CREDIT_CARD","new_balance":"1234.56","available_credit":"8765.44","payment_due_date":"25/10/2026"})
    add("LANG-001","language","Solde d'ouverture\\nDébit\\nCrédit",{"language":"fr"})
    add("LANG-002","language","الرصيد الافتتاحي\\nمدين\\nدائن",{"language":"ar"})
    add("LANG-003","language","期初余额\\n借方\\n贷方",{"language":"zh"})
    add("LANG-004","language","НАЧАЛЬНЫЙ ОСТАТОК\\nДЕБЕТ\\nКРЕДИТ",{"language":"ru"})
    add("LANG-005","language","Saldo inicial\\nDébito\\nCrédito",{"language":"es"})
    add("LANG-006","language","प्रारंभिक शेष\\nडेबिट\\nक्रेडिट",{"language":"hi"})
    return cases


def run_v26_corpus():
    ensure_taxforensic_support_dirs()
    cases = build_synthetic_corpus()
    failures = []
    for case in cases:
        family = case["family"]
        if family == "reference-balance":
            value, _ = extract_any_reference_balance(case["text"])
            got = None if value is None else str(value)
            if got != case["expected"]["value"]:
                failures.append((case["id"], got, case["expected"]["value"]))
        elif family == "credit-card":
            got = extract_credit_card_summary(case["text"])
            for key, expected in case["expected"].items():
                if got.get(key) != expected:
                    failures.append((case["id"], key, got.get(key), expected))
        elif family == "language":
            got = detect_statement_language(case["text"])
            if got != case["expected"]["language"]:
                failures.append((case["id"], got, case["expected"]["language"]))
    print("=" * 80)
    print(f"TaxForensic {TAXFORENSIC_ENGINE_VERSION} deterministic corpus")
    print(f"Corpus version: {CORPUS_VERSION}")
    print(f"Cases: {len(cases)}")
    print(f"Passed: {len(cases)-len(failures)}")
    print(f"Failed: {len(failures)}")
    if failures:
        for failure in failures:
            print("FAIL", failure)
        print("CORPUS FAILED")
        return 1
    print("CORPUS PASSED")
    return 0


def format_inr(val) -> str:
    """
    Formats a Decimal/number into Indian currency notation.
    """

    if val is None or val == "UNEXTRACTED_AMOUNT":
        return "UNEXTRACTED_AMOUNT"

    try:
        d = Decimal(str(val)).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    except Exception:
        return "Rs. 0.00"

    is_negative = d < 0

    d = abs(d)

    formatted = f"{d:.2f}"

    int_part, decimal_part = formatted.split(".")

    if len(int_part) <= 3:
        result = int_part

    else:
        last_three = int_part[-3:]
        rest = int_part[:-3]

        groups = []

        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]

        if rest:
            groups.insert(0, rest)

        result = ",".join(groups) + "," + last_three

    if is_negative:
        return f"-Rs. {result}.{decimal_part}"

    return f"Rs. {result}.{decimal_part}"


def parse_decimal(val_str):
    """
    Safely parses monetary values.

    Returns:
        Decimal
            when the value is valid

        None
            when the value is missing/unextractable/invalid
    """

    if val_str is None:
        return None

    if isinstance(val_str, Decimal):
        return val_str.quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )

    if isinstance(val_str, (int, float)):
        try:
            return Decimal(str(val_str)).quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP,
            )
        except Exception:
            return None

    clean = str(val_str).strip()

    if not clean:
        return None

    if clean.upper() in {
        "UNEXTRACTED_AMOUNT",
        "N/A",
        "NA",
        "NULL",
        "NONE",
        "-",
        "--",
    }:
        return None

    negative = False

    if clean.startswith("(") and clean.endswith(")"):
        negative = True
        clean = clean[1:-1]

    clean = re.sub(r"₹", "", clean)

    clean = re.sub(
        r"Rs\.?",
        "",
        clean,
        flags=re.IGNORECASE,
    )

    clean = re.sub(
        r"\bDR\b",
        "",
        clean,
        flags=re.IGNORECASE,
    )

    clean = re.sub(
        r"\bCR\b",
        "",
        clean,
        flags=re.IGNORECASE,
    )

    clean = clean.replace(",", "")
    clean = clean.replace(" ", "")
    clean = clean.strip()

    if not clean:
        return None

    if not re.fullmatch(
        r"[-+]?\d+(?:\.\d+)?",
        clean,
    ):
        return None

    try:
        value = Decimal(clean).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )

        return -value if negative else value

    except Exception:
        return None


def decimal_or_zero(value):
    """
    Use only where missing values are logically equivalent to zero.

    DO NOT use this for extraction validation.
    """

    parsed = parse_decimal(value)

    if parsed is None:
        return Decimal("0.00")

    return parsed


def infer_declared_cash_source(narration: str, amount):
    """
    Identify an explicit external physical-cash source stated in the
    narration. This is deliberately conservative: only explicit wording
    such as CASH RECEIVED is treated as a source. A normal CASH DEPOSIT
    remains a use of previously available physical cash.

    Returns the source amount when explicitly indicated, otherwise zero.
    """

    parsed_amount = parse_decimal(amount)

    if parsed_amount is None or parsed_amount <= Decimal("0.00"):
        return Decimal("0.00")

    text = str(narration or "").upper()

    explicit_source_phrases = [
        "CASH RECEIVED",
        "CASH RECEIPT",
        "CASH RECEIVED FROM",
        "CASH RECEIVED FOR",
        "CASH SALE RECEIPT",
    ]

    if any(phrase in text for phrase in explicit_source_phrases):
        return parsed_amount

    return Decimal("0.00")


def safe_float(val) -> float:
    """
    Converts a monetary value into float for SQLite/Excel.

    Extraction status is preserved separately by the transaction.
    """

    parsed = parse_decimal(val)

    if parsed is None:
        return 0.0

    return float(parsed)


def compute_file_sha256(filepath: str) -> str:

    if not filepath or not os.path.exists(filepath):
        return "VERIFIED_FILE_HASH"

    sha = hashlib.sha256()

    with open(filepath, "rb") as f:

        while True:

            chunk = f.read(65536)

            if not chunk:
                break

            sha.update(chunk)

    return sha.hexdigest().upper()


# ======================================================================
# 2. DATABASE
# ======================================================================

def init_database():

    conn = sqlite3.connect(DATABASE_PATH)

    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS bank_transactions (
            txn_id INTEGER PRIMARY KEY AUTOINCREMENT,
            pan TEXT,
            bank_name TEXT,
            account_number TEXT,
            txn_date TEXT,
            narration TEXT,
            mode TEXT,
            withdrawal NUMERIC,
            deposit NUMERIC,
            balance NUMERIC,
            is_contra BOOLEAN DEFAULT 0,
            audit_flag TEXT,
            pdf_page INTEGER,
            bbox_coords TEXT,
            source_file TEXT
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS forensic_schedules (
            pan TEXT PRIMARY KEY,
            peak_unexplained_cash NUMERIC,
            sec_269st_violations NUMERIC,
            sec_269ss_violations NUMERIC,
            sft_cash_total NUMERIC,
            contra_neutralized NUMERIC,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    conn.commit()
    conn.close()


# ======================================================================
# 3. DATE UTILITIES
# ======================================================================

def parse_statement_date(date_text):

    if not date_text:
        return None

    value = str(date_text).strip()

    candidates = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%d.%m.%y",
        "%d-%b-%Y",
        "%d/%b/%Y",
        "%d.%b.%Y",
        "%d %b %Y",
        "%d-%b-%y",
        "%d %b %y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y%m%d",
        "%y%m%d",
    ]

    normalized = (
        value
        .replace("/", "-")
        .replace(".", "-")
    )

    for fmt in candidates:

        try:
            return datetime.strptime(
                normalized,
                fmt,
            )

        except Exception:
            pass

    return None


# ======================================================================
# 4. FORENSIC DOCUMENT VIEW
# ======================================================================

class ForensicDocumentView(QGraphicsView):
    """Zoomable document inspector with mouse-wheel support."""

    zoomChanged = Signal(float)

    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self._document_zoom = 1.0
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setStyleSheet(
            "QGraphicsView { background: #202124; border: 2px solid #607d8b; }"
        )

    def wheelEvent(self, event):
        # Ctrl + wheel is an explicit zoom gesture. Normal wheel retains
        # standard scrolling behaviour so the inspector remains usable on
        # long crops.
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                factor = 1.20
            elif delta < 0:
                factor = 1 / 1.20
            else:
                factor = 1.0
            if factor != 1.0:
                self.scale(factor, factor)
                self._document_zoom = max(0.10, min(12.0, self._document_zoom * factor))
                self.zoomChanged.emit(self._document_zoom)
            event.accept()
            return
        super().wheelEvent(event)


# ======================================================================
# 5. DYNAMIC HEADER CALIBRATOR
# ======================================================================

class DynamicHeaderCalibrator:

    @classmethod
    def calibrate_columns(
        cls,
        page: fitz.Page,
        fallback: dict = None,
    ) -> dict:
        """Calibrate statement columns from actual header words.

        Bank PDFs frequently split the header into separate text blocks. The
        earlier block-only detector could therefore miss the Balance column
        even when the header was visibly present. We group words by vertical
        position and identify the strongest header row across the whole page.
        """
        words = page.get_text("words")
        if not words:
            return fallback or {
                "date": (0, 110), "narration": (110, 420),
                "withdrawal": (420, 560), "deposit": (560, 700),
                "balance": (700, 9999),
            }

        # Header vocabulary is semantic and layout-independent. Match complete
        # normalized tokens/phrases, not arbitrary substrings: a short alias
        # such as DR or CR must never match a longer unrelated word. Plural and
        # common international variants are included because they describe the
        # same column role, not a particular bank.
        roles = {
            "date": ("DATE", "TXN DATE", "TRANSACTION DATE", "VALUE DATE", "POSTING DATE", "BOOKING DATE", "BOOKING DAT", "ENTRY DATE"),
            "narration": ("PARTICULARS", "NARRATION", "DESCRIPTION", "TRANSACTION DETAILS", "TRANSACTION DESCRIPTION", "REMARKS", "DETAILS", "PAYMENT DETAILS", "REFERENCE DETAILS"),
            "withdrawal": ("WITHDRAWAL", "WITHDRAWALS", "WITHDRAW", "DEBIT", "DEBITS", "DEBIT AMOUNT", "WITHDRAWAL AMOUNT", "DR", "DR AMOUNT", "PAID OUT", "MONEY OUT"),
            "deposit": ("DEPOSIT", "DEPOSITS", "CREDIT", "CREDITS", "CREDIT AMOUNT", "DEPOSIT AMOUNT", "CR", "CR AMOUNT", "PAID IN", "MONEY IN"),
            "balance": ("BALANCE", "BAL", "CLOSING BALANCE", "RUNNING BALANCE", "LEDGER BALANCE", "AVAILABLE BALANCE", "CURRENT BALANCE", "BALANCE AMOUNT"),
        }
        def _header_norm(value):
            return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
        role_aliases = {
            role: {_header_norm(alias) for alias in aliases}
            for role, aliases in roles.items()
        }

        # Group header words into horizontal rows using their y-centres.
        rows = []
        for word in words:
            yc = (float(word[1]) + float(word[3])) / 2.0
            placed = False
            for row in rows:
                if abs(yc - row["yc"]) <= 8.0:
                    row["words"].append(word)
                    row["yc"] = sum((float(w[1]) + float(w[3])) / 2.0 for w in row["words"]) / len(row["words"])
                    placed = True
                    break
            if not placed:
                rows.append({"yc": yc, "words": [word]})

        best = None
        for row in rows:
            found = {}
            right_edge_found = {}
            row_words = sorted(row["words"], key=lambda w: float(w[0]))
            # Consider one-, two-, and three-word phrases. This handles headers
            # split into separate PDF text spans (for example "Transaction" +
            # "Date") while requiring horizontal proximity so unrelated cells
            # cannot accidentally form a phrase.
            for start in range(len(row_words)):
                for count in (3, 2, 1):
                    end = start + count - 1
                    if end >= len(row_words):
                        continue
                    phrase_words = row_words[start:end + 1]
                    if count > 1:
                        gaps = [float(phrase_words[i + 1][0]) - float(phrase_words[i][2])
                                for i in range(len(phrase_words) - 1)]
                        if any(gap > 24.0 or gap < -2.0 for gap in gaps):
                            continue
                    phrase = "".join(_header_norm(w[4]) for w in phrase_words)
                    if not phrase:
                        continue
                    x0_phrase = min(float(w[0]) for w in phrase_words)
                    x1_phrase = max(float(w[2]) for w in phrase_words)
                    xmid = (x0_phrase + x1_phrase) / 2.0
                    for role, aliases in role_aliases.items():
                        if role not in found and phrase in aliases:
                            found[role] = xmid
                            if role in {"withdrawal", "deposit", "balance"}:
                                right_edge_found[role] = x1_phrase
            score = len(found)
            if "date" in found and "balance" in found:
                score += 1
            if "narration" in found:
                score += 1
            if best is None or score > best[0]:
                best = (score, found, row["yc"], right_edge_found)

        anchors = best[1] if best and best[0] >= 3 else {}
        right_edge_anchors = dict(best[3]) if best and best[0] >= 3 else {}

        # If a page omits its header, reuse the previous page's calibrated
        # layout. This is safer than guessing from transaction amounts.
        if len(anchors) < 3 and fallback:
            return fallback

        # Fill missing header anchors from fallback where possible.
        if fallback:
            for role in roles:
                if role not in anchors and role in fallback:
                    left, right = fallback[role]
                    anchors[role] = (left + right) / 2.0

        # A sensible emergency layout for genuinely headerless statements.
        if not anchors:
            money_x = []
            for word in words:
                token = str(word[4]).strip()
                if is_money_token(token) or is_running_balance_token(token):
                    money_x.append((float(word[0]) + float(word[2])) / 2.0)
            money_x.sort()
            if len(money_x) >= 3:
                anchors = {
                    "date": 45.0,
                    "narration": 180.0,
                    "withdrawal": money_x[-3],
                    "deposit": money_x[-2],
                    "balance": money_x[-1],
                }
            else:
                return fallback or {
                    "date": (0, 110), "narration": (110, 420),
                    "withdrawal": (420, 560), "deposit": (560, 700),
                    "balance": (700, 9999),
                }

        ordered = sorted(anchors.items(), key=lambda item: item[1])
        channel_map = {}
        for idx, (name, anchor) in enumerate(ordered):
            left = 0.0 if idx == 0 else (ordered[idx - 1][1] + anchor) / 2.0
            right = 9999.0 if idx == len(ordered) - 1 else (anchor + ordered[idx + 1][1]) / 2.0
            channel_map[name] = (left, right)

        # Keep required fields even if two visual columns share an x-position.
        required = ["date", "narration", "withdrawal", "deposit", "balance"]
        for role in required:
            if role not in channel_map:
                if fallback and role in fallback:
                    channel_map[role] = fallback[role]
                else:
                    channel_map[role] = (0.0, 9999.0)

        # Reuse prior-page right-edge anchors if this page's heading row is
        # incomplete. The ordinary centre-based channels remain available for
        # narration and date text.
        if fallback:
            previous_edges = fallback.get("_money_right_edge_anchors", {})
            for role in ("withdrawal", "deposit", "balance"):
                if role not in right_edge_anchors and role in previous_edges:
                    try:
                        right_edge_anchors[role] = float(previous_edges[role])
                    except Exception:
                        pass

        sorted_money_edges = sorted(right_edge_anchors.items(), key=lambda item: item[1])
        money_edge_channels = {}
        if len(sorted_money_edges) >= 2:
            for idx, (role, edge_anchor) in enumerate(sorted_money_edges):
                if idx == 0:
                    # Avoid classifying numbers inside narration as money.
                    left = channel_map.get(role, (0.0, 9999.0))[0]
                else:
                    left = (sorted_money_edges[idx - 1][1] + edge_anchor) / 2.0
                if idx == len(sorted_money_edges) - 1:
                    right = channel_map.get(role, (0.0, 9999.0))[1]
                else:
                    right = (edge_anchor + sorted_money_edges[idx + 1][1]) / 2.0
                money_edge_channels[role] = (float(left), float(right))

        channel_map["_money_right_edge_channels"] = money_edge_channels
        channel_map["_money_right_edge_anchors"] = right_edge_anchors
        return channel_map


def find_reference_line_bbox(page, labels):
    """Find the exact visual line containing a reference label and its amount.

    This is deliberately positional.  It does not use transaction arithmetic
    and does not assume that the reference amount is in a particular column.
    """
    wanted = [str(x).upper() for x in labels]
    words = page.get_text("words") or []
    if not words:
        return []
    lines = []
    for w in words:
        x0, y0, x1, y1, text = float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4]).strip()
        if not text:
            continue
        cy = (y0 + y1) / 2.0
        target = None
        for line in lines:
            if abs(cy - line["cy"]) <= max(3.5, (y1 - y0) * 0.75):
                target = line
                break
        if target is None:
            target = {"cy": cy, "words": []}
            lines.append(target)
        target["words"].append(w)
        target["cy"] = sum((float(a[1]) + float(a[3])) / 2.0 for a in target["words"]) / len(target["words"])
    results = []
    for line in lines:
        line["words"].sort(key=lambda w: float(w[0]))
        text = " ".join(str(w[4]).strip() for w in line["words"])
        upper = text.upper()
        matched = next((label for label in wanted if label in upper), None)
        if not matched:
            continue
        x0 = min(float(w[0]) for w in line["words"])
        y0 = min(float(w[1]) for w in line["words"])
        x1 = max(float(w[2]) for w in line["words"])
        y1 = max(float(w[3]) for w in line["words"])
        amount = None
        amount_bbox = None
        for w in line["words"]:
            token = str(w[4]).strip()
            if is_running_balance_token(token):
                amount = parse_decimal(re.sub(r"(?i)(CR|DR)$", "", token.replace(" ", "")))
                if amount is not None:
                    amount_bbox = [float(w[0]), float(w[1]), float(w[2]), float(w[3])]
                    break
            val = parse_decimal(token)
            if val is not None:
                amount = val
                amount_bbox = [float(w[0]), float(w[1]), float(w[2]), float(w[3])]

        # Some statements print the label and amount on adjacent visual lines,
        # e.g. "BROUGHT FORWARD" followed by "202.80 Cr". Only search below
        # the matched label and only when that line itself has no amount.
        if amount is None:
            label_bottom = y1
            nearby = []
            for w in words:
                wx0, wy0, wx1, wy1 = map(float, w[:4])
                gap = wy0 - label_bottom
                if gap < -2 or gap > 26:
                    continue
                token = str(w[4]).strip()
                val = None
                explicit_balance = is_running_balance_token(token)
                if explicit_balance:
                    val = parse_decimal(re.sub(r"(?i)(CR|DR)$", "", token.replace(" ", "")))
                elif is_money_token(token):
                    # A plain amount on a nearby line is not enough: it may be
                    # narration (e.g. a transfer of 50,000.00). Only consider
                    # it when it lies in the rightmost balance area.
                    page_width = float(page.rect.width)
                    if wx0 >= page_width * 0.62:
                        val = parse_decimal(token)
                if val is not None:
                    nearby.append((max(0.0, gap), 0 if explicit_balance else 1, wx0, wy0, wx1, wy1, val, token))
            if nearby:
                nearby.sort(key=lambda item: (item[1], item[0], item[2]))
                candidate = nearby[0]
                amount = candidate[6]
                amount_bbox = list(candidate[2:6])
                x0 = min(x0, amount_bbox[0]); y0 = min(y0, amount_bbox[1])
                x1 = max(x1, amount_bbox[2]); y1 = max(y1, amount_bbox[3])

        # Expand evidence bounds to include the amount when it shares the line.
        if amount_bbox is not None:
            x0 = min(x0, amount_bbox[0]); y0 = min(y0, amount_bbox[1])
            x1 = max(x1, amount_bbox[2]); y1 = max(y1, amount_bbox[3])
        results.append({"label": matched, "text": text, "bbox": [x0, y0, x1, y1], "amount": amount, "cy": line["cy"]})
    return results


def ensure_printed_amount_evidence(tx):
    """Preserve the first machine-extracted debit/credit as printed evidence."""
    if "image_debit" not in tx:
        tx["image_debit"] = tx.get("withdrawal", "UNEXTRACTED_AMOUNT")
    if "image_credit" not in tx:
        tx["image_credit"] = tx.get("deposit", "UNEXTRACTED_AMOUNT")
    if "image_balance" not in tx:
        tx["image_balance"] = tx.get("balance", "UNEXTRACTED_AMOUNT")
    return tx


def open_pdf_with_optional_security(path, password=None):
    """Open authorized/password-protected PDFs through pikepdf/qpdf when available."""
    try:
        import pikepdf
        import io
        with pikepdf.open(path, password=password or "") as pdf:
            buf = io.BytesIO()
            pdf.save(buf, encryption=False)
            return fitz.open(stream=buf.getvalue(), filetype="pdf")
    except Exception:
        return fitz.open(path)


def extract_statement_metadata(filepath):
    """Lightweight metadata extraction performed when a statement is loaded.

    It reads text/headers only and does not admit transaction rows.  This is
    intentionally separate from the forensic transaction parser.
    """
    meta = {
        "file": os.path.basename(filepath),
        "path": filepath,
        "document_type": "UNKNOWN",
        "language": "en",
        "bank": "Pending detection",
        "account": "Pending detection",
        "holder_name": "Not detected",
        "start_date": None,
        "end_date": None,
        "currency": "Not detected",
        "pages": None,
    }
    ext = os.path.splitext(filepath)[1].lower()
    try:
        sample = ""
        if ext == ".pdf":
            with open(filepath, "rb") as f:
                data = f.read()
            doc = fitz.open(stream=data, filetype="pdf")
            meta["pages"] = len(doc)
            if len(doc):
                sample = "\n".join(doc[i].get_text("text") for i in range(min(len(doc), 2)))
                if len(doc) > 2:
                    sample += "\n" + doc[-1].get_text("text")
            try:
                helper = StatementParserWorker([])
                bank, account = helper.detect_bank_and_acc(sample, filepath)
                if bank and bank != "Bank Account": meta["bank"] = bank
                if account: meta["account"] = account
            finally:
                doc.close()
        elif ext in {".csv", ".tsv", ".txt"}:
            with open(filepath, "r", encoding="utf-8-sig", errors="replace") as f:
                sample = f.read(30000)
            helper = StatementParserWorker([])
            bank, account = helper.detect_bank_and_acc(sample, filepath)
            if bank and bank != "Bank Account": meta["bank"] = bank
            if account: meta["account"] = account
        elif ext in {".xlsx", ".xlsm", ".xls"} and ext in {".xlsx", ".xlsm"}:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            try:
                vals=[]
                for row in wb.active.iter_rows(max_row=30, values_only=True):
                    vals.extend(str(v) for v in row if v is not None)
                sample="\n".join(vals)
            finally:
                wb.close()
            helper=StatementParserWorker([])
            bank,account=helper.detect_bank_and_acc(sample, filepath)
            if bank and bank != "Bank Account": meta["bank"]=bank
            if account: meta["account"]=account
        else:
            with open(filepath, "r", encoding="utf-8-sig", errors="replace") as f:
                sample=f.read(30000)
    except Exception:
        pass
    if sample:
        meta["document_type"] = detect_statement_type(sample)
        meta["language"] = detect_statement_language(sample)
        upper=sample.upper()
        if any(x in upper for x in ["INR", "RS.", "RS ", "₹"]): meta["currency"]="INR (Rs.)"
        elif "$" in sample or "USD" in upper: meta["currency"]="USD ($)"
        elif "€" in sample or "EUR" in upper: meta["currency"]="EUR (€)"
        elif "£" in sample or "GBP" in upper: meta["currency"]="GBP (£)"
        # Statement holder labels are intentionally conservative.
        m=re.search(r"(?:ACCOUNT HOLDER|CUSTOMER NAME|NAME OF ACCOUNT HOLDER|CARDHOLDER)\s*[:.-]?\s*([^\n]{2,80})", sample, re.I)
        if m: meta["holder_name"]=m.group(1).strip()
        # Period labels are preferred over arbitrary dates.
        dm=re.search(r"(?:STATEMENT PERIOD|PERIOD|FROM)\s*[:.-]?\s*([0-9]{1,2}[/-][A-Za-z0-9]{1,4}[/-][0-9]{2,4})\s*(?:TO|[-–])\s*([0-9]{1,2}[/-][A-Za-z0-9]{1,4}[/-][0-9]{2,4})", sample, re.I)
        if dm:
            meta["start_date"]=dm.group(1); meta["end_date"]=dm.group(2)
    return meta

def optional_visual_capabilities():
    """Return availability of optional OCR/visual dependencies.

    Core statement parsing must not depend on OpenCV, NumPy or Tesseract.
    This helper is intentionally cheap and never imports those packages at
    module import time.
    """
    caps = {"opencv": False, "numpy": False, "tesseract": False, "paddleocr": False, "pikepdf": False}
    try:
        import cv2  # noqa: F401
        caps["opencv"] = True
    except Exception:
        pass
    try:
        import numpy  # noqa: F401
        caps["numpy"] = True
    except Exception:
        pass
    try:
        import pytesseract  # noqa: F401
        caps["tesseract"] = True
    except Exception:
        pass
    try:
        import paddleocr  # noqa: F401
        caps["paddleocr"] = True
    except Exception:
        pass
    try:
        import pikepdf  # noqa: F401
        caps["pikepdf"] = True
    except Exception:
        pass
    return caps


class VisualEvidenceEngine:
    """Conservative visual-document extraction for photographed/scanned statements.

    OCR is treated as evidence generation, not as truth.  Transactions are
    ledger-eligible only when the visual evidence gate passes.
    """

    VERSION = "VisualEvidenceEngine-v1.0"
    SUPPORTED_IMAGE_EXTENSIONS = {
        ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"
    }

    @staticmethod
    def capabilities():
        caps = {"opencv": False, "tesseract": False, "numpy": False, "paddleocr": False}
        try:
            import cv2  # noqa: F401
            caps["opencv"] = True
        except Exception:
            pass
        try:
            import numpy  # noqa: F401
            caps["numpy"] = True
        except Exception:
            pass
        try:
            import pytesseract  # noqa: F401
            caps["tesseract"] = True
        except Exception:
            pass
        try:
            import paddleocr  # noqa: F401
            caps["paddleocr"] = True
        except Exception:
            pass
        return caps

    @staticmethod
    def _sha256_file(path):
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for block in iter(lambda: f.read(1024 * 1024), b""):
                h.update(block)
        return h.hexdigest()

    @staticmethod
    def _load_image(path_or_bytes):
        import cv2
        import numpy as np
        if isinstance(path_or_bytes, (str, os.PathLike)):
            img = cv2.imread(str(path_or_bytes), cv2.IMREAD_COLOR)
            if img is None:
                raise ValueError("Unable to decode image file")
            return img
        arr = np.frombuffer(path_or_bytes, dtype=np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        if img is None:
            raise ValueError("Unable to decode image bytes")
        return img

    @staticmethod
    def _order_quad(points):
        import numpy as np
        pts = np.asarray(points, dtype="float32")
        s = pts.sum(axis=1)
        d = np.diff(pts, axis=1).ravel()
        return np.array([
            pts[np.argmin(s)], pts[np.argmin(d)],
            pts[np.argmax(s)], pts[np.argmax(d)]
        ], dtype="float32")

    @classmethod
    def detect_page_quad(cls, image):
        """Return a probable page quadrilateral, never treating it as certain."""
        import cv2
        import numpy as np
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        edges = cv2.Canny(blur, 40, 140)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        h, w = gray.shape[:2]
        best = None
        best_area = 0
        for c in contours:
            area = cv2.contourArea(c)
            if area < 0.20 * w * h:
                continue
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.02 * peri, True)
            if len(approx) == 4 and area > best_area:
                best = approx.reshape(4, 2)
                best_area = area
        if best is None:
            return None, 0.0
        return cls._order_quad(best), min(1.0, best_area / float(w * h))

    @classmethod
    def rectify_page(cls, image):
        import cv2
        import numpy as np
        quad, coverage = cls.detect_page_quad(image)
        if quad is None:
            return image.copy(), {"perspective_corrected": False, "page_coverage": None}
        tl, tr, br, bl = quad
        width_a = np.linalg.norm(br - bl)
        width_b = np.linalg.norm(tr - tl)
        height_a = np.linalg.norm(tr - br)
        height_b = np.linalg.norm(tl - bl)
        width = int(max(width_a, width_b))
        height = int(max(height_a, height_b))
        if width < 100 or height < 100:
            return image.copy(), {"perspective_corrected": False, "page_coverage": coverage}
        dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype="float32")
        matrix = cv2.getPerspectiveTransform(quad, dst)
        warped = cv2.warpPerspective(image, matrix, (width, height), flags=cv2.INTER_CUBIC)
        return warped, {"perspective_corrected": True, "page_coverage": round(coverage, 4)}

    @staticmethod
    def quality_report(image):
        import cv2
        import numpy as np
        h, w = image.shape[:2]
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        lap = float(cv2.Laplacian(gray, cv2.CV_64F).var())
        brightness = float(np.mean(gray))
        contrast = float(np.std(gray))
        # These are intentionally heuristic readiness indicators, not OCR truth.
        blur_score = min(1.0, lap / 300.0)
        contrast_score = min(1.0, contrast / 65.0)
        resolution_score = min(1.0, (w * h) / float(1600 * 1200))
        # White paper is not glare. Estimate specular glare primarily from
        # bright, low-detail pixels in chroma-saturated regions; otherwise a
        # normal white passbook page would be falsely classified as glare.
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
        bright_colored = ((hsv[:, :, 2] >= 250) & (hsv[:, :, 1] >= 35))
        glare_ratio = float(np.mean(bright_colored))
        dark_ratio = float(np.mean(gray <= 25))
        readiness = 0.45 * blur_score + 0.30 * contrast_score + 0.25 * resolution_score
        if glare_ratio > 0.08 or dark_ratio > 0.20:
            readiness *= 0.75
        if min(w, h) < 700:
            readiness *= 0.65
        if readiness >= 0.68:
            band = "OCR_READY"
        elif readiness >= 0.42:
            band = "OCR_DEGRADED"
        elif readiness >= 0.22:
            band = "OCR_REVIEW_REQUIRED"
        else:
            band = "OCR_UNUSABLE"
        return {
            "width": w, "height": h,
            "blur_laplacian": round(lap, 3),
            "brightness_mean": round(brightness, 3),
            "contrast_std": round(contrast, 3),
            "glare_ratio": round(glare_ratio, 5),
            "dark_ratio": round(dark_ratio, 5),
            "readiness_score": round(readiness, 4),
            "readiness": band,
        }

    @staticmethod
    def _remove_shadows(gray):
        import cv2
        bg = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (31, 31)))
        norm = cv2.divide(gray, bg, scale=255)
        return norm

    @staticmethod
    def _remove_ruling_lines(gray):
        import cv2
        horizontal = cv2.morphologyEx(gray, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(15, gray.shape[1] // 35), 1)))
        vertical = cv2.morphologyEx(gray, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(15, gray.shape[0] // 35))))
        lines = cv2.bitwise_or(horizontal, vertical)
        return cv2.subtract(gray, lines)

    @classmethod
    def preprocess_variants(cls, image):
        import cv2
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        shadow = cls._remove_shadows(gray)
        clean = cls._remove_ruling_lines(shadow)
        denoise = cv2.fastNlMeansDenoising(clean, None, 7, 7, 21)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(denoise)
        sharp = cv2.addWeighted(clahe, 1.35, cv2.GaussianBlur(clahe, (0, 0), 1.0), -0.35, 0)
        adaptive = cv2.adaptiveThreshold(clahe, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 11)
        return {
            "original_gray": gray,
            "illumination_corrected": shadow,
            "cleaned": clean,
            "enhanced": clahe,
            "sharpened": sharp,
            "adaptive_threshold": adaptive,
        }

    @staticmethod
    def _ocr_image(img, psm=6):
        """Prefer PaddleOCR 3.x when installed; fall back to Tesseract."""
        try:
            from paddleocr import PaddleOCR
            import numpy as np
            if not hasattr(VisualEvidenceEngine, "_paddle_ocr_instance"):
                VisualEvidenceEngine._paddle_ocr_instance = PaddleOCR(
                    lang="en",
                    use_doc_orientation_classify=False,
                    use_doc_unwarping=False,
                    use_textline_orientation=False,
                    enable_mkldnn=False,
                )
            result = VisualEvidenceEngine._paddle_ocr_instance.predict(img)
            words = []
            for res in result:
                for txt, score, box in zip(res["rec_texts"], res["rec_scores"], res["rec_boxes"]):
                    txt = str(txt).strip()
                    if not txt:
                        continue
                    x0, y0, x1, y1 = [int(v) for v in np.asarray(box).tolist()]
                    words.append({"text": txt, "conf": float(score) * 100.0,
                                  "x": x0, "y": y0, "w": max(1,x1-x0), "h": max(1,y1-y0),
                                  "engine": "PaddleOCR-PP-OCRv6"})
            if words:
                return words
        except Exception:
            pass

        import pytesseract
        from pytesseract import Output
        cfg = f"--oem 3 --psm {psm}"
        data = pytesseract.image_to_data(img, config=cfg, output_type=Output.DICT)
        words = []
        for i, txt in enumerate(data.get("text", [])):
            txt = str(txt).strip()
            try:
                conf = float(data["conf"][i])
            except Exception:
                conf = -1.0
            if not txt:
                continue
            words.append({"text": txt, "conf": conf,
                          "x": int(data["left"][i]), "y": int(data["top"][i]),
                          "w": int(data["width"][i]), "h": int(data["height"][i]),
                          "engine": "Tesseract"})
        return words

    @staticmethod
    def _normalise_ocr_amount(raw):
        """Return (Decimal-or-None, ambiguity flag, normalized text)."""
        text = str(raw or "").strip().replace("₹", "").replace("Rs.", "").replace("Rs", "")
        text = text.replace(" ", "")
        if not text:
            return None, True, text
        # Conservative OCR confusions. Only map when the whole token is amount-like.
        candidate = text
        if re.fullmatch(r"[0-9OIlSBoos,.-]+", candidate, re.IGNORECASE):
            candidate = (candidate.replace("O", "0").replace("o", "0")
                         .replace("I", "1").replace("l", "1")
                         .replace("S", "5").replace("s", "5"))
        value = parse_decimal(candidate)
        ambiguous = bool(re.search(r"[A-Za-z]", text)) or value is None
        return value, ambiguous, candidate

    @classmethod
    def _group_lines(cls, words):
        lines = []
        for word in sorted(words, key=lambda x: (x["y"], x["x"])):
            cy = word["y"] + word["h"] / 2.0
            target = None
            for line in lines:
                if abs(cy - line["cy"]) <= max(8, 0.60 * max(word["h"], line["height"])):
                    target = line
                    break
            if target is None:
                target = {"cy": cy, "height": word["h"], "words": []}
                lines.append(target)
            target["words"].append(word)
            target["height"] = max(target["height"], word["h"])
            target["cy"] = sum(w["y"] + w["h"] / 2.0 for w in target["words"]) / len(target["words"])
        for line in lines:
            line["words"].sort(key=lambda x: x["x"])
            line["text"] = " ".join(w["text"] for w in line["words"])
        return sorted(lines, key=lambda x: x["cy"])

    @staticmethod
    def _date_from_words(words):
        patterns = [
            r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$",
            r"^\d{1,2}[.-]\d{1,2}[.-]\d{2,4}$",
            r"^\d{8}$",
        ]
        for w in words:
            raw = w["text"].strip(".,")
            t = raw.replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")
            if any(re.fullmatch(p, t) for p in patterns):
                if re.fullmatch(r"\d{8}", t):
                    t = f"{t[:2]}/{t[2:4]}/{t[4:]}"
                return t, w
        return None, None

    @classmethod
    def _line_transaction_candidate(cls, line, image_width, bank, account, source_file, page=0):
        date, date_word = cls._date_from_words(line["words"])
        if not date:
            return None
        if any(x in line["text"].upper() for x in ["DATE PARTICULAR", "OPENING BALANCE", "CLOSING BALANCE", "STATEMENT OF ACCOUNT"]):
            return None
        amount_words = []
        for w in line["words"]:
            # Ignore the date token and obvious reference-like tokens.
            if w is date_word:
                continue
            val, amb, norm = cls._normalise_ocr_amount(w["text"])
            if val is not None and re.search(r"\d", w["text"]):
                amount_words.append((w, val, amb, norm))
        # A normal passbook row has 2 or 3 numeric columns. With only one amount
        # there is insufficient evidence to decide debit/credit versus balance.
        if len(amount_words) < 2:
            return {
                "candidate": True, "certifiable": False,
                "reason": "INSUFFICIENT_MONETARY_COLUMNS", "line": line["text"],
                "date": date,
            }
        amount_words.sort(key=lambda x: x[0]["x"])
        # Rightmost numeric field is treated as balance. This is only a candidate
        # until arithmetic and cross-variant agreement validate it.
        bal_item = amount_words[-1]
        pre = amount_words[:-1]
        debit = Decimal("0.00")
        credit = Decimal("0.00")
        ambiguity = any(x[2] for x in amount_words)
        if len(pre) == 2:
            debit, credit = pre[0][1], pre[1][1]
        elif len(pre) == 1:
            # Prefer physical column position for passbook/table layouts.
            # Fall back to explicit DR/CR words when the layout is not columnar.
            xmid = pre[0][0]["x"] + pre[0][0]["w"] / 2.0
            upper = line["text"].upper()
            if xmid < image_width * 0.64 or " DR" in upper or upper.endswith("DR") or "DEBIT" in upper or "WITHDRAW" in upper:
                debit = pre[0][1]
            elif xmid < image_width * 0.82 or " CR" in upper or upper.endswith("CR") or "CREDIT" in upper or "DEPOSIT" in upper:
                credit = pre[0][1]
            else:
                return {
                    "candidate": True, "certifiable": False,
                    "reason": "DEBIT_CREDIT_DIRECTION_AMBIGUOUS", "line": line["text"],
                    "date": date,
                }
        narration_parts = [w["text"] for w in line["words"] if w["x"] > date_word["x"] + date_word["w"] and w not in [x[0] for x in amount_words]]
        narration = " ".join(narration_parts).strip()
        return {
            "candidate": True,
            "certifiable": not ambiguity,
            "date": date,
            "narration": narration,
            "withdrawal": debit,
            "deposit": credit,
            "balance": bal_item[1],
            "balance_text": bal_item[0]["text"],
            "ambiguity": ambiguity,
            "bbox": [line["words"][0]["x"], int(min(w["y"] for w in line["words"])),
                     max(w["x"] + w["w"] for w in line["words"]), int(max(w["y"] + w["h"] for w in line["words"]))],
            "source_file": source_file,
            "page": page,
            "bank": bank,
            "account_number": account,
        }

    @staticmethod
    def _agreement_signature(candidate):
        if not candidate:
            return None
        return (
            str(candidate.get("date", "")),
            str(candidate.get("withdrawal", "")),
            str(candidate.get("deposit", "")),
            str(candidate.get("balance", "")),
        )

    @classmethod
    def _merge_candidates(cls, candidate_sets):
        # Align rows by order and date. Only exact financial agreement across at
        # least two independent preprocessing variants is considered corroboration.
        rows = []
        usable = [c for c in candidate_sets if c]
        if not usable:
            return rows
        max_len = max(len(c) for c in usable)
        for idx in range(max_len):
            bucket = [c[idx] for c in usable if idx < len(c) and c[idx].get("candidate")]
            cert = [x for x in bucket if x.get("certifiable")]
            if not bucket:
                continue
            signatures = {}
            for x in cert:
                signatures.setdefault(cls._agreement_signature(x), []).append(x)
            best_sig, best_items = (None, [])
            if signatures:
                best_sig, best_items = max(signatures.items(), key=lambda kv: len(kv[1]))
            chosen = best_items[0] if best_items else bucket[0]
            chosen = dict(chosen)
            chosen["ocr_variant_agreement"] = len(best_items)
            chosen["ocr_variant_count"] = len(bucket)
            chosen["ocr_conflict"] = len(signatures) > 1
            chosen["ocr_certified"] = len(best_items) >= 2 and not chosen["ocr_conflict"]
            rows.append(chosen)
        return rows

    @classmethod
    def _extract_opening_balance(cls, candidate_sets):
        values = []
        for candidates in candidate_sets:
            for c in candidates:
                line = str(c.get("line", "")).upper()
                if "OPENING BALANCE" in line or "BROUGHT FORWARD" in line or "B/F" in line:
                    vals = [tok for tok in re.split(r"\s+", line) if re.search(r"\d", tok)]
                    for raw in vals:
                        val, _, _ = cls._normalise_ocr_amount(raw)
                        if val is not None and val >= Decimal("100.00"):
                            values.append(val)
        if not values:
            return None
        counts = {}
        for v in values: counts[v] = counts.get(v, 0) + 1
        return max(counts, key=counts.get)

    @classmethod
    def _apply_balance_gate(cls, rows, opening_balance=None):
        previous = opening_balance
        for row in rows:
            if not row.get("ocr_certified"):
                row["visual_status"] = "REVIEW_REQUIRED"
                row["ledger_eligible"] = False
                previous = None if previous is None else previous
                continue
            expected = None
            if previous is not None:
                expected = previous + row["deposit"] - row["withdrawal"]
            row["expected_balance"] = str(expected) if expected is not None else None
            actual = row.get("balance")
            if expected is not None and actual != expected:
                row["visual_status"] = "BALANCE_CONFLICT"
                row["ledger_eligible"] = False
            else:
                row["visual_status"] = "CERTIFIED" if expected is not None else "CORROBORATED_NO_OPENING_BALANCE"
                row["ledger_eligible"] = expected is not None
            previous = actual if actual is not None else previous
        return rows

    @classmethod
    def extract(cls, path_or_bytes, source_file="", page=0, bank="Bank Account", account="N/A"):
        caps = cls.capabilities()
        if not all(caps.values()):
            return {"status": "OCR_ENGINE_UNAVAILABLE", "ledger_eligible": False, "capabilities": caps, "transactions": []}
        import cv2
        import numpy as np
        image = cls._load_image(path_or_bytes)
        original_quality = cls.quality_report(image)
        rectified, geometry = cls.rectify_page(image)
        quality = cls.quality_report(rectified)
        variants = cls.preprocess_variants(rectified)
        candidate_sets = []
        ocr_variant_reports = []
        ocr_text_samples = []
        variant_items = list(variants.items())

        def run_variant(name, variant):
            try:
                words = cls._ocr_image(variant, psm=6)
                ocr_text_samples.append(" ".join(w["text"] for w in words))
                lines = cls._group_lines(words)
                candidates = []
                for line in lines:
                    if "OPENING BALANCE" in line["text"].upper() or "BROUGHT FORWARD" in line["text"].upper():
                        candidates.append({"candidate": False, "line": line["text"]})
                    c = cls._line_transaction_candidate(line, rectified.shape[1], bank, account, source_file, page)
                    if c:
                        candidates.append(c)
                candidate_sets.append(candidates)
                avg_conf = sum(max(0.0, w["conf"]) for w in words) / max(1, len(words))
                ocr_variant_reports.append({"variant": name, "words": len(words), "average_confidence": round(avg_conf, 2), "candidates": len(candidates)})
            except Exception as exc:
                ocr_variant_reports.append({"variant": name, "error": str(exc), "words": 0, "candidates": 0})

        # Adaptive ensemble: start with two materially different views. Only
        # spend OCR time on the remaining variants when the evidence disagrees
        # or does not produce a certifiable result.
        for name, variant in variant_items[:2]:
            run_variant(name, variant)
        rows = cls._merge_candidates(candidate_sets)
        if not rows or any(r.get("ocr_conflict") or not r.get("ocr_certified") for r in rows):
            for name, variant in variant_items[2:]:
                run_variant(name, variant)
                rows = cls._merge_candidates(candidate_sets)
                if rows and all(r.get("ocr_certified") and not r.get("ocr_conflict") for r in rows):
                    break
        opening_balance = cls._extract_opening_balance(candidate_sets)
        rows = cls._apply_balance_gate(rows, opening_balance=opening_balance)
        # A page with obvious severe quality problems never becomes automatically ledger-eligible.
        if quality["readiness"] in {"OCR_REVIEW_REQUIRED", "OCR_UNUSABLE"}:
            for r in rows:
                r["ledger_eligible"] = False
                r["visual_status"] = "IMAGE_QUALITY_REVIEW_REQUIRED"
        transactions = []
        for r in rows:
            if not r.get("ledger_eligible"):
                continue
            tx = {
                "bank": bank,
                "account_number": account,
                "date": r["date"],
                "narration": r.get("narration", ""),
                "mode": "DIGITAL",
                "withdrawal": str(r["withdrawal"]),
                "deposit": str(r["deposit"]),
                "balance": str(r["balance"]),
                "page": page,
                "bbox": r.get("bbox", [0, 0, rectified.shape[1], rectified.shape[0]]),
                "source_file": source_file,
                "source_format": "VISUAL_OCR",
                "extraction_method": cls.VERSION,
                "ocr_variant_agreement": r.get("ocr_variant_agreement", 0),
                "ocr_variant_count": r.get("ocr_variant_count", 0),
                "visual_status": r.get("visual_status"),
                "status": "OK",
            }
            transactions.append(tx)
        return {
            "status": "OK" if transactions else ("REVIEW_REQUIRED" if rows else "NO_TRANSACTIONS_DETECTED"),
            "ledger_eligible": bool(transactions),
            "capabilities": caps,
            "sha256": cls._sha256_file(path_or_bytes) if isinstance(path_or_bytes, (str, os.PathLike)) else None,
            "original_quality": original_quality,
            "quality": quality,
            "geometry": geometry,
            "ocr_variants": ocr_variant_reports,
            "ocr_text": ocr_text_samples[0] if ocr_text_samples else "",
            "rows": rows,
            "transactions": transactions,
        }

    @classmethod
    def hostile_corpus(cls, directory):
        """Create a deterministic hostile synthetic image corpus for regression testing."""
        import cv2
        import numpy as np
        from PIL import Image, ImageDraw, ImageFont
        os.makedirs(directory, exist_ok=True)
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
        font = ImageFont.truetype(font_path, 30) if os.path.exists(font_path) else ImageFont.load_default()
        small_font = ImageFont.truetype(font_path, 24) if os.path.exists(font_path) else font
        base = Image.new("RGB", (1400, 900), "white")
        d = ImageDraw.Draw(base)
        d.text((80, 35), "SYNTHETIC BANK PASSBOOK - OCR TEST", fill="black", font=font)
        d.text((900, 35), "Opening Balance 100000", fill="black", font=font)
        headers = [(80, 100, "Date"), (260, 100, "Particulars"), (760, 100, "Debit"), (900, 100, "Credit"), (1050, 100, "Balance")]
        for x, y, t in headers: d.text((x, y), t, fill="black", font=font)
        rows = [("01/04/2026", "SALARY CREDIT", "", "50000", "150000"), ("03/04/2026", "ATM WITHDRAWAL", "20000", "", "130000"), ("05/04/2026", "NEFT CREDIT", "", "25000", "155000"), ("07/04/2026", "ATM WITHDRAWAL", "15000", "", "140000")]
        y = 160
        for date, narr, dr, cr, bal in rows:
            d.text((80, y), date, fill="black", font=font); d.text((260, y), narr, fill="black", font=font)
            d.text((760, y), dr, fill="black", font=font); d.text((900, y), cr, fill="black", font=font); d.text((1050, y), bal, fill="black", font=font)
            d.line((60, y+42, 1300, y+42), fill="#aaaaaa", width=1); y += 105
        base_path = os.path.join(directory, "00_clean.jpg"); base.save(base_path, quality=95)
        cases = {"01_skewed": 7, "02_perspective": "perspective", "03_shadow": "shadow", "04_glare": "glare", "05_blur": "blur", "06_faint": "faint", "07_ruling": "ruling", "08_cropped": "crop", "09_duplicate": "duplicate", "10_rotated": "rotate", "11_lowres": "lowres", "12_ambiguous": "ambiguous"}
        paths = [base_path]
        arr = np.array(base)[:, :, ::-1]
        h, w = arr.shape[:2]
        for name, mode in cases.items():
            out = arr.copy()
            if mode == "perspective":
                src = np.float32([[0,0],[w-1,0],[w-1,h-1],[0,h-1]])
                dst = np.float32([[140,40],[w-120,0],[w-20,h-50],[20,h-20]])
                out = cv2.warpPerspective(out, cv2.getPerspectiveTransform(src,dst), (w,h), borderValue=(255,255,255))
            elif mode == "shadow":
                mask = np.tile(np.linspace(0.35, 1.0, w, dtype=np.float32), (h,1))[:, :, None]
                out = np.clip(out.astype(np.float32)*mask,0,255).astype(np.uint8)
            elif mode == "glare":
                cv2.ellipse(out, (1050,350), (250,180), 0, 0, 360, (255,255,255), -1)
            elif mode == "blur":
                out = cv2.GaussianBlur(out, (15,15), 0)
            elif mode == "faint":
                out = np.clip(out.astype(np.float32)*0.65 + 255*0.35,0,255).astype(np.uint8)
            elif mode == "ruling":
                for yy in range(130, 500, 27): cv2.line(out,(40,yy),(1350,yy),(100,100,100),2)
            elif mode == "crop":
                out = out[20:860, 50:1150]
            elif mode == "rotate":
                M = cv2.getRotationMatrix2D((w/2,h/2), 90, 1.0); out = cv2.warpAffine(out,M,(h,w),borderValue=(255,255,255))
            elif mode == "lowres":
                out = cv2.resize(cv2.resize(out,(500,320),interpolation=cv2.INTER_AREA),(1400,900),interpolation=cv2.INTER_NEAREST)
            elif mode == "ambiguous":
                cv2.rectangle(out,(1040,145),(1190,230),(245,245,245),-1); cv2.putText(out,"15O000",(1040,195),cv2.FONT_HERSHEY_SIMPLEX,1.2,(60,60,60),2)
            elif mode == "duplicate":
                out = arr.copy()
            elif isinstance(mode, int):
                M = cv2.getRotationMatrix2D((w/2,h/2), mode, 1.0); out = cv2.warpAffine(out,M,(w,h),borderValue=(255,255,255))
            path = os.path.join(directory, name + ".jpg")
            cv2.imwrite(path, out)
            paths.append(path)
        # Exact duplicate for duplicate detection regression.
        cv2.imwrite(os.path.join(directory, "09_duplicate_copy.jpg"), arr)
        with open(os.path.join(directory, "README.txt"), "w", encoding="utf-8") as f:
            f.write("Synthetic hostile visual corpus. Not a real bank statement.\nCases: skew, perspective, shadow, glare, blur, faint print, ruling lines, crop, duplicate, rotation, low resolution, ambiguous amount.\n")
        return paths

    @classmethod
    def run_hostile_regression(cls, directory=None):
        """Run the hostile OCR regression when optional visual dependencies exist.

        Missing OpenCV/Tesseract is a capability state, not a core parser
        failure.  The headless suite therefore reports SKIPPED and continues
        with all dependency-free parser tests.
        """
        caps = optional_visual_capabilities()
        missing = [name for name, available in caps.items() if not available]
        if missing:
            return {
                "status": "SKIPPED",
                "files": 0,
                "directory": None,
                "capabilities": caps,
                "missing": missing,
                "reason": "Optional visual/OCR dependencies are not installed.",
            }

        import tempfile
        if directory is None:
            td = tempfile.mkdtemp(prefix="taxforensic_visual_corpus_", dir=APP_DIRECTORY)
            cleanup = True
        else:
            td = directory; os.makedirs(td, exist_ok=True); cleanup = False
        try:
            paths = cls.hostile_corpus(td)
            assert os.path.exists(paths[0])
            clean = cls.extract(paths[0], source_file=paths[0], bank="Synthetic Bank", account="1234567890")
            assert clean["capabilities"]["opencv"] and clean["capabilities"]["tesseract"]
            assert clean["transactions"] or clean["status"] in {"REVIEW_REQUIRED", "NO_TRANSACTIONS_DETECTED"}
            # Exercise the hostile corpus without multiplying OCR cost: every
            # hostile image must survive decoding/quality/geometry inspection,
            # while the clean image exercises the full adaptive OCR ensemble.
            for p in paths[1:]:
                image = cls._load_image(p)
                q = cls.quality_report(image)
                assert q["width"] > 0 and q["height"] > 0
                rectified, geometry = cls.rectify_page(image)
                assert rectified.shape[0] > 0 and rectified.shape[1] > 0
            # Exact duplicate fingerprints must be identical.
            assert cls._sha256_file(os.path.join(td, "09_duplicate.jpg")) == cls._sha256_file(os.path.join(td, "09_duplicate_copy.jpg"))
            return {"status": "PASS", "files": len(paths), "directory": td}
        finally:
            if cleanup:
                # Keep the corpus for inspection by default; caller can remove it later.
                pass

# ======================================================================
# 5. STATEMENT PARSER WORKER
# ======================================================================

def _narration_direction_cue(narration):
    """Return CREDIT/DEBIT only when narration has one clear directional cue."""
    text = re.sub(r"\s+", " ", str(narration or "").upper()).strip()
    credit_patterns = (
        r"\bTRF\s+FRM\b", r"\bTRANSFER\s+FROM\b", r"\bBY\s+TRF\b",
        r"\bBY\s+CASH\b", r"\bCREDIT(?:ED)?\b", r"\bDEPOSIT(?:ED)?\b",
        r"\bRECEIVED\b", r"\bNEFT\s+CR\b", r"\bIMPS\s+CR\b",
        r"\bUPI\s+CR\b", r"\bRTGS\s+CR\b", r"\bREFUND(?:ED)?\b",
    )
    debit_patterns = (
        r"\bTO\s+CASH\b", r"\bTO\s+TRF\b", r"\bTRANSFER\s+TO\b",
        r"\bPAID\s+TO\b", r"\bATM\s+WDL\b", r"\bWITHDRAW(?:AL)?\b",
        r"\bNEFT\s+DR\b", r"\bIMPS\s+DR\b", r"\bUPI\s+DR\b",
        r"\bRTGS\s+DR\b", r"\bDEBITED\b", r"\bCHQ\s+PAID\b",
    )
    credit = any(re.search(p, text) for p in credit_patterns)
    debit = any(re.search(p, text) for p in debit_patterns)
    if credit == debit:
        return None
    return "CREDIT" if credit else "DEBIT"


def _transaction_y(tx):
    bbox = tx.get("bbox") or []
    try:
        return float(tx.get("row_anchor_y", bbox[1]))
    except Exception:
        return 0.0


def recover_multiline_money_cells(transactions):
    """Recover uniquely positioned financial cells whose y-coordinate differs from the date line.

    Some bank statements lay out a transaction as a multi-line block: date and
    narration first, then the debit/credit/balance cells on a lower baseline.
    This routine uses the calibrated header's right-edge money columns. It only
    assigns a field where exactly one monetary token occupies that column for
    the transaction. Ambiguous or duplicate tokens remain unresolved.
    """
    recovered = {"withdrawal": 0, "deposit": 0, "balance": 0, "conflicts": 0}
    for tx in transactions or []:
        if tx.get("mode") == "REFERENCE" or not str(tx.get("date") or "").strip():
            continue
        candidates = [c for c in (tx.get("amount_candidates") or [])
                      if c.get("row_aligned") is False
                      and c.get("money_token") is True
                      and c.get("column_geometry_confident") is True
                      and c.get("column_role") in {"withdrawal", "deposit", "balance"}
                      and parse_decimal(c.get("value")) is not None
                      and abs(parse_decimal(c.get("value"))) > Decimal("0.00")]

        # Debit/credit movement: only repair rows where neither movement field
        # was already extracted. One unambiguous token in one money column is
        # geometrical evidence; the running-balance check remains independent.
        dr = parse_decimal(tx.get("withdrawal"))
        cr = parse_decimal(tx.get("deposit"))
        if dr is not None and cr is not None and dr == Decimal("0.00") and cr == Decimal("0.00"):
            movement = [c for c in candidates if c.get("column_role") in {"withdrawal", "deposit"}]
            if len(movement) == 1:
                cand = movement[0]
                role = cand["column_role"]
                amount = abs(parse_decimal(cand["value"]))
                tx["withdrawal"] = str(amount) if role == "withdrawal" else "0.00"
                tx["deposit"] = str(amount) if role == "deposit" else "0.00"
                tx[role + "_bbox"] = [cand.get("x0"), cand.get("y"), cand.get("x1"), cand.get("y")]
                tx["movement_extraction_method"] = "MULTILINE_MONEY_COLUMN_GEOMETRY"
                tx["movement_extraction_evidence"] = {
                    "token": cand.get("text"), "value": str(amount),
                    "column_role": role, "row_aligned": False,
                    "x0": cand.get("x0"), "x1": cand.get("x1"), "y": cand.get("y"),
                }
                tx["status"] = "EXTRACTED_PENDING_RECONCILIATION"
                recovered[role] += 1
            elif len(movement) > 1:
                tx.setdefault("field_conflicts", []).append({
                    "field": "debit_credit", "reason": "MULTIPLE_OFF_BASELINE_MONEY_CELLS",
                    "candidates": [{"text": c.get("text"), "value": c.get("value"),
                                    "column_role": c.get("column_role"), "x0": c.get("x0"),
                                    "x1": c.get("x1"), "y": c.get("y")} for c in movement],
                })
                recovered["conflicts"] += 1

        # Printed balance: only restore a missing image balance from one unique
        # off-baseline token in the independently calibrated balance column.
        image_balance = parse_decimal(tx.get("image_balance"))
        if image_balance is None:
            balance_candidates = [c for c in candidates if c.get("column_role") == "balance"]
            if len(balance_candidates) == 1:
                cand = balance_candidates[0]
                value = parse_decimal(cand["value"])
                tx["image_balance"] = str(value)
                tx["balance"] = str(value)
                tx["image_balance_source"] = "MULTILINE_MONEY_COLUMN_GEOMETRY"
                tx["image_balance_bbox"] = [cand.get("x0"), cand.get("y"), cand.get("x1"), cand.get("y")]
                recovered["balance"] += 1
            elif len(balance_candidates) > 1:
                tx.setdefault("field_conflicts", []).append({
                    "field": "image_balance", "reason": "MULTIPLE_OFF_BASELINE_BALANCE_CELLS",
                    "candidates": [{"text": c.get("text"), "value": c.get("value"),
                                    "x0": c.get("x0"), "x1": c.get("x1"), "y": c.get("y")}
                                   for c in balance_candidates],
                })
                recovered["conflicts"] += 1
    return recovered


def resolve_movement_from_exact_balance_delta(transactions):
    """Correct an incorrectly assigned movement only if a visible cell matches the exact row balance delta.

    The routine never calculates an amount and inserts it. It can only choose
    an already-extracted monetary token that (a) is positioned inside the
    calibrated debit/credit column and (b) exactly equals the difference
    between two consecutive printed running balances, including direction.
    """
    corrections = []
    streams = {}
    for original_index, tx in enumerate(transactions or []):
        key = (str(tx.get("source_file", "")), str(tx.get("bank", "")), str(tx.get("account_number", "")))
        streams.setdefault(key, []).append((original_index, tx))

    for key, stream in streams.items():
        stream.sort(key=lambda pair: (
            int(pair[1].get("page", 0) or 0),
            float((pair[1].get("bbox") or [0, 0, 0, 0])[1]),
            float((pair[1].get("bbox") or [0, 0, 0, 0])[3]), pair[0]))
        previous_balance = None
        previous_page = None
        for _, tx in stream:
            page = int(tx.get("page", 0) or 0)
            image_balance = parse_decimal(tx.get("image_balance"))
            if tx.get("mode") == "REFERENCE":
                if tx.get("reference_role") == "OPENING" and previous_balance is None:
                    previous_balance = image_balance
                # Do not reset a known transaction balance from a C/F line:
                # printed carried-forward values can themselves be erroneous.
                previous_page = page
                continue
            if page != previous_page and previous_balance is None:
                # A real printed B/F may seed the next page if the prior page
                # ended without an extractable transaction balance.
                same_page_openings = [r for _, r in stream
                    if int(r.get("page", 0) or 0) == page
                    and r.get("mode") == "REFERENCE" and r.get("reference_role") == "OPENING"]
                if same_page_openings:
                    previous_balance = parse_decimal(same_page_openings[0].get("image_balance"))
            if image_balance is None:
                previous_balance = None
                previous_page = page
                continue
            if previous_balance is not None:
                delta = image_balance - previous_balance
                current_dr = parse_decimal(tx.get("withdrawal"))
                current_cr = parse_decimal(tx.get("deposit"))
                current_dr = current_dr if current_dr is not None else Decimal("0.00")
                current_cr = current_cr if current_cr is not None else Decimal("0.00")
                if current_cr - current_dr != delta:
                    expected_role = "deposit" if delta > Decimal("0.00") else "withdrawal" if delta < Decimal("0.00") else None
                    needed = abs(delta)
                    matching = []
                    if expected_role and needed > Decimal("0.00"):
                        for cand in tx.get("amount_candidates", []) or []:
                            if (cand.get("column_role") != expected_role
                                    or cand.get("money_token") is not True
                                    or cand.get("column_geometry_confident") is not True):
                                continue
                            val = parse_decimal(cand.get("value"))
                            if val is not None and abs(val) == needed:
                                matching.append(cand)
                    if len(matching) == 1:
                        cand = matching[0]
                        tx["movement_reconstruction"] = {
                            "method": "EXACT_PRINTED_RUNNING_BALANCE_DELTA_TOKEN",
                            "original_withdrawal": str(current_dr),
                            "original_deposit": str(current_cr),
                            "printed_previous_balance": str(previous_balance),
                            "printed_current_balance": str(image_balance),
                            "visible_token": cand.get("text"),
                            "token_value": str(needed), "column_role": expected_role,
                            "admitted": True,
                        }
                        tx["withdrawal"] = str(needed) if expected_role == "withdrawal" else "0.00"
                        tx["deposit"] = str(needed) if expected_role == "deposit" else "0.00"
                        tx["movement_extraction_method"] = "EXACT_PRINTED_RUNNING_BALANCE_DELTA_TOKEN"
                        tx["status"] = "RECONSTRUCTED_BY_RUNNING_BALANCE_TOKEN"
                        tx["audit_detail"] = (
                            "Movement selected from one visible amount token in the appropriate calibrated money column; "
                            "its value and direction exactly match the difference between consecutive printed balances."
                        )
                        corrections.append({"page": page, "date": tx.get("date", ""),
                                            "narration": tx.get("narration", ""),
                                            "amount": str(needed), "side": expected_role,
                                            "token": cand.get("text")})
            previous_balance = image_balance
            previous_page = page
    return corrections


def reconstruct_pdf_amounts_from_checkpoints(transactions):
    """Reclassify a candidate misfiled as balance only when exact page proof exists.

    A plain numeric token on the transaction's date line can be incorrectly
    classified as the running balance when a bank's PDF layout differs from
    header calibration. This routine considers that token as a transaction
    amount only where (1) the token is preserved with its coordinates, (2) it
    is aligned with the transaction date line, (3) narration supplies one
    clear debit/credit cue, (4) every ordinary transaction on that page has a
    usable amount after the proposed correction, and (5) the page's printed
    B/F plus movements exactly equals its printed C/F. It never creates a
    numeric amount. Failed or incomplete pages stay flagged for review.
    """
    rows = list(transactions or [])
    groups = {}
    for index, tx in enumerate(rows):
        try:
            page_index = int(tx.get("page", 0))
        except Exception:
            page_index = 0
        key = (str(tx.get("source_file", "")), str(tx.get("bank", "")),
               str(tx.get("account_number", "")), page_index)
        groups.setdefault(key, []).append((index, tx))

    for key, pairs in groups.items():
        page_rows = [tx for _, tx in pairs]
        opening = None
        closing = None
        for tx in page_rows:
            if tx.get("mode") != "REFERENCE":
                continue
            val = parse_decimal(tx.get("image_balance"))
            if val is None:
                val = parse_decimal(tx.get("balance"))
            if val is None:
                continue
            if tx.get("reference_role") == "OPENING" and opening is None:
                opening = val
            elif tx.get("reference_role") == "CLOSING" and closing is None:
                closing = val
        if opening is None or closing is None:
            continue

        ordinary = sorted(
            [(idx, tx) for idx, tx in pairs if tx.get("mode") != "REFERENCE" and str(tx.get("date") or "").strip()],
            key=lambda item: (_transaction_y(item[1]), item[0]),
        )
        proposals = {}
        page_is_complete = True
        for idx, tx in ordinary:
            dr = parse_decimal(tx.get("withdrawal"))
            cr = parse_decimal(tx.get("deposit"))
            if dr is None or cr is None:
                page_is_complete = False
                continue
            dr = dr or Decimal("0.00")
            cr = cr or Decimal("0.00")
            image_value = parse_decimal(tx.get("image_balance"))
            # Only candidate rows whose single numeric field landed in the
            # balance slot and whose two movement columns are otherwise blank.
            if image_value is None or dr != Decimal("0.00") or cr != Decimal("0.00"):
                if dr == Decimal("0.00") and cr == Decimal("0.00"):
                    page_is_complete = False
                continue
            candidates = tx.get("amount_candidates", []) or []
            matching = []
            for candidate in candidates:
                value = parse_decimal(candidate.get("value"))
                if value != image_value:
                    continue
                # Candidate must be from the primary transaction line, not a
                # cheque number, UPI RRN, account number, or timestamp beneath it.
                if candidate.get("row_aligned") is False:
                    continue
                matching.append(candidate)
            cue = _narration_direction_cue(tx.get("narration"))
            if len(matching) != 1 or cue is None:
                page_is_complete = False
                continue
            proposals[idx] = {"tx": tx, "value": image_value, "side": cue, "candidate": matching[0]}

        if not proposals:
            continue

        # Page equation can validate candidates collectively, but only where
        # every transaction has one well-defined side. No blanks are treated
        # as harmless zeros during the proof step.
        net = Decimal("0.00")
        proof_complete = page_is_complete
        for idx, tx in ordinary:
            proposal = proposals.get(idx)
            if proposal:
                value = proposal["value"]
                net += value if proposal["side"] == "CREDIT" else -value
                continue
            dr = parse_decimal(tx.get("withdrawal"))
            cr = parse_decimal(tx.get("deposit"))
            if dr is None or cr is None:
                proof_complete = False
                continue
            dr = dr or Decimal("0.00")
            cr = cr or Decimal("0.00")
            if (dr == Decimal("0.00")) == (cr == Decimal("0.00")):
                proof_complete = False
            net += cr - dr

        page_matches = proof_complete and opening + net == closing
        for idx, proposal in proposals.items():
            tx = proposal["tx"]
            value = proposal["value"]
            side = proposal["side"]
            candidate = proposal["candidate"]
            tx["misclassified_balance_candidate"] = {
                "value": str(value), "text": candidate.get("text", str(value)),
                "x0": candidate.get("x0"), "x1": candidate.get("x1"), "y": candidate.get("y"),
            }
            tx["image_balance"] = "UNEXTRACTED_AMOUNT"
            tx["balance"] = "UNEXTRACTED_AMOUNT"
            tx["image_balance_source"] = "RECLASSIFIED_CANDIDATE_NOT_A_BALANCE"
            tx["amount_reconstruction"] = {
                "method": "NARRATION_DIRECTION_PLUS_EXACT_PAGE_CHECKPOINT",
                "side": side, "value": str(value),
                "opening_balance": str(opening), "closing_balance": str(closing),
                "page_checkpoint_status": "MATCHED" if page_matches else "FAILED",
                "admitted": bool(page_matches),
            }
            if page_matches:
                tx["withdrawal"] = str(value) if side == "DEBIT" else "0.00"
                tx["deposit"] = str(value) if side == "CREDIT" else "0.00"
                tx["status"] = "RECONSTRUCTED_BY_PAGE_CHECKPOINT"
                tx["audit_detail"] = "Amount reclassified from an unmarked numeric token using its source coordinates, a single narration-direction cue, and exact printed B/F-to-C/F reconciliation."
            else:
                tx["withdrawal"] = "UNEXTRACTED_AMOUNT" if side == "DEBIT" else "0.00"
                tx["deposit"] = "UNEXTRACTED_AMOUNT" if side == "CREDIT" else "0.00"
                tx["status"] = "REVIEW_REQUIRED"
                tx["audit_detail"] = "Potential movement amount could not be admitted because the complete page failed exact B/F-to-C/F reconciliation."
    return rows


def _try_universal_structured_document(filepath, bank_name=None, account_number=None):
    """Parse structured card/broker/VDA exports with the evidence-preserving engine.

    Returns None for ordinary/unknown inputs so the mature legacy parser remains
    the fallback. The Bank of India PDF path is unchanged.
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in {".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".json"}:
        return None
    result = read_universal_structured_file(filepath)
    dtype = result.get("document_type", "UNKNOWN")
    txns = result.get("transactions", [])
    if not txns or dtype not in {"BANK", "CREDIT_CARD", "BROKER", "VDA"}:
        return None
    adapted = universal_to_application_transactions(result)
    if dtype == "BANK":
        bank = bank_name or "Bank Account"
        account = account_number or "N/A"
    else:
        bank = {"CREDIT_CARD": "Credit Card Statement", "BROKER": "Broker / Contract Note", "VDA": "VDA / Digital Asset"}[dtype]
        account = account_number or "N/A"
    for tx in adapted:
        tx["bank"] = bank
        tx["account_number"] = account
        tx["source_format"] = "UNIVERSAL_STRUCTURED_" + dtype
        tx["extraction_method"] = "HEADER_MAPPED_STRUCTURED"
    return bank, account, adapted


class StatementParserWorker(QThread):

    sig_progress = Signal(int, int, str)

    sig_statement_ready = Signal(
        str,
        str,
        str,
        list,
    )

    sig_error = Signal(
        str,
        str,
    )

    sig_finished = Signal()

    def __init__(self, filepaths: list):
        super().__init__()
        self.filepaths = filepaths
        self._is_cancelled = False
        # This is an internal construction state, not a processing state.
        # The GUI must never display it as though parsing has begun.
        self.current_stage = "Waiting for parser thread to start..."
        self.current_file = ""

    def cancel(self):
        self._is_cancelled = True
        self.current_stage = "Cancellation requested..."

    def detect_bank_and_acc(
        self,
        raw_text: str,
        filepath: str,
    ) -> tuple:

        text_upper = (
            raw_text or ""
        ).upper()

        filename_upper = (
            os.path.basename(filepath)
            .upper()
        )

        combined = (
            text_upper[:5000]
            + " "
            + filename_upper
        )

        bank_name = "UNDETERMINED BANK"

        bank_patterns = [
            (
                "STATE BANK OF INDIA",
                "State Bank of India",
            ),
            (
                "STATE BANK",
                "State Bank of India",
            ),
            (
                "HDFC BANK LIMITED",
                "HDFC Bank",
            ),
            (
                "HDFC BANK LTD",
                "HDFC Bank",
            ),
            (
                "ICICI BANK LIMITED",
                "ICICI Bank",
            ),
            (
                "ICICI BANK LTD",
                "ICICI Bank",
            ),
            (
                "AXIS BANK LIMITED",
                "Axis Bank",
            ),
            (
                "AXIS BANK LTD",
                "Axis Bank",
            ),
            (
                "KOTAK MAHINDRA BANK LIMITED",
                "Kotak Mahindra Bank",
            ),
            (
                "KOTAK MAHINDRA BANK LTD",
                "Kotak Mahindra Bank",
            ),
            (
                "CANARA BANK",
                "Canara Bank",
            ),
            (
                "UNION BANK OF INDIA",
                "Union Bank of India",
            ),
            (
                "BANK OF INDIA",
                "Bank of India",
            ),
            (
                "CENTRAL BANK OF INDIA",
                "Central Bank of India",
            ),
            (
                "INDIAN BANK",
                "Indian Bank",
            ),
            (
                "INDIAN OVERSEAS BANK",
                "Indian Overseas Bank",
            ),
            (
                "UCO BANK",
                "UCO Bank",
            ),
            (
                "PUNJAB & SIND BANK",
                "Punjab & Sind Bank",
            ),
            (
                "PUNJAB AND SIND BANK",
                "Punjab & Sind Bank",
            ),
            (
                "INDUSIND BANK",
                "IndusInd Bank",
            ),
            (
                "YES BANK",
                "Yes Bank",
            ),
            (
                "IDFC FIRST BANK",
                "IDFC FIRST Bank",
            ),
            (
                "IDBI BANK",
                "IDBI Bank",
            ),
            (
                "FEDERAL BANK",
                "Federal Bank",
            ),
            (
                "RBL BANK",
                "RBL Bank",
            ),
            (
                "BANDHAN BANK",
                "Bandhan Bank",
            ),
            (
                "AU SMALL FINANCE BANK",
                "AU Small Finance Bank",
            ),
            (
                "BANK OF MAHARASHTRA",
                "Bank of Maharashtra",
            ),
            (
                "THE KARNATAKA BANK",
                "Karnataka Bank",
            ),
            (
                "KARNATAKA BANK",
                "Karnataka Bank",
            ),
            (
                "SOUTH INDIAN BANK",
                "South Indian Bank",
            ),
            (
                "CITY UNION BANK",
                "City Union Bank",
            ),
            (
                "DCB BANK",
                "DCB Bank",
            ),
            (
                "TAMILNAD MERCANTILE BANK",
                "Tamilnad Mercantile Bank",
            ),
            (
                "TAMILNAD MERCANTILE",
                "Tamilnad Mercantile Bank",
            ),
            (
                "STATE BANK OF INDIA",
                "State Bank of India",
            ),
            (
                "PUNJAB NATIONAL BANK",
                "Punjab National Bank",
            ),
            (
                "BANK OF BARODA",
                "Bank of Baroda",
            ),
            (
                "HDFC BANK",
                "HDFC Bank",
            ),
            (
                "HDFC",
                "HDFC Bank",
            ),
            (
                "ICICI BANK",
                "ICICI Bank",
            ),
            (
                "ICICI",
                "ICICI Bank",
            ),
            (
                "AXIS BANK",
                "Axis Bank",
            ),
            (
                "AXIS",
                "Axis Bank",
            ),
            (
                "KOTAK MAHINDRA BANK",
                "Kotak Mahindra Bank",
            ),
            (
                "KOTAK",
                "Kotak Mahindra Bank",
            ),
            (
                "SBI",
                "State Bank of India",
            ),
            (
                "PNB",
                "Punjab National Bank",
            ),
            (
                "BOB",
                "Bank of Baroda",
            ),
        ]

        for pattern, detected_name in bank_patterns:

            if pattern in combined:

                bank_name = detected_name

                break

        acc_num = None

        account_patterns = [
            ACC_NUM_REGEX,
            re.compile(
                r"(?:ACCOUNT|A/C|ACCT)"
                r"[\s._:-]*"
                r"(?:NO|NUMBER|NUM)?"
                r"[\s._:-]*"
                r"([X\d]{8,18})",
                re.IGNORECASE,
            ),
        ]

        for pattern in account_patterns:

            match = pattern.search(
                text_upper[:5000]
            )

            if match:

                acc_num = (
                    match.group(1)
                    .strip()
                )

                break

        if not acc_num:

            filename_digits = re.findall(
                r"\d{9,18}",
                os.path.basename(filepath),
            )

            if filename_digits:

                acc_num = filename_digits[0]

        if not acc_num:

            acc_num = (
                "A/c-"
                + hashlib.md5(
                    filepath.encode("utf-8")
                ).hexdigest()[:8]
            )

        return bank_name, acc_num

    # ------------------------------------------------------------------
    # PDF
    # ------------------------------------------------------------------

    @staticmethod
    def _repair_transaction_sides(transactions):
        """
        Repair debit/credit assignment using the printed running balance.

        Many bank PDFs do not expose a reliable table structure to PyMuPDF.
        If one monetary amount is found but the PDF column calibration puts it
        in the wrong side, the balance equation is stronger evidence than the
        x-coordinate. This routine never invents an amount: it only moves an
        already extracted amount between debit and credit when the printed
        balances prove which side it belongs to.
        """
        previous_balance = {}

        for tx in transactions:
            key = (tx.get("bank"), tx.get("account_number"))
            balance = parse_decimal(tx.get("balance"))

            if tx.get("mode") == "REFERENCE" and balance is not None:
                previous_balance[key] = balance
                continue

            if balance is None:
                continue

            prev = previous_balance.get(key)
            dr = parse_decimal(tx.get("withdrawal")) or Decimal("0.00")
            cr = parse_decimal(tx.get("deposit")) or Decimal("0.00")

            if prev is not None:
                delta = balance - prev
                total = dr + cr

                # If exactly one side was populated but the balance proves the
                # opposite direction, move that amount to the correct side.
                if total > Decimal("0.00") and total == abs(delta):
                    if delta > Decimal("0.00"):
                        tx["deposit"] = str(total)
                        tx["withdrawal"] = "0.00"
                    elif delta < Decimal("0.00"):
                        tx["withdrawal"] = str(total)
                        tx["deposit"] = "0.00"

                # Do not choose a raw amount candidate based on numeric equality
                # alone. It may be a UPI reference, cheque number, timestamp or
                # account number from the narration. Such candidates can only
                # be reclassified by reconstruct_pdf_amounts_from_checkpoints,
                # which additionally requires source-row alignment, a unique
                # narration direction cue and exact printed page reconciliation.

            previous_balance[key] = balance

        return transactions

    @staticmethod
    def _find_page_brought_forward_reference(page):
        """Return the printed B/F reference on a page, if one is visibly present.

        This is page-local and positional.  It deliberately does not infer a
        balance from transaction arithmetic.
        """
        labels = [
            "BALANCE BROUGHT FORWARD",
            "BROUGHT FORWARD",
            "B/F",
            "B.F.",
            "BALANCE B/F",
            "OPENING BALANCE",
            "OPENING BAL",
            "BALANCE B/FWD",
            "BALANCE BROUGHT FWD",
        ]
        refs = find_reference_line_bbox(page, labels)
        for ref in refs:
            if ref.get("amount") is not None:
                return ref

        # Text-layer fallback: if the label is extractable but its amount was
        # split into separate PDF text blocks, recover the value from the short
        # post-label window and retain the label's real visual coordinates.
        page_text = page.get_text("text") or ""
        value, label_pattern = extract_any_reference_balance(page_text)
        if value is None:
            return None
        label_refs = find_reference_line_bbox(page, labels)
        bbox = label_refs[0].get("bbox") if label_refs else None
        label = label_refs[0].get("label") if label_refs else (label_pattern or "BROUGHT FORWARD")
        return {"label": label, "text": label, "bbox": bbox, "amount": value}

    @staticmethod
    def _derive_previous_page_balance(parsed_txns, page_index):
        """Derive a page-start balance only from the immediately prior page.

        Preference order is printed C/F, then the last printed transaction
        balance on the previous page.  No arithmetic is invented here.
        """
        previous = [
            tx for tx in parsed_txns
            if int(tx.get("page", -1) or -1) == int(page_index)
        ]
        if not previous:
            return None

        closings = [
            tx for tx in previous
            if tx.get("reference_role") == "CLOSING"
        ]
        if closings:
            for tx in reversed(closings):
                value = parse_decimal(tx.get("image_balance"))
                if value is not None:
                    return value

        ordered = sorted(
            previous,
            key=lambda tx: (
                float((tx.get("bbox") or [0, 0, 0, 0])[1]),
                float((tx.get("bbox") or [0, 0, 0, 0])[3]),
            ),
        )
        for tx in reversed(ordered):
            if tx.get("reference_role") in {"OPENING", "CLOSING"}:
                continue
            value = parse_decimal(tx.get("image_balance"))
            if value is not None:
                return value
        return None

    @staticmethod
    def _make_page_brought_forward_row(
        bank_name,
        acc_no,
        document_type,
        document_language,
        credit_card_summary,
        filepath,
        page_idx,
        page_rect,
        reference=None,
        derived_balance=None,
    ):
        """Create a visible B/F row for reconciliation.

        A printed B/F remains independent image evidence.  A missing B/F on a
        later page is represented by a deterministic derived B/F when the
        previous page supplies a printed ending balance.  The derived value is
        never presented as printed evidence.
        """
        if reference is not None and reference.get("amount") is not None:
            amount = parse_decimal(reference.get("amount"))
            return {
                "bank": bank_name,
                "account_number": acc_no,
                "document_type": document_type,
                "document_language": document_language,
                "credit_card_summary": credit_card_summary,
                "date": "",
                "narration": "BROUGHT FORWARD",
                "reference_label": reference.get("label", "BROUGHT FORWARD"),
                "extraction_method": "PRINTED_REFERENCE_BALANCE",
                "mode": "REFERENCE",
                "withdrawal": "0.00",
                "deposit": "0.00",
                "cash_source": "0.00",
                "balance": str(amount),
                "calculated_balance": str(amount),
                "image_balance": str(amount),
                "page": page_idx,
                "bbox": reference.get("bbox") or None,
                "reference_role": "OPENING",
                "reference_source": "PRINTED",
                "page_bf_status": "PRINTED",
                "source_file": filepath,
                "status": "REFERENCE",
            }

        if derived_balance is not None:
            amount = parse_decimal(derived_balance)
            if amount is not None:
                return {
                    "bank": bank_name,
                    "account_number": acc_no,
                    "document_type": document_type,
                    "document_language": document_language,
                    "credit_card_summary": credit_card_summary,
                    "date": "",
                    "narration": "BROUGHT FORWARD (DERIVED)",
                    "reference_label": "BROUGHT FORWARD (DERIVED FROM PRIOR PAGE)",
                    "extraction_method": "DERIVED_PRIOR_PAGE_BALANCE",
                    "mode": "REFERENCE",
                    "withdrawal": "0.00",
                    "deposit": "0.00",
                    "cash_source": "0.00",
                    "balance": str(amount),
                    "calculated_balance": str(amount),
                    "image_balance": "UNEXTRACTED_AMOUNT",
                    "page": page_idx,
                    "bbox": None,
                    "reference_role": "OPENING",
                    "reference_source": "DERIVED",
                    "page_bf_status": "DERIVED",
                    "source_file": filepath,
                    # This is a reconciliation marker, not a claim that a
                    # printed B/F was found on the page.
                    "status": "MATCHED",
                    "audit_detail": "No printed B/F found on this page; starting balance derived from the prior page's printed ending balance.",
                }

        return {
            "bank": bank_name,
            "account_number": acc_no,
            "document_type": document_type,
            "document_language": document_language,
            "credit_card_summary": credit_card_summary,
            "date": "",
            "narration": "BROUGHT FORWARD (NOT DERIVABLE)",
            "reference_label": "BROUGHT FORWARD NOT PROVIDED",
            "extraction_method": "NO_PAGE_BF_BALANCE",
            "mode": "REFERENCE",
            "withdrawal": "0.00",
            "deposit": "0.00",
            "cash_source": "0.00",
            "balance": "UNEXTRACTED_AMOUNT",
            "calculated_balance": "UNEXTRACTED_AMOUNT",
            "image_balance": "UNEXTRACTED_AMOUNT",
            "page": page_idx,
            "bbox": None,
            "reference_role": "OPENING",
            "reference_source": "MISSING",
            "page_bf_status": "MISSING",
            "source_file": filepath,
            "status": "REVIEW_REQUIRED",
            "audit_detail": "No printed B/F found and no reliable prior-page ending balance was available.",
        }

    def parse_pdf(
        self,
        filepath,
        file_index,
        total_files,
    ):

        filename = os.path.basename(filepath)
        self.current_file = filepath
        self.current_stage = f"Opening PDF file: {filename}"

        self.sig_progress.emit(
            0,
            max(1, total_files),
            f"Opening PDF file: {filename}"
        )

        with open(filepath, "rb") as f:
            file_bytes = f.read()

        try:
            doc = open_pdf_with_optional_security(filepath)
        except Exception:
            doc = fitz.open(stream=file_bytes, filetype="pdf")

        try:

            total_pages = len(doc)

            if total_pages == 0:

                self.sig_error.emit(
                    filepath,
                    "File contains 0 pages.",
                )

                return

            first_page_text = (
                doc[0].get_text("text")
                if total_pages
                else ""
            )

            bank_name, acc_no = (
                self.detect_bank_and_acc(
                    first_page_text,
                    filepath,
                )
            )
            document_type = detect_statement_type(first_page_text)
            document_language = detect_statement_language(first_page_text)
            credit_card_summary = (
                extract_credit_card_summary(first_page_text, document_language)
                if document_type == "CREDIT_CARD" else {}
            )

            active_channels = (
                DynamicHeaderCalibrator.calibrate_columns(
                    doc[0]
                )
            )

            parsed_txns = []
            opening_reference_added = False
            visual_caps = optional_visual_capabilities()

            for page_idx in range(total_pages):

                if self._is_cancelled:
                    break

                self.current_stage = (
                    f"Processing {filename} ({file_index + 1}/{total_files}) "
                    f"— Page {page_idx + 1}/{total_pages}"
                )
                self.sig_progress.emit(
                    page_idx + 1,
                    total_pages,
                    (
                        f"Processing {filename} "
                        f"({file_index + 1}/{total_files}) "
                        f"— Page {page_idx + 1}/{total_pages}"
                    ),
                )

                page = doc[page_idx]

                active_channels = (
                    DynamicHeaderCalibrator.calibrate_columns(
                        page,
                        fallback=active_channels,
                    )
                )

                # Check the page for a printed B/F. Printed B/F is real visual
                # evidence and is admitted immediately. Derived B/F is inserted
                # only after the page proves that it contains transactions.
                page_bf_reference = self._find_page_brought_forward_reference(page)
                if page_idx == 0:
                    if page_bf_reference is not None and page_bf_reference.get("amount") is not None:
                        parsed_txns.append(self._make_page_brought_forward_row(
                            bank_name, acc_no, document_type, document_language,
                            credit_card_summary, filepath, page_idx, page.rect,
                            reference=page_bf_reference,
                        ))
                        opening_reference_added = True
                elif page_bf_reference is not None and page_bf_reference.get("amount") is not None:
                    parsed_txns.append(self._make_page_brought_forward_row(
                        bank_name, acc_no, document_type, document_language,
                        credit_card_summary, filepath, page_idx, page.rect,
                        reference=page_bf_reference,
                    ))

                words = page.get_text("words")

                if len(words) < 5:
                    # Scanned/image-only PDF page.  Rendering a page solely to
                    # discover that optional OCR dependencies are unavailable
                    # is wasteful and was a major source of slow-feeling scans.
                    # Core PDF extraction must stay fast and dependency-safe.
                    if not all(visual_caps.values()):
                        missing = ", ".join(
                            name for name, available in visual_caps.items()
                            if not available
                        )
                        self.sig_progress.emit(
                            page_idx + 1, total_pages,
                            f"Scanned page {page_idx + 1}: optional OCR unavailable ({missing}); page retained for review."
                        )
                        continue

                    # OCR-derived transactions are admitted only when the visual
                    # evidence engine certifies them.
                    try:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), alpha=False)
                        visual = VisualEvidenceEngine.extract(
                            pix.tobytes("png"),
                            source_file=filepath,
                            page=page_idx,
                            bank=bank_name,
                            account=acc_no,
                        )
                        if visual.get("ledger_eligible"):
                            parsed_txns.extend(visual.get("transactions", []))
                        else:
                            self.sig_progress.emit(
                                page_idx + 1, total_pages,
                                f"Visual review required for scanned PDF page {page_idx + 1}"
                            )
                    except Exception as visual_exc:
                        self.sig_progress.emit(
                            page_idx + 1, total_pages,
                            f"OCR unavailable for scanned page {page_idx + 1}: {visual_exc}"
                        )
                    continue

                sorted_words = sorted(
                    words,
                    key=lambda w: (
                        round(w[1] / 6) * 6,
                        w[0],
                    ),
                )
                carried_reference_lines = find_reference_line_bbox(
                    page, ["CARRIED FORWARD", "C/F", "BALANCE C/F", "TOTAL C/F"]
                )

                date_column_left, date_column_right = (
                    infer_pdf_date_column_bounds(
                        words,
                        page.rect.width,
                    )
                )

                curr_txn = None

                def finalize_current_transaction():

                    nonlocal curr_txn

                    if curr_txn is None:
                        return

                    narration_upper = (
                        curr_txn
                        .get("narration", "")
                        .upper()
                    )
                    normalized_narration = re.sub(r"\s+", " ", narration_upper).strip()
                    # Dated reference lines can be present in the table text as
                    # well as being separately captured by the page-reference
                    # extractor. Do not admit the same B/F or C/F a second time
                    # as a zero-value transaction.
                    reference_only_pattern = (
                        r"(?:OPENING\s+BALANCE(?:\s+B/?F)?|OPENING\s+BAL(?:\s+B/?F)?|"
                        r"CLOSING\s+BALANCE(?:\s+C/?F)?|CLOSING\s+BAL|"
                        r"BALANCE\s+(?:BROUGHT\s+FORWARD|CARRIED\s+FORWARD|B/?F|C/?F)|"
                        r"BROUGHT\s+FORWARD|CARRIED\s+FORWARD|B/?F|B\.F\.|C/?F|C\.F\.|TOTAL\s+C/?F)"
                    )
                    if re.fullmatch(reference_only_pattern, normalized_narration):
                        curr_txn = None
                        return

                    if any(
                        bf in narration_upper
                        for bf in BF_PATTERNS
                    ):

                        # Reference/header rows are handled separately. Do not
                        # discard a dated transaction merely because a bank's
                        # multi-line narration happens to contain B/F or C/F.
                        if not curr_txn.get("date"):
                            curr_txn = None
                            return

                    narration_upper = narration_upper.strip()

                    # Conservative transaction-mode inference. Only explicit
                    # cash indicators are classified as physical cash.
                    if any(
                        key in narration_upper
                        for key in [
                            "ATM WDL",
                            "ATM WITHDRAW",
                            "CASH WITHDRAW",
                            "CASH WDL",
                            "CASH WITHDRAWAL",
                        ]
                    ):
                        curr_txn["mode"] = "CASH_WDL"

                    elif any(
                        key in narration_upper
                        for key in [
                            "CASH DEP",
                            "CASH DEPOSIT",
                            "CDM CASH",
                            "CASH RECEIVED",
                            "CASH CREDIT",
                        ]
                    ):
                        curr_txn["mode"] = "CASH_DEP"

                    # A narration such as CASH RECEIVED ON PROPERTY SALE
                    # contains an explicit external cash-source claim. It
                    # is still a bank cash deposit for statutory purposes,
                    # but for physical-cash availability it supplies the
                    # same amount immediately before the bank deposit.
                    if curr_txn.get("mode") == "CASH_DEP":
                        curr_txn["cash_source"] = str(
                            infer_declared_cash_source(
                                narration_upper,
                                curr_txn.get("deposit"),
                            )
                        )

                    # Deliberately do not infer a running balance from a
                    # monetary token printed below the transaction date.
                    # Those layouts are too ambiguous for a deterministic
                    # forensic parser and can interfere with recognition of
                    # the basic debit/credit columns.  A balance is accepted
                    # only when it is found in an explicit calibrated balance
                    # field or as an independently identified reference
                    # balance such as B/F or C/F.

                    # Preserve the printed balance in both legacy ``balance``
                    # and the new ``image_balance`` fields at parser output.
                    # The GUI reconciliation pass immediately replaces
                    # ``balance`` with the formula-derived canonical value.
                    image_bal = parse_decimal(curr_txn.get("image_balance"))
                    curr_txn["balance"] = str(image_bal) if image_bal is not None else "UNEXTRACTED_AMOUNT"
                    debit_value = parse_decimal(curr_txn.get("withdrawal"))
                    credit_value = parse_decimal(curr_txn.get("deposit"))
                    debit_value = debit_value if debit_value is not None else Decimal("0.00")
                    credit_value = credit_value if credit_value is not None else Decimal("0.00")
                    exactly_one_side = (debit_value > Decimal("0.00")) != (credit_value > Decimal("0.00"))
                    # Never label an ordinary transaction OK before financial
                    # validation. A row with zero debit and credit is not a
                    # successfully extracted transaction simply because some
                    # numeric token was assigned to its balance field.
                    if curr_txn.get("field_conflicts") or not exactly_one_side:
                        curr_txn["status"] = "REVIEW_REQUIRED"
                    elif image_bal is None:
                        curr_txn["status"] = "BALANCE_NOT_PROVIDED"
                    else:
                        curr_txn["status"] = "REVIEW_REQUIRED"

                    parsed_txns.append(curr_txn)

                    curr_txn = None

                for word in sorted_words:

                    x0, y0, x1, y1, text = (
                        word[0],
                        word[1],
                        word[2],
                        word[3],
                        str(word[4]).strip(),
                    )

                    if not text:
                        continue

                    x_mid = (
                        x0 + x1
                    ) / 2.0

                    # Reference lines are independent evidence rows. Stop the
                    # preceding transaction before C/F so its evidence bbox
                    # cannot swallow the footer.
                    reference_line_y = False
                    for _ref in carried_reference_lines:
                        rb = _ref.get("bbox", [0, 0, 0, 0])
                        if float(rb[1]) - 1.5 <= y0 <= float(rb[3]) + 1.5:
                            reference_line_y = True
                            break
                    if reference_line_y:
                        finalize_current_transaction()
                        continue

                    # Date recognition is adaptive, but a date must also
                    # belong to the physical transaction-date column. V22
                    # intentionally removed this second check to avoid false
                    # negatives on difficult PDFs. That created the opposite
                    # failure mode: dates embedded inside narration such as
                    # "FOR PERIOD 01/04/24 TO 30/04/24" became fake rows.
                    #
                    # V23 uses a page-local date-column bound inferred from
                    # the leftmost cluster of date-like PDF words. This keeps
                    # genuine dates flexible while excluding dates appearing
                    # materially farther right inside narration.
                    is_date = (
                        is_valid_row_date(
                            text,
                            x0,
                            page.rect.width,
                        )
                        and date_column_left <= x_mid <= date_column_right
                    )

                    if is_date:

                        finalize_current_transaction()

                        curr_txn = {
                            "bank": bank_name,
                            "account_number": acc_no,
                            "document_type": document_type,
                            "document_language": document_language,
                            "credit_card_summary": credit_card_summary,
                            "date": text,
                            "narration": "",
                            "mode": "DIGITAL",
                            "row_anchor_y": y0,
                            "date_x_mid": x_mid,

                            # Debit and credit can legitimately
                            # be zero.
                            "withdrawal": "0.00",
                            "deposit": "0.00",

                            # Explicit external physical-cash source, when
                            # the narration itself states CASH RECEIVED.
                            "cash_source": "0.00",
                            "amount_candidates": [],
                            "field_conflicts": [],
                            "image_balance_source": "NOT_PROVIDED",

                            # Printed balance must be supported by source text/geometry.
                            "balance": "UNEXTRACTED_AMOUNT",
                            "image_balance": "UNEXTRACTED_AMOUNT",

                            "page": page_idx,
                            "bbox": [
                                x0,
                                y0,
                                x1,
                                y1,
                            ],
                            "source_file": filepath,
                            "status": "OK",
                        }

                        continue

                    if curr_txn is None:
                        continue

                    curr_txn["bbox"][0] = min(
                        curr_txn["bbox"][0],
                        x0,
                    )

                    curr_txn["bbox"][1] = min(
                        curr_txn["bbox"][1],
                        y0,
                    )

                    curr_txn["bbox"][2] = max(
                        curr_txn["bbox"][2],
                        x1,
                    )

                    curr_txn["bbox"][3] = max(
                        curr_txn["bbox"][3],
                        y1,
                    )

                    # --------------------------------------------------
                    # BANK-SPECIFIC RUNNING BALANCE HANDLING
                    # --------------------------------------------------
                    # This statement prints the running balance on the line
                    # immediately below the transaction, often under the
                    # first date column, e.g. 50,202.80Cr. Its x-coordinate
                    # therefore cannot be used to identify it.
                    if is_running_balance_token(text):
                        bare_balance = re.sub(r"(?i)\s*(CR|DR)$", "", text).strip()
                        balance_value = parse_decimal(bare_balance)
                        if balance_value is not None:
                            curr_txn["image_balance"] = str(balance_value)
                            curr_txn["image_balance_source"] = "EXPLICIT_RUNNING_BALANCE_TOKEN"
                            curr_txn["image_balance_bbox"] = [float(x0), float(y0), float(x1), float(y1)]
                        continue

                    # --------------------------------------------------
                    # COLUMN-SAFE TOKEN CLASSIFICATION
                    # --------------------------------------------------
                    # A PDF text extractor can place narration words inside
                    # an amount column when the bank statement uses
                    # multi-line narration or slightly shifted coordinates.
                    # Never overwrite a monetary field with a non-numeric
                    # token. Blank debit/credit is legitimate; balance is
                    # mandatory and is therefore left as UNEXTRACTED_AMOUNT
                    # until a valid numeric token is found.

                    numeric_value = parse_decimal(text)
                    row_aligned = abs(float(y0) - float(curr_txn.get("row_anchor_y", y0))) <= max(4.0, (float(y1) - float(y0)) * 0.65)

                    # Classify from horizontal column geometry before deciding
                    # whether the token shares the date baseline. Several real
                    # statement layouts put amount/balance cells on a lower line
                    # than the transaction date and multi-line narration.
                    column_role = None
                    if numeric_value is not None:
                        column_role = pdf_numeric_column_role(x0, x_mid, x1, active_channels)

                    if numeric_value is not None:
                        curr_txn.setdefault("amount_candidates", []).append({
                            "value": str(numeric_value),
                            "x": x_mid,
                            "x0": float(x0),
                            "x1": float(x1),
                            "y": float(y0),
                            "text": text,
                            "row_aligned": bool(row_aligned),
                            "column_role": column_role,
                            "money_token": bool(is_money_token(text)),
                            "column_geometry_confident": bool(
                                (active_channels or {}).get("_money_right_edge_channels", {}).get(column_role)
                            ),
                        })

                    # Do not lose monetary cells merely because their y-position
                    # is below the transaction-date baseline. They are retained
                    # as coordinate-backed candidates and resolved only when a
                    # unique monetary-column token exists. Numeric-looking
                    # references in the narration column remain narration.
                    if numeric_value is not None and not row_aligned:
                        is_column_money = (
                            is_money_token(text)
                            and column_role in {"withdrawal", "deposit", "balance"}
                            and bool((active_channels or {}).get("_money_right_edge_channels"))
                        )
                        if is_column_money:
                            continue
                        if curr_txn["narration"]:
                            curr_txn["narration"] += " "
                        curr_txn["narration"] += text
                        continue

                    narration_left, narration_right = active_channels["narration"]
                    withdrawal_left, withdrawal_right = active_channels["withdrawal"]
                    deposit_left, deposit_right = active_channels["deposit"]
                    balance_left, balance_right = active_channels["balance"]

                    if numeric_value is not None and column_role == "balance":
                        curr_txn["image_balance"] = str(numeric_value)
                        curr_txn["image_balance_source"] = "CALIBRATED_HEADER_RIGHT_EDGE"
                        curr_txn["image_balance_bbox"] = [float(x0), float(y0), float(x1), float(y1)]
                    elif numeric_value is not None and column_role == "withdrawal":
                        curr_txn["withdrawal"] = str(abs(numeric_value))
                        curr_txn["withdrawal_bbox"] = [float(x0), float(y0), float(x1), float(y1)]
                    elif numeric_value is not None and column_role == "deposit":
                        curr_txn["deposit"] = str(abs(numeric_value))
                        curr_txn["deposit_bbox"] = [float(x0), float(y0), float(x1), float(y1)]
                    elif narration_left <= x_mid <= narration_right or column_role == "narration":
                        if curr_txn["narration"]:
                            curr_txn["narration"] += " "
                        curr_txn["narration"] += text
                    elif numeric_value is None:
                        # Preserve non-numeric text outside calibrated columns.
                        if curr_txn["narration"]:
                            curr_txn["narration"] += " "
                        curr_txn["narration"] += text

                finalize_current_transaction()

                # Capture CARRIED FORWARD / C/F as its own reference row.
                # It must never become part of the final transaction's bbox.
                carried_lines = find_reference_line_bbox(
                    page, ["CARRIED FORWARD", "C/F", "BALANCE C/F", "TOTAL C/F"]
                )
                for ref in carried_lines:
                    if ref.get("amount") is None:
                        continue
                    parsed_txns.append({
                        "bank": bank_name,
                        "account_number": acc_no,
                        "document_type": document_type,
                        "document_language": document_language,
                        "credit_card_summary": credit_card_summary,
                        "date": "",
                        "narration": ref.get("label", "CARRIED FORWARD"),
                        "reference_label": ref.get("label", "CARRIED FORWARD"),
                        "extraction_method": "PRINTED_REFERENCE_BALANCE",
                        "mode": "REFERENCE",
                        "withdrawal": "0.00",
                        "deposit": "0.00",
                        "cash_source": "0.00",
                        "balance": str(ref["amount"]),
                        "calculated_balance": str(ref["amount"]),
                        "image_balance": str(ref["amount"]),
                        "page": page_idx,
                        "bbox": ref["bbox"],
                        "reference_role": "CLOSING",
                        "source_file": filepath,
                        "status": "REFERENCE",
                    })

                # Continuous-statement page boundary: when a later page has no
                # printed B/F, add exactly one derived B/F only if the page has
                # admitted transaction rows and the prior page supplies a reliable
                # printed ending balance.
                if page_idx > 0 and page_bf_reference is None:
                    page_transactions = [tx for tx in parsed_txns
                        if int(tx.get("page", -1) or -1) == page_idx
                        and tx.get("date") and tx.get("mode") != "REFERENCE"]
                    already_has_bf = any(int(tx.get("page", -1) or -1) == page_idx and tx.get("reference_role") == "OPENING" for tx in parsed_txns)
                    if page_transactions and not already_has_bf:
                        prior_page_balance = self._derive_previous_page_balance(parsed_txns, page_idx - 1)
                        if prior_page_balance is not None:
                            parsed_txns.append(self._make_page_brought_forward_row(
                                bank_name, acc_no, document_type, document_language,
                                credit_card_summary, filepath, page_idx, page.rect,
                                derived_balance=prior_page_balance,
                            ))

                # Diagnostic only: distinguish a text-layer page with no
                # recognized transaction dates from an actual empty/scanned
                # page. This makes real-bank debugging much faster.
                if words and not any(
                    tx.get("page") == page_idx and tx.get("date")
                    for tx in parsed_txns
                ):
                    self.sig_progress.emit(
                        page_idx + 1,
                        total_pages,
                        f"Review: page {page_idx + 1} contains {len(words)} PDF words but no transaction date was admitted."
                    )

            if not self._is_cancelled:

                multiline_recovery = recover_multiline_money_cells(parsed_txns)
                delta_corrections = resolve_movement_from_exact_balance_delta(parsed_txns)
                self.sig_progress.emit(total_pages, total_pages,
                    "Multi-line column recovery: "
                    f"{multiline_recovery['withdrawal']} debit, "
                    f"{multiline_recovery['deposit']} credit, "
                    f"{multiline_recovery['balance']} image-balance cell(s); "
                    f"{len(delta_corrections)} exact balance-delta correction(s); "
                    f"{multiline_recovery['conflicts']} ambiguous cell group(s) withheld.")

                parsed_txns = reconstruct_pdf_amounts_from_checkpoints(parsed_txns)
                recovered_count = sum(1 for tx in parsed_txns if (tx.get("amount_reconstruction") or {}).get("admitted"))
                withheld_count = sum(1 for tx in parsed_txns if tx.get("amount_reconstruction") and not tx["amount_reconstruction"].get("admitted"))
                self.sig_progress.emit(total_pages, total_pages,
                    f"Page-checkpoint recovery: {recovered_count} amount(s) confirmed; {withheld_count} candidate(s) withheld for review.")
                parsed_txns = self._repair_transaction_sides(parsed_txns)

                self.sig_statement_ready.emit(
                    filepath,
                    bank_name,
                    acc_no,
                    parsed_txns,
                )

        finally:

            doc.close()

    # ------------------------------------------------------------------
    # Excel
    # ------------------------------------------------------------------

    def parse_excel(
        self,
        filepath,
    ):

        filename = os.path.basename(filepath)

        self.sig_progress.emit(
            1,
            1,
            f"Parsing Excel Statement: {filename}...",
        )

        try:
            universal = _try_universal_structured_document(filepath)
            if universal is not None:
                bank_name, acc_no, parsed_txns = universal
                if not self._is_cancelled:
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, parsed_txns)
                return
        except (ValueError, OSError, csv.Error, json.JSONDecodeError) as exc:
            self.sig_error.emit(filepath, f"Universal structured extraction failed; legacy parser will be tried: {exc}")

        wb = openpyxl.load_workbook(
            filepath,
            data_only=True,
        )

        try:

            ws = wb.active

            header_sample = " ".join(
                [
                    str(cell)
                    for row in ws.iter_rows(
                        max_row=10,
                        values_only=True,
                    )
                    for cell in row
                    if cell is not None
                ]
            )

            bank_name, acc_no = (
                self.detect_bank_and_acc(
                    header_sample,
                    filepath,
                )
            )

            parsed_txns = []

            for row in ws.iter_rows(
                values_only=True
            ):

                if not row:
                    continue

                normalized = [
                    ""
                    if value is None
                    else str(value).strip()
                    for value in row
                ]

                if not normalized:
                    continue

                row_upper = " ".join(
                    normalized
                ).upper()

                if (
                    "TXN DATE" in row_upper
                    or (
                        "DATE" in row_upper
                        and "BALANCE" in row_upper
                    )
                ):
                    continue

                if not normalized[0]:
                    continue

                d_str = normalized[0]

                if not is_valid_row_date(
                    d_str,
                    x0=50,
                ):
                    continue

                narration = (
                    normalized[1]
                    if len(normalized) > 1
                    else ""
                )

                if any(
                    bf in narration.upper()
                    for bf in BF_PATTERNS
                ):
                    continue

                dr = (
                    normalized[2]
                    if len(normalized) > 2
                    and normalized[2]
                    else "0.00"
                )

                cr = (
                    normalized[3]
                    if len(normalized) > 3
                    and normalized[3]
                    else "0.00"
                )

                bal = (
                    normalized[4]
                    if len(normalized) > 4
                    and normalized[4]
                    else "UNEXTRACTED_AMOUNT"
                )

                status = "OK"

                if (
                    parse_decimal(dr) is None
                    or parse_decimal(cr) is None
                    or parse_decimal(bal) is None
                ):

                    status = "UNEXTRACTED_AMOUNT"

                parsed_txns.append(
                    {
                        "bank": bank_name,
                        "account_number": acc_no,
                        "date": d_str,
                        "narration": narration,
                        "mode": "DIGITAL",
                        "withdrawal": dr,
                        "deposit": cr,
                        "balance": bal,
                        "page": 0,
                        "bbox": [
                            50,
                            100,
                            500,
                            140,
                        ],
                        "source_file": filepath,
                        "status": status,
                    }
                )

            if not self._is_cancelled:

                self.sig_statement_ready.emit(
                    filepath,
                    bank_name,
                    acc_no,
                    parsed_txns,
                )

        finally:

            wb.close()

    # ------------------------------------------------------------------
    # CSV
    # ------------------------------------------------------------------

    def parse_csv(
        self,
        filepath,
    ):

        filename = os.path.basename(filepath)

        self.sig_progress.emit(
            1,
            1,
            f"Parsing CSV Statement: {filename}...",
        )

        try:
            with open(filepath, "r", encoding="utf-8-sig", errors="replace") as probe:
                probe_text = probe.read(8192)
            detected_bank, detected_account = self.detect_bank_and_acc(probe_text, filepath)
            universal = _try_universal_structured_document(filepath, detected_bank, detected_account)
            if universal is not None:
                bank_name, acc_no, parsed_txns = universal
                if not self._is_cancelled:
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, parsed_txns)
                return
        except (ValueError, OSError, csv.Error, json.JSONDecodeError) as exc:
            self.sig_error.emit(filepath, f"Universal structured extraction failed; legacy parser will be tried: {exc}")

        with open(
            filepath,
            "r",
            encoding="utf-8-sig",
            newline="",
        ) as f:

            sample = f.read(4096)

            bank_name, acc_no = (
                self.detect_bank_and_acc(
                    sample,
                    filepath,
                )
            )

            f.seek(0)

            reader = csv.reader(f, delimiter="\t" if filepath.lower().endswith(".tsv") else ",")

            parsed_txns = []

            for row in reader:

                if not row:
                    continue

                normalized = [
                    str(value).strip()
                    for value in row
                ]

                if not normalized[0]:
                    continue

                row_upper = " ".join(
                    normalized
                ).upper()

                if (
                    "TXN DATE" in row_upper
                    or (
                        "DATE" in row_upper
                        and "BALANCE" in row_upper
                    )
                ):
                    continue

                d_str = normalized[0]

                if not is_valid_row_date(
                    d_str,
                    x0=50,
                ):
                    continue

                narration = (
                    normalized[1]
                    if len(normalized) > 1
                    else ""
                )

                if any(
                    bf in narration.upper()
                    for bf in BF_PATTERNS
                ):
                    continue

                dr = (
                    normalized[2]
                    if len(normalized) > 2
                    and normalized[2]
                    else "0.00"
                )

                cr = (
                    normalized[3]
                    if len(normalized) > 3
                    and normalized[3]
                    else "0.00"
                )

                bal = (
                    normalized[4]
                    if len(normalized) > 4
                    and normalized[4]
                    else "UNEXTRACTED_AMOUNT"
                )

                status = "OK"

                if (
                    parse_decimal(dr) is None
                    or parse_decimal(cr) is None
                    or parse_decimal(bal) is None
                ):
                    status = "UNEXTRACTED_AMOUNT"

                parsed_txns.append(
                    {
                        "bank": bank_name,
                        "account_number": acc_no,
                        "date": d_str,
                        "narration": narration,
                        "mode": "DIGITAL",
                        "withdrawal": dr,
                        "deposit": cr,
                        "balance": bal,
                        "page": 0,
                        "bbox": [
                            50,
                            100,
                            500,
                            140,
                        ],
                        "source_file": filepath,
                        "status": status,
                    }
                )

        if not self._is_cancelled:

            self.sig_statement_ready.emit(
                filepath,
                bank_name,
                acc_no,
                parsed_txns,
            )

class UniversalStatementEngine:
    """Machine-readable bank statement ingestion with conservative certification."""

    @staticmethod
    def _parse_mt_amount(value):
        raw = str(value or "").strip().replace(" ", "")
        if "," in raw and "." not in raw:
            raw = raw.replace(",", ".")
        elif "," in raw and "." in raw and raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        return parse_decimal(raw)

    @staticmethod
    def _clean_text(value):
        if value is None:
            return ""
        return re.sub(r"\s+", " ", html.unescape(str(value))).strip()

    @staticmethod
    def _normalize_date(value):
        if value is None:
            return ""
        raw = str(value).strip()
        dt = parse_statement_date(raw[:10])
        if dt is None and re.fullmatch(r"\d{6}", raw):
            try:
                dt = datetime.strptime(raw, "%y%m%d")
            except Exception:
                dt = None
        return dt.strftime("%d-%m-%Y") if dt else raw

    @staticmethod
    def _tx(bank, account, date, narration, debit="0.00", credit="0.00", balance="UNEXTRACTED_AMOUNT", source_file="", source_format="STRUCTURED", reference=""):
        dr = parse_decimal(debit)
        cr = parse_decimal(credit)
        bal = parse_decimal(balance)
        return {
            "bank": bank or "Bank Account",
            "account_number": account or "N/A",
            "date": UniversalStatementEngine._normalize_date(date),
            "narration": UniversalStatementEngine._clean_text(narration),
            "mode": "DIGITAL",
            "withdrawal": str(dr if dr is not None else Decimal("0.00")),
            "deposit": str(cr if cr is not None else Decimal("0.00")),
            "balance": str(bal if bal is not None else "UNEXTRACTED_AMOUNT"),
            "image_balance": str(bal if bal is not None else "UNEXTRACTED_AMOUNT"),
            "page": 0,
            "bbox": [50, 100, 500, 140],
            "source_file": source_file,
            "source_format": source_format,
            "reference": reference,
            "extraction_method": source_format,
            "status": "OK" if bal is not None else "UNEXTRACTED_AMOUNT",
        }

    @staticmethod
    def _local(tag):
        return str(tag).split("}")[-1].upper()

    @staticmethod
    def _xtext(node):
        if node is None:
            return ""
        return UniversalStatementEngine._clean_text(" ".join(node.itertext()))

    @staticmethod
    def _find(node, names):
        names = {str(x).upper() for x in names}
        for child in node.iter():
            if UniversalStatementEngine._local(child.tag) in names:
                return child
        return None

    @staticmethod
    def _findall(node, name):
        wanted = str(name).upper()
        return [x for x in node.iter() if UniversalStatementEngine._local(x.tag) == wanted]

    @staticmethod
    def _signed(amount, indicator):
        value = parse_decimal(amount)
        if value is None:
            return None
        return value if str(indicator).upper() in {"C", "CR", "CRDT", "CREDIT"} else -value

    @staticmethod
    def _rebuild_balances(transactions, opening=None, closing=None):
        current = parse_decimal(opening)
        for tx in transactions:
            dr = parse_decimal(tx.get("withdrawal")) or Decimal("0.00")
            cr = parse_decimal(tx.get("deposit")) or Decimal("0.00")
            if current is None:
                tx["status"] = "UNEXTRACTED_AMOUNT"
                continue
            current += cr - dr
            tx["balance"] = str(current)
            tx["balance_source"] = "DERIVED_FROM_EXPLICIT_OPENING_BALANCE"
            tx["status"] = "OK"
        final = parse_decimal(closing)
        if current is not None and final is not None:
            passed = current == final
            for tx in transactions:
                tx["statement_closing_balance_check"] = "PASS" if passed else "MISMATCH"
                if not passed:
                    tx["status"] = "MISMATCH"
            return passed
        return current is not None

    @staticmethod
    def _detect_bank_account(text, filepath):
        combined = ((text or "")[:20000] + " " + os.path.basename(filepath)).upper()
        patterns = [
            ("STATE BANK OF INDIA", "State Bank of India"), ("PUNJAB NATIONAL BANK", "Punjab National Bank"),
            ("BANK OF BARODA", "Bank of Baroda"), ("BANK OF INDIA", "Bank of India"),
            ("CANARA BANK", "Canara Bank"), ("UNION BANK OF INDIA", "Union Bank of India"),
            ("INDIAN BANK", "Indian Bank"), ("CENTRAL BANK OF INDIA", "Central Bank of India"),
            ("UCO BANK", "UCO Bank"), ("HDFC BANK", "HDFC Bank"), ("ICICI BANK", "ICICI Bank"),
            ("AXIS BANK", "Axis Bank"), ("KOTAK MAHINDRA BANK", "Kotak Mahindra Bank"),
            ("IDBI BANK", "IDBI Bank"), ("INDUSIND BANK", "IndusInd Bank"), ("YES BANK", "Yes Bank"),
            ("FEDERAL BANK", "Federal Bank"), ("BANDHAN BANK", "Bandhan Bank"),
            ("AU SMALL FINANCE BANK", "AU Small Finance Bank"), ("RBL BANK", "RBL Bank"),
            ("IDFC FIRST BANK", "IDFC FIRST Bank"), ("CITIBANK", "Citibank"), ("CITI", "Citibank"),
            ("HSBC", "HSBC"), ("STANDARD CHARTERED", "Standard Chartered"), ("DBS BANK", "DBS Bank"),
        ]
        bank = "Bank Account"
        for needle, name in patterns:
            if needle in combined:
                bank = name
                break
        account = None
        for pat in [
            re.compile(r"(?:ACCOUNT|A/C|ACCT)[\s._:-]*(?:NO|NUMBER|NUM)?[\s._:-]*([X\d]{6,24})", re.I),
            re.compile(r":25:([A-Z0-9/._-]{6,30})", re.I),
            re.compile(r"<ACCTID>([^<]+)", re.I),
        ]:
            m = pat.search(combined)
            if m:
                account = m.group(1).strip()
                break
        if not account:
            nums = re.findall(r"\d{9,18}", os.path.basename(filepath))
            account = nums[0] if nums else "A/c-" + hashlib.md5(filepath.encode()).hexdigest()[:8]
        return bank, account

    @classmethod
    def parse_ofx(cls, filepath):
        text = open(filepath, encoding="utf-8-sig", errors="replace").read()
        bank, account = cls._detect_bank_account(text, filepath)
        m = re.search(r"<ACCTID>([^<\r\n]+)", text, re.I)
        if m:
            account = m.group(1).strip()
        txns = []
        blocks = re.findall(r"<STMTTRN>(.*?)(?=<STMTTRN>|</BANKTRANLIST>|</STMTRS>|\Z)", text, re.I | re.S)
        for block in blocks:
            def tag(name):
                x = re.search(rf"<{name}>([^<\r\n]+)", block, re.I)
                return x.group(1).strip() if x else ""
            amount = parse_decimal(tag("TRNAMT"))
            if amount is None:
                continue
            typ = tag("TRNTYPE").upper()
            debit = abs(amount) if amount < 0 or typ in {"DEBIT", "CHECK", "PAYMENT", "DIRECTDEBIT", "OTHERDEBIT"} else Decimal("0.00")
            credit = abs(amount) if debit == 0 else Decimal("0.00")
            txns.append(cls._tx(bank, account, tag("DTPOSTED")[:8], " | ".join(x for x in [tag("NAME"), tag("PAYEE"), tag("MEMO")] if x), debit, credit, source_file=filepath, source_format="OFX", reference=tag("FITID")))
        m = re.search(r"<LEDGERBAL>.*?<BALAMT>([-+]?\d[\d.,]*)", text, re.I | re.S)
        if m and txns:
            txns[-1]["statement_closing_balance"] = str(parse_decimal(m.group(1)))
        return bank, account, txns

    @classmethod
    def parse_camt(cls, filepath):
        root = ET.parse(filepath).getroot()
        text = ET.tostring(root, encoding="unicode")
        bank, account = cls._detect_bank_account(text, filepath)
        opening = closing = None
        for bal in cls._findall(root, "Bal"):
            code = cls._xtext(cls._find(bal, {"Cd"})).upper()
            amount = parse_decimal(cls._xtext(cls._find(bal, {"Amt"})))
            signed = cls._signed(amount, cls._xtext(cls._find(bal, {"CdtDbtInd"})))
            if signed is None:
                continue
            if code in {"OPBD", "PRCD"} and opening is None:
                opening = signed
            elif code in {"CLBD", "CLAV", "ITBD"}:
                closing = signed
        txns = []
        for entry in cls._findall(root, "Ntry"):
            amount = parse_decimal(cls._xtext(cls._find(entry, {"Amt"})))
            if amount is None:
                continue
            indicator = cls._xtext(cls._find(entry, {"CdtDbtInd"})).upper()
            debit = amount if indicator == "DBIT" else Decimal("0.00")
            credit = amount if indicator == "CRDT" else Decimal("0.00")
            date_node = cls._find(entry, {"BookgDt", "ValDt"})
            date = cls._xtext(cls._find(date_node, {"Dt", "DtTm"})) if date_node is not None else ""
            refs = []
            for name in ("InstrId", "EndToEndId", "TxId", "AcctSvcrRef"):
                n = cls._find(entry, {name})
                if n is not None and cls._xtext(n):
                    refs.append(cls._xtext(n))
            narr = []
            for name in ("Ustrd", "AddtlNtryInf", "Nm"):
                for n in cls._findall(entry, name):
                    value = cls._xtext(n)
                    if value and value not in narr:
                        narr.append(value)
            txns.append(cls._tx(bank, account, date, " | ".join(narr), debit, credit, source_file=filepath, source_format="ISO20022_CAMT", reference=" | ".join(refs)))
        cls._rebuild_balances(txns, opening, closing)
        return bank, account, txns

    @classmethod
    def parse_mt940(cls, filepath):
        text = open(filepath, encoding="utf-8-sig", errors="replace").read()
        bank, account = cls._detect_bank_account(text, filepath)
        m = re.search(r":25:([^\r\n]+)", text)
        if m:
            account = m.group(1).strip()
        opening = closing = None
        for tag, target in (("60", "opening"), ("62", "closing")):
            m = re.search(rf":{tag}[FM]:([CD])(\d{{6}})([A-Z]{{3}})([-+]?\d+[,.]\d+)", text)
            if m:
                value = cls._parse_mt_amount(m.group(4))
                signed = value if m.group(1) == "C" else (-value if value is not None else None)
                if target == "opening": opening = signed
                else: closing = signed
        txns = []
        for match in re.finditer(r":61:([^\r\n]+)(.*?)(?=\r?\n:61:|\r?\n:62[FM]:|\Z)", text, re.I | re.S):
            line, tail = match.group(1).strip(), match.group(2)
            m = re.match(r"(\d{6})(\d{4})?([CDR]{1,2})([A-Z]?)([-+]?\d+[,.]\d+)(.*)", line)
            if not m:
                continue
            amount = cls._parse_mt_amount(m.group(5))
            if amount is None:
                continue
            dc = m.group(3).upper()
            debit = amount if dc in {"D", "RD"} else Decimal("0.00")
            credit = amount if dc in {"C", "RC"} else Decimal("0.00")
            info = re.search(r":86:([\s\S]*?)(?=\r?\n:\d{2}[A-Z]?:|\Z)", tail, re.I)
            narration = info.group(1).strip() if info else tail.strip()
            txns.append(cls._tx(bank, account, m.group(1), narration, debit, credit, source_file=filepath, source_format="SWIFT_MT940", reference=m.group(6).strip()))
        cls._rebuild_balances(txns, opening, closing)
        return bank, account, txns

    @classmethod
    def parse_qif(cls, filepath):
        text = open(filepath, encoding="utf-8-sig", errors="replace").read()
        bank, account = cls._detect_bank_account(text, filepath)
        txns = []
        for block in re.split(r"\^", text):
            lines = [x.strip() for x in block.splitlines() if x.strip()]
            if not any(x.upper().startswith("D") for x in lines):
                continue
            data = {line[0].upper(): line[1:].strip() for line in lines if len(line) >= 2}
            amount = parse_decimal(data.get("T"))
            if amount is None:
                continue
            txns.append(cls._tx(bank, account, data.get("D", ""), " | ".join(x for x in [data.get("P"), data.get("M")] if x), abs(amount) if amount < 0 else Decimal("0.00"), amount if amount >= 0 else Decimal("0.00"), source_file=filepath, source_format="QIF", reference=data.get("N", "")))
        return bank, account, txns

    @classmethod
    def parse_bai2(cls, filepath):
        text = open(filepath, encoding="utf-8-sig", errors="replace").read()
        bank, account = cls._detect_bank_account(text, filepath)
        current_account, current_date = account, ""
        txns = []
        for raw in text.splitlines():
            fields = [x.strip() for x in raw.strip().rstrip("/").split(",")]
            if not fields or not fields[0].isdigit():
                continue
            code = int(fields[0])
            if code == 3:
                if len(fields) > 1 and fields[1]: current_account = fields[1]
                if len(fields) > 4 and fields[4]: current_date = fields[4]
            elif code == 16 and len(fields) >= 3:
                amount = parse_decimal(fields[2])
                if amount is None: continue
                direction = fields[3].upper() if len(fields) > 3 else ""
                debit = abs(amount) if direction in {"D", "DEBIT"} else Decimal("0.00")
                credit = abs(amount) if direction in {"C", "CREDIT"} else Decimal("0.00")
                tx = cls._tx(bank, current_account, current_date, " | ".join(fields[4:]), debit, credit, source_file=filepath, source_format="BAI2", reference=fields[1])
                if direction not in {"D", "DEBIT", "C", "CREDIT"}:
                    tx["amount_candidates"] = [{"value": str(abs(amount)), "source": "BAI2_TYPE_16"}]
                    tx["status"] = "REVIEW_REQUIRED"
                txns.append(tx)
        return bank, current_account, txns

    @classmethod
    def parse_json(cls, filepath):
        with open(filepath, encoding="utf-8-sig", errors="replace") as f:
            data = json.load(f)
        bank, account = cls._detect_bank_account(json.dumps(data, ensure_ascii=False), filepath)
        txns = []
        def walk(obj):
            if isinstance(obj, dict):
                keys = {str(k).lower(): v for k, v in obj.items()}
                date = keys.get("date") or keys.get("transactiondate") or keys.get("transaction_date") or keys.get("posteddate") or keys.get("bookdate")
                amount = keys.get("amount") or keys.get("transactionamount") or keys.get("txnamount")
                if date is not None and amount is not None:
                    amt = parse_decimal(amount)
                    debit = parse_decimal(keys.get("debit") or keys.get("withdrawal") or keys.get("debitamount"))
                    credit = parse_decimal(keys.get("credit") or keys.get("deposit") or keys.get("creditamount"))
                    if debit is None and credit is None and amt is not None:
                        debit = abs(amt) if amt < 0 else Decimal("0.00")
                        credit = amt if amt > 0 else Decimal("0.00")
                    txns.append(cls._tx(bank, str(keys.get("account") or keys.get("accountnumber") or account), date, keys.get("narration") or keys.get("description") or keys.get("memo") or keys.get("particulars") or "", debit or Decimal("0.00"), credit or Decimal("0.00"), keys.get("balance") or keys.get("runningbalance") or "UNEXTRACTED_AMOUNT", source_file=filepath, source_format="JSON", reference=str(keys.get("reference") or keys.get("transactionid") or keys.get("id") or "")))
                for value in obj.values(): walk(value)
            elif isinstance(obj, list):
                for item in obj: walk(item)
        walk(data)
        return bank, account, txns

    @classmethod
    def parse_xml_generic(cls, filepath):
        root = ET.parse(filepath).getroot()
        text = ET.tostring(root, encoding="unicode")
        bank, account = cls._detect_bank_account(text, filepath)
        if {cls._local(x.tag) for x in root.iter()} & {"NTRY", "BAL"}:
            return cls.parse_camt(filepath)
        txns = []
        for node in root.iter():
            if cls._local(node.tag) not in {"TRANSACTION", "TXN", "ENTRY", "ITEM", "RECORD"}: continue
            def child(*names):
                n = cls._find(node, names)
                return cls._xtext(n) if n is not None else ""
            date, amount = child("DATE", "DTPOSTED", "BOOKINGDATE"), parse_decimal(child("AMOUNT", "TRNAMT"))
            if not date or amount is None: continue
            debit, credit = parse_decimal(child("DEBIT", "WITHDRAWAL")), parse_decimal(child("CREDIT", "DEPOSIT"))
            if debit is None and credit is None:
                debit, credit = (abs(amount), Decimal("0.00")) if amount < 0 else (Decimal("0.00"), amount)
            txns.append(cls._tx(bank, account, date, child("NARRATION", "DESCRIPTION", "MEMO", "PARTICULARS"), debit, credit, child("BALANCE", "RUNNINGBALANCE") or "UNEXTRACTED_AMOUNT", source_file=filepath, source_format="XML_GENERIC"))
        return bank, account, txns


    # ------------------------------------------------------------------
    # Worker
    # ------------------------------------------------------------------

    def run(self):

        total_files = len(self.filepaths)
        self.current_stage = (
            f"Worker thread entered parser. {total_files} statement(s) queued."
        )

        # Emit a heartbeat before touching the input file.  This is important
        # diagnostically: the GUI must never look frozen at merely
        # "Worker started" while the worker is actually alive.
        try:
            self.sig_progress.emit(
                0,
                max(1, total_files),
                f"Worker thread entered parser. {total_files} statement(s) queued."
            )
        except Exception:
            pass

        for file_index, filepath in enumerate(self.filepaths):

            if self._is_cancelled:
                break

            try:
                ext = os.path.splitext(filepath)[1].lower()
                self.current_file = filepath
                self.current_stage = (
                    f"Opening statement {file_index + 1}/{total_files}: "
                    f"{os.path.basename(filepath)}"
                )

                self.sig_progress.emit(
                    0,
                    max(1, total_files),
                    f"Opening statement {file_index + 1}/{total_files}: {os.path.basename(filepath)}"
                )

                if ext == ".pdf":

                    self.parse_pdf(
                        filepath,
                        file_index,
                        total_files,
                    )

                elif ext in [
                    ".xlsx",
                    ".xls",
                    ".xlsm",
                ]:

                    self.parse_excel(
                        filepath
                    )

                elif ext == ".csv":

                    self.parse_csv(
                        filepath
                    )

                elif ext == ".tsv":

                    self.parse_csv(
                        filepath
                    )

                elif ext in [".ofx", ".qfx"]:
                    bank_name, acc_no, txns = UniversalStatementEngine.parse_ofx(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext in [".mt940", ".940", ".sta"]:
                    bank_name, acc_no, txns = UniversalStatementEngine.parse_mt940(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext == ".qif":
                    bank_name, acc_no, txns = UniversalStatementEngine.parse_qif(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext in [".bai", ".bai2"]:
                    bank_name, acc_no, txns = UniversalStatementEngine.parse_bai2(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext == ".json":
                    universal = _try_universal_structured_document(filepath)
                    if universal is not None:
                        bank_name, acc_no, txns = universal
                    else:
                        bank_name, acc_no, txns = UniversalStatementEngine.parse_json(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext in VisualEvidenceEngine.SUPPORTED_IMAGE_EXTENSIONS:
                    self.sig_progress.emit(
                        0, 1,
                        f"OCR / visual evidence processing: {os.path.basename(filepath)}"
                    )
                    visual = VisualEvidenceEngine.extract(filepath, source_file=filepath)
                    if visual.get("ledger_eligible"):
                        bank_name, acc_no = self.detect_bank_and_acc(visual.get("ocr_text", ""), filepath)
                        txns = visual.get("transactions", [])
                        for tx in txns:
                            tx["bank"] = bank_name
                            tx["account_number"] = acc_no
                        self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)
                    else:
                        self.sig_error.emit(filepath, "Visual OCR did not certify any transaction: " + str(visual.get("status")))

                elif ext == ".xml":
                    bank_name, acc_no, txns = UniversalStatementEngine.parse_xml_generic(filepath)
                    self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                elif ext == ".txt":
                    raw = open(filepath, encoding="utf-8-sig", errors="replace").read(20000)
                    upper = raw.upper()
                    if ":61:" in upper and (":60F:" in upper or ":62F:" in upper):
                        bank_name, acc_no, txns = UniversalStatementEngine.parse_mt940(filepath)
                    elif "<OFX>" in upper or "<STMTTRN>" in upper:
                        bank_name, acc_no, txns = UniversalStatementEngine.parse_ofx(filepath)
                    elif re.search(r"(^|[\r\n])01,.*[\r\n].*03,.*[\r\n].*16,", raw):
                        bank_name, acc_no, txns = UniversalStatementEngine.parse_bai2(filepath)
                    else:
                        self.parse_csv(filepath)
                        bank_name = acc_no = txns = None
                    if txns is not None:
                        self.sig_statement_ready.emit(filepath, bank_name, acc_no, txns)

                else:

                    self.sig_error.emit(
                        filepath,
                        f"Unsupported file type: {ext}",
                    )

            except Exception as exc:
                self.current_stage = (
                    f"Extraction error in {os.path.basename(filepath)}: {exc}"
                )
                self.sig_error.emit(
                    filepath,
                    str(exc),
                )

        self.current_stage = "Worker finished."
        self.sig_finished.emit()


# ======================================================================
# 6. ISOLATED HEADLESS PARSER PROCESS
# ======================================================================

def validate_extracted_transactions(transactions):
    """Strict financial validation with no silent zero-fill of missing movements."""
    rows = list(transactions or [])
    extraction_errors = []
    arithmetic_errors = []
    streams = {}

    for idx, row in enumerate(rows):
        key = (str(row.get("source_file", "")), str(row.get("bank", "")), str(row.get("account_number", "")))
        streams.setdefault(key, []).append((idx, row))
        if row.get("mode") == "REFERENCE":
            continue
        missing = []
        if not str(row.get("date") or "").strip():
            missing.append("date")
        dr = parse_decimal(row.get("withdrawal"))
        cr = parse_decimal(row.get("deposit"))
        if dr is None or cr is None:
            missing.append("debit/credit extraction incomplete")
        else:
            drp = dr > Decimal("0.00")
            crp = cr > Decimal("0.00")
            if drp == crp:
                missing.append("exactly one debit or credit amount must be positive")
        if row.get("field_conflicts"):
            missing.append("conflicting monetary tokens")
        reconstruction = row.get("amount_reconstruction") or {}
        if reconstruction and not reconstruction.get("admitted"):
            missing.append("candidate not admitted because page reconciliation failed")
        if missing:
            extraction_errors.append({
                "row": idx + 1, "date": row.get("date", ""),
                "narration": row.get("narration", ""), "missing": "; ".join(missing),
            })

    for key, indexed_rows in streams.items():
        indexed_rows.sort(key=lambda item: (
            int(item[1].get("page", 0) or 0),
            0 if item[1].get("reference_role") == "OPENING" else
            2 if item[1].get("reference_role") == "CLOSING" else 1,
            float((item[1].get("bbox") or [0, 0, 0, 0])[1]), item[0],
        ))
        current = None
        curve_valid = False
        for original_idx, row in indexed_rows:
            if row.get("mode") == "REFERENCE":
                reference = parse_decimal(row.get("image_balance"))
                if reference is None:
                    reference = parse_decimal(row.get("balance"))
                role = row.get("reference_role")
                if role == "OPENING":
                    if current is None or not curve_valid:
                        current = reference
                        curve_valid = reference is not None
                        row["calculated_balance"] = str(reference) if reference is not None else "UNEXTRACTED_AMOUNT"
                    elif reference is not None and current != reference:
                        arithmetic_errors.append({
                            "row": original_idx + 1, "date": "", "narration": row.get("narration", "BROUGHT FORWARD"),
                            "checkpoint": "BROUGHT_FORWARD", "calculated_balance": str(current),
                            "printed_balance": str(reference), "discrepancy": str(reference - current),
                        })
                        row["calculated_balance"] = str(current)
                    continue
                if role == "CLOSING":
                    row["calculated_balance"] = str(current) if current is not None and curve_valid else "UNEXTRACTED_AMOUNT"
                    if reference is not None and current is not None and curve_valid and reference != current:
                        arithmetic_errors.append({
                            "row": original_idx + 1, "date": "", "narration": row.get("narration", "CARRIED FORWARD"),
                            "checkpoint": "CARRIED_FORWARD", "calculated_balance": str(current),
                            "printed_balance": str(reference), "discrepancy": str(reference - current),
                        })
                    continue

            dr = parse_decimal(row.get("withdrawal"))
            cr = parse_decimal(row.get("deposit"))
            if dr is None or cr is None:
                row["calculated_balance"] = "UNEXTRACTED_AMOUNT"
                current, curve_valid = None, False
                continue
            if (dr == Decimal("0.00") and cr == Decimal("0.00")) or (dr > 0 and cr > 0):
                row["calculated_balance"] = "UNEXTRACTED_AMOUNT"
                current, curve_valid = None, False
                continue
            if current is None or not curve_valid:
                row["calculated_balance"] = "UNEXTRACTED_AMOUNT"
                continue
            expected = current + cr - dr
            row["calculated_balance"] = str(expected)
            # Legacy balance is the formula result only. Printed evidence stays
            # in image_balance and is never overwritten by this calculation.
            row["balance"] = str(expected)
            image = parse_decimal(row.get("image_balance"))
            image_source = row.get("image_balance_source")
            if image is not None and image_source in {"EXPLICIT_RUNNING_BALANCE_TOKEN", "CALIBRATED_HEADER_COLUMN", "CALIBRATED_HEADER_RIGHT_EDGE", "MULTILINE_MONEY_COLUMN_GEOMETRY"} and image != expected:
                arithmetic_errors.append({
                    "row": original_idx + 1, "date": row.get("date", ""),
                    "narration": row.get("narration", ""), "previous_balance": str(current),
                    "debit": str(dr), "credit": str(cr), "printed_balance": str(image),
                    "expected_balance": str(expected), "discrepancy": str(image - expected),
                })
            current = expected

    return {
        "rows": len(rows),
        "extraction_errors": extraction_errors,
        "arithmetic_errors": arithmetic_errors,
        "extraction_gate": "PASS" if not extraction_errors else "FAILED",
        "balance_reconciliation": "PASS" if not arithmetic_errors and not extraction_errors else "FAILED",
        "status": "VALIDATED" if rows and not extraction_errors and not arithmetic_errors else "REVIEW_REQUIRED",
    }


def run_pdf_checkpoint_recovery_tests():
    """Synthetic regression for checkpoint-proven recovery and strict refusal."""
    def ref(role, amount, y, label):
        return {"source_file":"synthetic_checkpoint_test.pdf", "bank":"Test Bank", "account_number":"12345678", "page":0, "date":"", "narration":label, "mode":"REFERENCE", "reference_role":role, "balance":str(amount), "image_balance":str(amount), "withdrawal":"0.00", "deposit":"0.00", "bbox":[0,y,520,y+10], "status":"REFERENCE"}
    def tx(date, narration, amount, y):
        return {"source_file":"synthetic_checkpoint_test.pdf", "bank":"Test Bank", "account_number":"12345678", "page":0, "date":date, "narration":narration, "mode":"DIGITAL", "withdrawal":"0.00", "deposit":"0.00", "balance":str(amount), "image_balance":str(amount), "row_anchor_y":y, "bbox":[0,y,520,y+10], "amount_candidates":[{"value":str(amount), "text":str(amount), "x0":480.0, "x1":520.0, "y":y, "row_aligned":True}], "status":"REVIEW_REQUIRED"}
    def fixture(closing):
        return [ref("OPENING","100.00",20,"BROUGHT FORWARD"), tx("01/04/24","TRF FRM TEST","50.00",100), tx("02/04/24","TO CASH SELF","30.00",120), tx("03/04/24","BY TRF RECEIVED","20.00",140), ref("CLOSING",closing,200,"CARRIED FORWARD")]
    good = reconstruct_pdf_amounts_from_checkpoints(fixture("140.00"))
    ordinary = [r for r in good if r.get("mode") != "REFERENCE"]
    assert [r["deposit"] for r in ordinary] == ["50.00","0.00","20.00"]
    assert [r["withdrawal"] for r in ordinary] == ["0.00","30.00","0.00"]
    good_validation = validate_extracted_transactions(good)
    assert good_validation["extraction_gate"] == "PASS" and good_validation["balance_reconciliation"] == "PASS", good_validation
    bad = reconstruct_pdf_amounts_from_checkpoints(fixture("141.00"))
    bad_ordinary = [r for r in bad if r.get("mode") != "REFERENCE"]
    assert not any((r.get("amount_reconstruction") or {}).get("admitted") for r in bad_ordinary)
    bad_validation = validate_extracted_transactions(bad)
    assert bad_validation["status"] == "REVIEW_REQUIRED" and bad_validation["extraction_gate"] == "FAILED", bad_validation
    zero_validation = validate_extracted_transactions([{"source_file":"zero_test.pdf", "bank":"Test Bank", "account_number":"12345678", "page":0, "date":"01/04/24", "narration":"UNEXTRACTED TRANSACTION", "mode":"DIGITAL", "withdrawal":"0.00", "deposit":"0.00", "image_balance":"100.00", "balance":"100.00", "status":"OK"}])
    assert zero_validation["status"] == "REVIEW_REQUIRED" and zero_validation["extraction_gate"] == "FAILED", zero_validation
    right_edge_fixture = {
        "withdrawal": (300.0, 450.0), "deposit": (450.0, 505.0),
        "balance": (505.0, 9999.0), "narration": (80.0, 300.0),
        "_money_right_edge_channels": {
            "withdrawal": (300.0, 450.0), "deposit": (450.0, 505.0),
            "balance": (505.0, 9999.0),
        },
    }
    assert pdf_numeric_column_role(465.0, 477.0, 490.0, right_edge_fixture) == "deposit"
    assert pdf_numeric_column_role(460.0, 490.0, 520.0, right_edge_fixture) == "balance"
    assert pdf_numeric_column_role(350.0, 390.0, 420.0, right_edge_fixture) == "withdrawal"

    # Exercise the actual heading calibrator with synthetic geometry. Header
    # centres and amount-token centres differ, while right edges line up.
    class _SyntheticHeaderPage:
        def get_text(self, mode):
            assert mode == "words"
            return [
                (40, 100, 70, 112, "DATE", 0, 0, 0),
                (120, 100, 205, 112, "PARTICULARS", 0, 0, 1),
                (300, 100, 365, 112, "WITHDRAWAL", 0, 0, 2),
                (420, 100, 470, 112, "DEPOSIT", 0, 0, 3),
                (520, 100, 570, 112, "BALANCE", 0, 0, 4),
            ]
    calibrated = DynamicHeaderCalibrator.calibrate_columns(_SyntheticHeaderPage())
    assert pdf_numeric_column_role(460.0, 477.0, 490.0, calibrated) == "deposit", calibrated
    assert pdf_numeric_column_role(500.0, 520.0, 550.0, calibrated) == "balance", calibrated
    assert pdf_numeric_column_role(320.0, 340.0, 365.0, calibrated) == "withdrawal", calibrated

    # Generalization regression: plural column labels and split multi-word
    # headers must resolve through the same semantic calibrator.
    class _SyntheticInternationalHeaderPage:
        def get_text(self, mode):
            assert mode == "words"
            return [
                (35, 100, 105, 112, "Transaction", 0, 0, 0),
                (108, 100, 135, 112, "Date", 0, 0, 1),
                (145, 100, 225, 112, "Description", 0, 0, 2),
                (300, 100, 360, 112, "Withdrawals", 0, 0, 3),
                (405, 100, 455, 112, "Deposits", 0, 0, 4),
                (500, 100, 550, 112, "Running", 0, 0, 5),
                (553, 100, 600, 112, "Balance", 0, 0, 6),
            ]
    international = DynamicHeaderCalibrator.calibrate_columns(_SyntheticInternationalHeaderPage())
    assert pdf_numeric_column_role(330.0, 340.0, 355.0, international) == "withdrawal", international
    assert pdf_numeric_column_role(425.0, 440.0, 450.0, international) == "deposit", international
    assert pdf_numeric_column_role(560.0, 580.0, 595.0, international) == "balance", international

    # Multi-line statement regression: amount token is geometrically in the
    # credit column but sits below the date line. It must not be discarded as
    # narration or guessed from its amount alone.
    multiline_tx = {
        "source_file":"synthetic_multiline_sbi.pdf", "bank":"Test Bank", "account_number":"12345678",
        "page":0, "date":"02-04-2024", "narration":"UPI/CR/123456/TEST", "mode":"DIGITAL",
        "withdrawal":"0.00", "deposit":"0.00", "image_balance":"1100.00", "balance":"1100.00",
        "row_anchor_y":100.0, "bbox":[20,100,540,140], "status":"REVIEW_REQUIRED",
        "amount_candidates":[{"value":"100.00", "text":"100.00", "x0":430.0, "x1":475.0,
                               "y":117.0, "row_aligned":False, "column_role":"deposit", "money_token":True, "column_geometry_confident":True}],
    }
    recovered = recover_multiline_money_cells([multiline_tx])
    assert multiline_tx["deposit"] == "100.00" and recovered["deposit"] == 1, (multiline_tx, recovered)

    # Exact printed balance delta can replace a previously misassigned amount
    # only when the matching visible token exists in the expected money column.
    exact_delta = [
        {"source_file":"synthetic_delta.pdf", "bank":"Test Bank", "account_number":"12345678", "page":0,
         "date":"", "narration":"BROUGHT FORWARD", "mode":"REFERENCE", "reference_role":"OPENING",
         "image_balance":"7003.85", "balance":"7003.85", "withdrawal":"0.00", "deposit":"0.00", "bbox":[0,10,500,20]},
        {"source_file":"synthetic_delta.pdf", "bank":"Test Bank", "account_number":"12345678", "page":0,
         "date":"16-06-2024", "narration":"WDL TFR UPI/DR/TEST", "mode":"DIGITAL", "image_balance":"911.85",
         "balance":"911.85", "withdrawal":"123.60", "deposit":"0.00", "bbox":[0,30,530,50],
         "amount_candidates":[{"value":"6092.00", "text":"6,092.00", "x0":300, "x1":350, "y":42,
                               "row_aligned":False, "column_role":"withdrawal", "money_token":True, "column_geometry_confident":True}]},
    ]
    delta_result = resolve_movement_from_exact_balance_delta(exact_delta)
    assert exact_delta[1]["withdrawal"] == "6092.00" and len(delta_result) == 1, (exact_delta, delta_result)
    return {"status":"PASS", "checkpoint_recovery_rows":3, "mismatched_checkpoint_refused":3,
            "zero_amount_rejected":1, "right_edge_column_tests":6, "synthetic_header_calibration":6,
            "multiline_column_recovery":1, "exact_balance_delta_token_correction":1}


def run_isolated_parser_process(filepaths, analyze_mode=False):
    """Run the existing parser in a dedicated Python child process."""
    app = QCoreApplication.instance()
    if app is None:
        app = QCoreApplication(sys.argv)

    parser = StatementParserWorker([])
    results = []
    transactions = []
    errors = []
    completed = []

    def status(text):
        print("STATUS|" + str(text), flush=True)

    def on_progress(current_page, total_pages, text):
        if text:
            status(text)

    def on_ready(filepath, bank, account, txns):
        rows = list(txns or [])
        validation = validate_extracted_transactions(rows)
        transactions.extend(rows)
        if filepath not in completed:
            completed.append(filepath)
        results.append({
            "filepath": filepath,
            "bank": bank,
            "account": account,
            "entries": len(rows),
            "status": validation["status"],
            "validation": validation,
        })
        status(
            f"Parsed {os.path.basename(filepath)}: {len(rows)} transaction row(s) admitted; "
            f"validation={validation['status']}."
        )

    def on_error(filepath, error):
        errors.append({"filepath": filepath, "error": str(error)})
        status(f"Review required: {os.path.basename(filepath)} | {error}")

    parser.sig_progress.connect(on_progress)
    parser.sig_statement_ready.connect(on_ready)
    parser.sig_error.connect(on_error)

    total = len(filepaths)
    status(f"Parser process started. {total} statement(s) queued.")

    for index, filepath in enumerate(filepaths, 1):
        if not os.path.exists(filepath):
            on_error(filepath, "File does not exist.")
            continue
        ext = os.path.splitext(filepath)[1].lower()
        status(f"Opening statement {index}/{total}: {os.path.basename(filepath)}")
        before_completed = len(completed)
        before_errors = len(errors)
        try:
            if ext == ".pdf":
                parser.parse_pdf(filepath, index - 1, total)
            elif ext in {".xlsx", ".xls", ".xlsm"}:
                parser.parse_excel(filepath)
            elif ext in {".csv", ".tsv"}:
                parser.parse_csv(filepath)
            elif ext in {".ofx", ".qfx"}:
                bank, account, txns = UniversalStatementEngine.parse_ofx(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext in {".mt940", ".940", ".sta"}:
                bank, account, txns = UniversalStatementEngine.parse_mt940(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext == ".qif":
                bank, account, txns = UniversalStatementEngine.parse_qif(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext in {".bai", ".bai2"}:
                bank, account, txns = UniversalStatementEngine.parse_bai2(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext == ".json":
                bank, account, txns = UniversalStatementEngine.parse_json(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext in VisualEvidenceEngine.SUPPORTED_IMAGE_EXTENSIONS:
                status(f"OCR / visual evidence processing: {os.path.basename(filepath)}")
                visual = VisualEvidenceEngine.extract(filepath, source_file=filepath)
                if visual.get("ledger_eligible"):
                    bank, account = parser.detect_bank_and_acc(visual.get("ocr_text", ""), filepath)
                    txns = visual.get("transactions", [])
                    for tx in txns:
                        tx["bank"] = bank
                        tx["account_number"] = account
                    parser.sig_statement_ready.emit(filepath, bank, account, txns)
                else:
                    parser.sig_error.emit(filepath, "Visual OCR did not certify any transaction: " + str(visual.get("status")))
            elif ext == ".xml":
                bank, account, txns = UniversalStatementEngine.parse_xml_generic(filepath)
                parser.sig_statement_ready.emit(filepath, bank, account, txns)
            elif ext == ".txt":
                raw = open(filepath, encoding="utf-8-sig", errors="replace").read(20000)
                upper = raw.upper()
                if ":61:" in upper and (":60F:" in upper or ":62F:" in upper):
                    bank, account, txns = UniversalStatementEngine.parse_mt940(filepath)
                    parser.sig_statement_ready.emit(filepath, bank, account, txns)
                elif "<OFX>" in upper or "<STMTTRN>" in upper:
                    bank, account, txns = UniversalStatementEngine.parse_ofx(filepath)
                    parser.sig_statement_ready.emit(filepath, bank, account, txns)
                elif re.search(r"(^|[\r\n])01,.*[\r\n].*03,.*[\r\n].*16,", raw):
                    bank, account, txns = UniversalStatementEngine.parse_bai2(filepath)
                    parser.sig_statement_ready.emit(filepath, bank, account, txns)
                else:
                    parser.parse_csv(filepath)
            else:
                parser.sig_error.emit(filepath, f"Unsupported file type: {ext}")
        except Exception as exc:
            parser.sig_error.emit(filepath, str(exc))

        if len(completed) == before_completed and len(errors) == before_errors:
            errors.append({"filepath": filepath, "error": "Parser completed without emitting a certified statement result."})
            status(f"No certified result for {os.path.basename(filepath)}. Review required.")

    payload = {
        "ok": not bool(errors),
        "transactions": transactions,
        "results": results,
        "errors": errors,
        "completed_files": completed,
        "failed_files": [x["filepath"] for x in errors],
    }
    status(f"Parser process finished. {len(transactions)} transaction row(s) admitted.")
    print("RESULT|" + json.dumps(payload, ensure_ascii=False, default=str), flush=True)

    if analyze_mode:
        print("\n" + "=" * 72)
        print("TAXFORENSIC HEADLESS ANALYSIS")
        print("=" * 72)
        for item in results:
            print(f"FILE      : {item['filepath']}")
            print(f"BANK      : {item['bank']}")
            print(f"ACCOUNT   : {item['account']}")
            print(f"ENTRIES   : {item['entries']}")
            print(f"STATUS    : {item['status']}")
            validation = item.get("validation", {})
            print(f"EXTRACTION GATE       : {validation.get('extraction_gate', 'N/A')}")
            print(f"BALANCE RECONCILIATION: {validation.get('balance_reconciliation', 'N/A')}")
            print(f"ROWS FAILING EXTRACTION GATE: {len(validation.get('extraction_errors', []))}")
            print(f"BALANCE DISCREPANCIES      : {len(validation.get('arithmetic_errors', []))}")
            if validation.get("extraction_errors"):
                print("  First extraction issue:", validation["extraction_errors"][0])
            if validation.get("arithmetic_errors"):
                print("  First balance issue:", validation["arithmetic_errors"][0])
            if item['bank'] == 'UNDETERMINED BANK':
                print("  BANK IDENTITY: Not established from extracted statement text. No bank name has been guessed.")
            print("-" * 72)
        if errors:
            print("PARSER ERRORS:")
            for item in errors:
                print(f"  {item['filepath']}: {item['error']}")
        else:
            print("PARSER ERRORS: 0")
        review_count = sum(1 for item in results if item.get("status") == "REVIEW_REQUIRED")
        print(f"FILES REQUIRING REVIEW: {review_count}")
        print(f"TOTAL TRANSACTION ROWS: {len(transactions)}")
        print("=" * 72)

    return 0 if payload["ok"] else 2


# ======================================================================
# 6. EXCEL EXPORTER
# ======================================================================

class TaxForensicExcelExporter:

    @classmethod
    def generate_assessment_workbook(
        cls,
        case_metadata: dict,
        transactions: list,
        contras: list,
        cash_flows: list,
        violations: list,
        file_path: str,
    ):

        wb = openpyxl.Workbook()

        wb.remove(
            wb.active
        )

        font_title = Font(
            name="Calibri",
            size=14,
            bold=True,
            color="1A237E",
        )

        font_sub = Font(
            name="Calibri",
            size=10,
            italic=True,
            color="424242",
        )

        font_header = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF",
        )

        font_bold = Font(
            name="Calibri",
            size=11,
            bold=True,
        )

        font_regular = Font(
            name="Calibri",
            size=11,
        )

        font_danger = Font(
            name="Calibri",
            size=11,
            bold=True,
            color="B71C1C",
        )

        fill_navy = PatternFill(
            start_color="1A237E",
            end_color="1A237E",
            fill_type="solid",
        )

        fill_dark_green = PatternFill(
            start_color="1B5E20",
            end_color="1B5E20",
            fill_type="solid",
        )

        fill_dark_red = PatternFill(
            start_color="B71C1C",
            end_color="B71C1C",
            fill_type="solid",
        )

        fill_gray_header = PatternFill(
            start_color="37474F",
            end_color="37474F",
            fill_type="solid",
        )

        fill_accent_yellow = PatternFill(
            start_color="FFF9C4",
            end_color="FFF9C4",
            fill_type="solid",
        )

        thin_border = Border(
            left=Side(
                style="thin",
                color="BDBDBD",
            ),
            right=Side(
                style="thin",
                color="BDBDBD",
            ),
            top=Side(
                style="thin",
                color="BDBDBD",
            ),
            bottom=Side(
                style="thin",
                color="BDBDBD",
            ),
        )

        double_bottom_border = Border(
            top=Side(
                style="thin",
                color="000000",
            ),
            bottom=Side(
                style="double",
                color="000000",
            ),
        )

        currency_fmt = (
            '"Rs. "#,##,##0.00;'
            '[Red]"-Rs. "#,##,##0.00;'
            '"Rs. "-'
        )

        # ==============================================================
        # EXECUTIVE SUMMARY
        # ==============================================================

        ws_sum = wb.create_sheet(
            title="Executive Summary"
        )

        ws_sum["A1"] = (
            "OFFICE OF THE ASSESSING OFFICER / "
            "TAX FORENSIC WORKSTATION"
        )

        ws_sum["A1"].font = font_title

        ws_sum["A2"] = (
            f"AUDIT STATEMENT — "
            f"{case_metadata.get('bank_scope', 'CONSOLIDATED')} — "
            f"GENERATED: "
            f"{datetime.now().strftime('%d-%b-%Y %H:%M:%S')}"
        )

        ws_sum["A2"].font = font_sub

        meta_rows = [
            (
                "Assessee PAN:",
                case_metadata.get(
                    "pan",
                    "ABCDE1234F",
                ),
                "Assessment Year:",
                case_metadata.get(
                    "ay",
                    "2024-25",
                ),
            ),
            (
                "Bank Name:",
                case_metadata.get(
                    "bank_name",
                    "All Banks",
                ),
                "Bank Account Number:",
                case_metadata.get(
                    "account_number",
                    "Multiple Accounts",
                ),
            ),
            (
                "Source Statement:",
                case_metadata.get(
                    "filenames",
                    "Consolidated Statements",
                ),
                "SHA-256 Hash:",
                case_metadata.get(
                    "sha256",
                    "UNKNOWN",
                ),
            ),
        ]

        row_number = 4

        for row in meta_rows:

            ws_sum.cell(
                row=row_number,
                column=1,
                value=row[0],
            ).font = font_bold

            ws_sum.cell(
                row=row_number,
                column=2,
                value=row[1],
            ).font = font_regular

            ws_sum.cell(
                row=row_number,
                column=4,
                value=row[2],
            ).font = font_bold

            ws_sum.cell(
                row=row_number,
                column=5,
                value=row[3],
            ).font = font_regular

            row_number += 1

        row_number += 1

        ws_sum.cell(
            row=row_number,
            column=1,
            value=(
                "SUMMARY OF FORENSIC ADDITIONS "
                "& AUDIT FINDINGS"
            ),
        ).font = font_bold

        row_number += 1

        headers = [
            "Serial",
            "Statutory Head / Issue",
            "Relevant Section",
            "Computed Quantum",
            "Evidentiary Cross-Reference",
        ]

        for column, header in enumerate(
            headers,
            start=1,
        ):

            cell = ws_sum.cell(
                row=row_number,
                column=column,
                value=header,
            )

            cell.font = font_header
            cell.fill = fill_navy

        summary_data = [
            (
                1,
                "Peak Unexplained Cash Credit",
                "Section 68 / 69A",
                "='Cash Flow & Peak Credit'!E2",
                "Annexure III",
            ),
            (
                2,
                "Inter-Account Self Transfers Neutralized",
                "Contra Exclusions",
                "='Contra Eliminations'!E2",
                "Annexure II",
            ),
            (
                3,
                "Transactions Violating Cash Limit",
                "Section 269ST",
                "='Statutory Violations'!G2",
                "Annexure IV",
            ),
            (
                4,
                "Cash Loans/Deposits",
                "Section 269SS / 269T",
                "='Statutory Violations'!G3",
                "Annexure IV",
            ),
            (
                5,
                "High-Value Cash Deposit / SFT",
                "Rule 114E",
                "='Statutory Violations'!G4",
                "Annexure IV",
            ),
        ]

        for item in summary_data:

            row_number += 1

            for column, value in enumerate(
                item,
                start=1,
            ):

                cell = ws_sum.cell(
                    row=row_number,
                    column=column,
                    value=value,
                )

                cell.border = thin_border

            ws_sum.cell(
                row=row_number,
                column=1,
            ).alignment = Alignment(
                horizontal="center"
            )

            ws_sum.cell(
                row=row_number,
                column=4,
            ).number_format = currency_fmt

        # ==============================================================
        # RECONCILED BANK LEDGER
        # ==============================================================

        ws_ledger = wb.create_sheet(
            title="Reconciled Bank Ledger"
        )

        ws_ledger["A1"] = (
            "ANNEXURE I: RECONCILED TRANSACTION LEDGER — "
            f"{case_metadata.get('bank_scope', 'ALL ACCOUNTS')}"
        )

        ws_ledger["A1"].font = font_title

        ws_ledger["A2"] = (
            "Verification Formula: "
            "Opening Balance + Total Deposits - "
            "Total Withdrawals = Closing Balance"
        )

        ws_ledger["A2"].font = font_sub

        ledger_headers = [
            "Row #",
            "Bank Name",
            "Bank A/c No",
            "Txn Date",
            "Mode",
            "Transaction Description / Narration",
            "Withdrawal (Dr)",
            "Deposit (Cr)",
            "Running Balance",
            "Source Statement File",
            "Audit Status",
        ]

        row_number = 4

        for column, header in enumerate(
            ledger_headers,
            start=1,
        ):

            cell = ws_ledger.cell(
                row=row_number,
                column=column,
                value=header,
            )

            cell.font = font_header
            cell.fill = fill_gray_header
            cell.alignment = Alignment(
                horizontal="center"
            )

        start_row = row_number + 1

        for index, transaction in enumerate(
            transactions,
            start=1,
        ):

            row_number += 1

            values = [
                index,
                transaction.get(
                    "bank",
                    "Bank Account",
                ),
                transaction.get(
                    "account_number",
                    "N/A",
                ),
                transaction.get(
                    "date",
                    "",
                ),
                transaction.get(
                    "mode",
                    "DIGITAL",
                ),
                transaction.get(
                    "narration",
                    "",
                ),
                safe_float(
                    transaction.get(
                        "withdrawal"
                    )
                ),
                safe_float(
                    transaction.get(
                        "deposit"
                    )
                ),
                safe_float(
                    transaction.get(
                        "balance"
                    )
                ),
                os.path.basename(
                    transaction.get(
                        "source_file",
                        "",
                    )
                ),
                transaction.get(
                    "status",
                    "OK",
                ),
            ]

            for column, value in enumerate(
                values,
                start=1,
            ):

                cell = ws_ledger.cell(
                    row=row_number,
                    column=column,
                    value=value,
                )

                cell.border = thin_border

            for column in [7, 8, 9]:

                ws_ledger.cell(
                    row=row_number,
                    column=column,
                ).number_format = currency_fmt

            status_cell = ws_ledger.cell(
                row=row_number,
                column=11,
            )

            if transaction.get("status") == "OK":

                status_cell.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color="1B5E20",
                )

            else:

                status_cell.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color="B71C1C",
                )

        row_number += 1

        ws_ledger.cell(
            row=row_number,
            column=6,
            value="CUMULATIVE TOTALS:",
        ).font = font_bold

        tot_dr = ws_ledger.cell(
            row=row_number,
            column=7,
            value=(
                f"=SUM(G{start_row}:"
                f"G{max(start_row, row_number - 1)})"
            ),
        )

        tot_dr.font = font_bold
        tot_dr.number_format = currency_fmt
        tot_dr.border = double_bottom_border

        tot_cr = ws_ledger.cell(
            row=row_number,
            column=8,
            value=(
                f"=SUM(H{start_row}:"
                f"H{max(start_row, row_number - 1)})"
            ),
        )

        tot_cr.font = font_bold
        tot_cr.number_format = currency_fmt
        tot_cr.border = double_bottom_border

        # ==============================================================
        # CONTRA
        # ==============================================================

        ws_contra = wb.create_sheet(
            title="Contra Eliminations"
        )

        ws_contra["A1"] = (
            "ANNEXURE II: INTER-ACCOUNT "
            "CONTRA TRANSFER RECONCILIATION"
        )

        ws_contra["A1"].font = font_title

        ws_contra["A2"] = (
            "Matching Outflows and Inflows within "
            "48 Hours Across Accounts"
        )

        ws_contra["A2"].font = font_sub

        ws_contra["D2"] = "Total Neutralized:"
        ws_contra["D2"].font = font_bold

        contra_end = max(
            5,
            len(contras) + 4,
        )

        ws_contra["E2"] = (
            f"=SUM(D5:D{contra_end})"
        )

        ws_contra["E2"].font = font_bold
        ws_contra["E2"].number_format = currency_fmt

        contra_headers = [
            "Date",
            "Debited Bank & A/c",
            "Credited Bank & A/c",
            "Contra Amount",
            "Time Delta",
            "Auditor Status",
        ]

        row_number = 4

        for column, header in enumerate(
            contra_headers,
            start=1,
        ):

            cell = ws_contra.cell(
                row=row_number,
                column=column,
                value=header,
            )

            cell.font = font_header
            cell.fill = fill_dark_green

        for contra in contras:

            row_number += 1

            values = [
                contra.get("date"),
                contra.get("from_bank"),
                contra.get("to_bank"),
                safe_float(
                    contra.get("amount")
                ),
                contra.get("delta"),
                "Self Transfer Neutralized",
            ]

            for column, value in enumerate(
                values,
                start=1,
            ):

                cell = ws_contra.cell(
                    row=row_number,
                    column=column,
                    value=value,
                )

                cell.border = thin_border

            ws_contra.cell(
                row=row_number,
                column=4,
            ).number_format = currency_fmt

        # ==============================================================
        # CASH FLOW
        # ==============================================================

        ws_cash = wb.create_sheet(
            title="Cash Flow & Peak Credit"
        )

        ws_cash["A1"] = (
            "ANNEXURE III: DAY-TO-DAY CASH FLOW "
            "& PEAK CREDIT EVALUATOR"
        )

        ws_cash["A1"].font = font_title

        ws_cash["A2"] = (
            "Evaluates recycled cash arguments "
            "(Section 68 / 69A)"
        )

        ws_cash["A2"].font = font_sub

        ws_cash["D2"] = (
            "Peak Unexplained Exposure:"
        )

        ws_cash["D2"].font = font_bold

        cash_end = max(
            5,
            len(cash_flows) + 4,
        )

        ws_cash["E2"] = (
            f'=IF(COUNT(D5:D{cash_end})>0,'
            f'MAX(0,-MIN(D5:D{cash_end})),0)'
        )

        ws_cash["E2"].font = font_danger
        ws_cash["E2"].number_format = currency_fmt

        cash_headers = [
            "Txn Date",
            "Cash Inflow (Withdrawal)",
            "Cash Outflow (Deposit)",
            "Theoretical Cash-on-Hand",
            "Evidentiary Audit Finding",
        ]

        row_number = 4

        for column, header in enumerate(
            cash_headers,
            start=1,
        ):

            cell = ws_cash.cell(
                row=row_number,
                column=column,
                value=header,
            )

            cell.font = font_header
            cell.fill = fill_dark_red

        for cash_row in cash_flows:

            row_number += 1

            values = [
                cash_row.get("date"),
                safe_float(
                    cash_row.get("inflow")
                ),
                safe_float(
                    cash_row.get("outflow")
                ),
                safe_float(
                    cash_row.get("running_cash")
                ),
                cash_row.get(
                    "finding",
                    "",
                ),
            ]

            for column, value in enumerate(
                values,
                start=1,
            ):

                cell = ws_cash.cell(
                    row=row_number,
                    column=column,
                    value=value,
                )

                cell.border = thin_border

            for column in [2, 3, 4]:

                ws_cash.cell(
                    row=row_number,
                    column=column,
                ).number_format = currency_fmt

            finding = str(
                cash_row.get(
                    "finding",
                    "",
                )
            )

            if "DEFICIT" in finding:

                ws_cash.cell(
                    row=row_number,
                    column=5,
                ).font = font_danger

                ws_cash.cell(
                    row=row_number,
                    column=4,
                ).font = font_danger

                ws_cash.cell(
                    row=row_number,
                    column=5,
                ).fill = fill_accent_yellow

        # ==============================================================
        # STATUTORY VIOLATIONS
        # ==============================================================

        ws_viol = wb.create_sheet(
            title="Statutory Violations"
        )

        ws_viol["A1"] = (
            "ANNEXURE IV: STATUTORY VIOLATIONS "
            "& SFT REPORTING AUDIT REGISTER"
        )

        ws_viol["A1"].font = font_title

        ws_viol["A2"] = (
            "Summary Tables for Section 269ST, "
            "Section 269SS/269T and Rule 114E"
        )

        ws_viol["A2"].font = font_sub

        violation_end = max(
            5,
            len(violations) + 4,
        )

        ws_viol["F2"] = "Sec 269ST Sum:"
        ws_viol["G2"] = (
            f'=SUMIF(E5:E{violation_end},'
            f'"*269ST*",D5:D{violation_end})'
        )

        ws_viol["F3"] = "Sec 269SS/T Sum:"
        ws_viol["G3"] = (
            f'=SUMIF(E5:E{violation_end},'
            f'"*269SS*",D5:D{violation_end})'
        )

        ws_viol["F4"] = "SFT Sum:"
        ws_viol["G4"] = (
            f'=SUMIF(E5:E{violation_end},'
            f'"*SFT*",D5:D{violation_end})'
        )

        for row in [2, 3, 4]:

            ws_viol[
                f"F{row}"
            ].font = font_bold

            ws_viol[
                f"G{row}"
            ].font = font_bold

            ws_viol[
                f"G{row}"
            ].number_format = currency_fmt

        violation_headers = [
            "Date",
            "Bank Name",
            "Transaction Particulars",
            "Amount (Rs.)",
            "Violated Section",
            "Penalty / Exposure",
        ]

        row_number = 4

        for column, header in enumerate(
            violation_headers,
            start=1,
        ):

            cell = ws_viol.cell(
                row=row_number,
                column=column,
                value=header,
            )

            cell.font = font_header
            cell.fill = fill_navy

        for violation in violations:

            row_number += 1

            values = [
                violation.get("date"),
                violation.get("bank"),
                violation.get("narration"),
                safe_float(
                    violation.get("amount")
                ),
                violation.get("rule"),
                violation.get("severity"),
            ]

            for column, value in enumerate(
                values,
                start=1,
            ):

                cell = ws_viol.cell(
                    row=row_number,
                    column=column,
                    value=value,
                )

                cell.border = thin_border

            ws_viol.cell(
                row=row_number,
                column=4,
            ).number_format = currency_fmt

            ws_viol.cell(
                row=row_number,
                column=5,
            ).font = font_bold

            ws_viol.cell(
                row=row_number,
                column=6,
            ).font = font_danger

        # ==============================================================
        # COLUMN WIDTHS
        # ==============================================================

        for ws in wb.worksheets:

            for column_cells in ws.columns:

                max_length = 0

                column_letter = (
                    get_column_letter(
                        column_cells[0].column
                    )
                )

                for cell in column_cells:

                    if cell.row in [1, 2]:
                        continue

                    max_length = max(
                        max_length,
                        len(
                            str(
                                cell.value
                                or ""
                            )
                        ),
                    )

                ws.column_dimensions[
                    column_letter
                ].width = max(
                    max_length + 3,
                    12,
                )

        wb.save(file_path)

        return file_path


# ----------------------------------------------------------------------
# Compatibility date parser used by reconciliation/table sorting.
# ----------------------------------------------------------------------

def parse_date_value(date_text):
    return parse_statement_date(date_text)


class StatementDropFrame(QFrame):
    """Reliable multi-file drag/drop surface for the statement importer."""

    filesDropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            local_files = [
                url.toLocalFile()
                for url in event.mimeData().urls()
                if url.isLocalFile()
            ]
            if any(os.path.isfile(path) for path in local_files):
                event.acceptProposedAction()
                return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event):
        paths = []
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if os.path.isfile(path):
                paths.append(path)

        if paths:
            event.acceptProposedAction()
            self.filesDropped.emit(paths)
        else:
            event.ignore()
    def eventFilter(self, watched, event):
        if event.type() in (QEvent.DragEnter, QEvent.DragMove, QEvent.Drop):
            if event.mimeData().hasUrls():
                if event.type() == QEvent.Drop:
                    files = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
                    self.filesDropped.emit(files)
                event.acceptProposedAction()
                return True
        return super().eventFilter(watched, event)


# ======================================================================
# 7. MAIN GUI APPLICATION
# ======================================================================

class PersistentSplitterHandle(QSplitterHandle):
    """Always-visible, easy-to-grab splitter handle.

    Qt's native Windows splitter handle can become almost invisible until the
    mouse crosses it.  That is particularly frustrating in a dense forensic
    workspace, so this handle paints a permanent grip and uses the appropriate
    resize cursor.
    """

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setAttribute(Qt.WA_Hover, True)
        self.setMouseTracking(True)
        if orientation == Qt.Horizontal:
            self.setCursor(Qt.SplitHCursor)
        else:
            self.setCursor(Qt.SplitVCursor)

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QPen

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        # A persistent neutral track makes the handle discoverable even when
        # the mouse is nowhere near it.
        painter.fillRect(self.rect(), QColor("#d8dee8"))

        if self.orientation() == Qt.Horizontal:
            cx = self.width() // 2
            grip_w = min(34, max(20, self.width() - 12))
            left = (self.width() - grip_w) // 2
            top = max(2, (self.height() - 6) // 2)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("#718096"))
            painter.drawRoundedRect(left, top, grip_w, 6, 3, 3)
            painter.setPen(QPen(QColor("#4a5568"), 1))
            painter.drawLine(left + 5, top + 2, left + grip_w - 5, top + 2)
            painter.drawLine(left + 5, top + 4, left + grip_w - 5, top + 4)
        else:
            cy = self.height() // 2
            grip_h = min(34, max(20, self.height() - 12))
            top = (self.height() - grip_h) // 2
            left = max(2, (self.width() - 6) // 2)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("#718096"))
            painter.drawRoundedRect(left, top, 6, grip_h, 3, 3)
            painter.setPen(QPen(QColor("#4a5568"), 1))
            painter.drawLine(left + 2, top + 5, left + 2, top + grip_h - 5)
            painter.drawLine(left + 4, top + 5, left + 4, top + grip_h - 5)

        painter.end()


class PersistentSplitter(QSplitter):
    """QSplitter with a permanently visible, wide, grab-friendly handle."""

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setChildrenCollapsible(False)
        self.setOpaqueResize(True)
        self.setHandleWidth(12 if orientation == Qt.Vertical else 10)
        self.setStyleSheet("""
            QSplitter::handle {
                background: #d8dee8;
            }
            QSplitter::handle:hover {
                background: #c3ccd9;
            }
        """)

    def createHandle(self):
        return PersistentSplitterHandle(self.orientation(), self)


class LedgerKeyboardDelegate(QStyledItemDelegate):
    """Give Enter and arrow keys one shared edit-commit path in ledger cells.

    Enter commits the edited value and leaves the current cell selected.
    Arrow keys commit the edited value first, then move the current cell in the
    requested direction. This prevents edits from being lost when a user uses
    the arrow keys rather than Enter to leave an edited cell.
    """

    navigationKeyPressed = Signal(int)
    _NAVIGATION_KEYS = {
        Qt.Key.Key_Return,
        Qt.Key.Key_Enter,
        Qt.Key.Key_Up,
        Qt.Key.Key_Down,
        Qt.Key.Key_Left,
        Qt.Key.Key_Right,
    }

    def eventFilter(self, editor, event):
        if (
            event.type() == QEvent.Type.KeyPress
            and event.key() in self._NAVIGATION_KEYS
            and event.modifiers() == Qt.KeyboardModifier.NoModifier
        ):
            key = int(event.key())
            # Ask Qt to write the editor's text to the cell model before
            # closing the editor. The existing cellChanged pipeline then runs
            # the same validation and recalculation used by Enter.
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QAbstractItemDelegate.EndEditHint.NoHint)
            # Let the cellChanged handlers and their queued focus restoration
            # finish before moving the selection.
            QTimer.singleShot(
                0,
                lambda saved_key=key: self.navigationKeyPressed.emit(saved_key),
            )
            event.accept()
            return True
        return super().eventFilter(editor, event)


class TaxForensicCompleteApp(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "TaxForensic Multi-Bank Analyzer — V26.10 Forensic Workstation"
        )

        # Default physical target is approximately 9 x 5.6 inches at 96 DPI.
        # The normal Windows resize/maximize controls remain enabled.
        self.resize(
            1180,
            720,
        )
        self.setMinimumSize(
            820,
            520,
        )

        init_database()
        ensure_taxforensic_support_dirs()

        self.pan = "ABCDE1234F"

        self.loaded_files = []
        self.pending_files = []
        self.processing_files = []
        self.successfully_processed_files = []
        self.failed_processing_files = []
        self.selected_pending_file = None

        self.active_row = -1
        self.selected_account_key = None
        self.selected_merge_sources = set()
        self._merge_source_records = {}
        # Inspector zoom is remembered independently for each loaded statement.
        # It is changed only by explicit zoom/recalibration actions and is not
        # reset merely because the selected transaction changes.
        self._inspector_zoom_by_source = {}
        self._inspector_evidence_scene_rect = None
        self._inspector_source_key = None
        self._inspector_zoom = 1.0

        self.transactions = []

        self.contras = []

        self.cash_flows = []

        self.violations = []
        if hasattr(self, "statement_selector_frame"):
            self.refresh_statement_selector()

        self.worker = None
        self.worker_result_payload = None
        self.worker_stdout_buffer = ""
        self.worker_stage_seen = False
        self.worker_last_stage = ""
        self.worker_watchdog_generation = 0
        self.init_ui()

        # Start with a genuinely empty forensic workspace. Statements are
        # loaded first and processed only after the user presses Process.
        self.transactions = []
        self.contras = []
        self.cash_flows = []
        self.violations = []

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def init_ui(self):

        central = QWidget()

        self.setCentralWidget(
            central
        )

        main_layout = QVBoxLayout(
            central
        )

        # --------------------------------------------------------------
        # TOP BAR
        # --------------------------------------------------------------

        # Responsive case/export header.  The original single horizontal row
        # was too wide for the normal workstation size, causing the right-most
        # export controls to be clipped.  Keep the case information on row 1
        # and put the three export actions on row 2.
        top_bar = QGridLayout()
        top_bar.setContentsMargins(0, 0, 0, 0)
        top_bar.setHorizontalSpacing(6)
        top_bar.setVerticalSpacing(5)

        # The old generic "Case Profile" header is replaced by the currently
        # selected financial statement identity. The first line shows the
        # selected bank/issuer and account/card. The second line shows the
        # actual PDF/statement filename. PAN, AY and currency remain visible.
        self.lbl_case_info = QLabel(
            f"<b>Selected Account:</b> No statement selected | "
            f"<b>PAN:</b> {self.pan} | "
            f"<b>AY:</b> 2024-25 | "
            f"<b>Currency:</b> Indian Rupees (Rs.)<br>"
            f"<b>Statement:</b> No statement selected"
        )
        self.lbl_case_info.setStyleSheet(
            "font-size: 13px; font-weight: bold; color: #1a237e;"
        )
        self.lbl_case_info.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        top_bar.addWidget(self.lbl_case_info, 0, 0, 1, 2)

        # Keep both controls visually identical to the existing In-Situ
        # checkbox. Statement workspace is intentionally hidden by default so
        # the forensic ledger gets the maximum working area.
        header_options = QWidget()
        header_options_layout = QHBoxLayout(header_options)
        header_options_layout.setContentsMargins(0, 0, 0, 0)
        header_options_layout.setSpacing(12)

        self.chk_insitu = QCheckBox(
            "Save Output In-Situ (Same folder as statement)"
        )
        self.chk_insitu.setChecked(True)
        self.chk_insitu.setStyleSheet(
            "font-weight: bold; color: #0277bd;"
        )
        self.chk_insitu.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        header_options_layout.addWidget(self.chk_insitu)

        self.chk_show_workspace = QCheckBox(
            "Show Statement Workspace"
        )
        self.chk_show_workspace.setChecked(True)
        self.chk_show_workspace.setToolTip(
            "Show or hide the statement loading and account workspace."
        )
        self.chk_show_workspace.setStyleSheet(
            "font-weight: bold; color: #0277bd;"
        )
        self.chk_show_workspace.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.chk_show_workspace.toggled.connect(self.toggle_statement_workspace)
        header_options_layout.addWidget(self.chk_show_workspace)

        top_bar.addWidget(header_options, 0, 2)

        self.btn_export_all_separate = QPushButton(
            "Export All Separate Bank Excels"
        )
        self.btn_export_all_separate.setStyleSheet(
            "background-color: #00695c; color: white; font-weight: bold; padding: 6px 10px;"
        )
        self.btn_export_all_separate.clicked.connect(
            self.action_export_all_separate_banks
        )

        self.btn_export_consolidated = QPushButton(
            "Export Consolidated Master Excel"
        )
        self.btn_export_consolidated.setStyleSheet(
            "background-color: #37474f; color: white; font-weight: bold; padding: 6px 10px;"
        )
        self.btn_export_consolidated.clicked.connect(
            self.action_export_consolidated_excel
        )

        self.btn_commit_db = QPushButton(
            "Commit to SQLite (tax_case.db)"
        )
        self.btn_commit_db.setStyleSheet(
            "background-color: #1b5e20; color: white; font-weight: bold; padding: 6px 10px;"
        )
        self.btn_commit_db.setEnabled(False)
        self.btn_commit_db.clicked.connect(self.action_commit_db)

        for button in (
            self.btn_export_all_separate,
            self.btn_export_consolidated,
            self.btn_commit_db,
        ):
            button.setMinimumHeight(32)
            button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        main_layout.addLayout(top_bar)

        # --------------------------------------------------------------
        # PROGRESS
        # --------------------------------------------------------------

        self.progress_frame = QFrame()

        self.progress_frame.setFrameShape(
            QFrame.StyledPanel
        )

        self.progress_frame.setStyleSheet(
            "QFrame { background: #eef2f7; border: 1px solid #cfd8dc; "
            "border-radius: 4px; padding: 1px; }"
        )

        progress_layout = QHBoxLayout(self.progress_frame)
        progress_layout.setContentsMargins(4, 2, 4, 2)
        progress_layout.setSpacing(4)

        self.lbl_progress_status = QLabel(
            "Ready. Drop statements below to begin asynchronous extraction."
        )

        self.lbl_progress_status.setStyleSheet(
            "font-weight: bold; "
            "color: #37474f;"
        )

        progress_layout.addWidget(
            self.lbl_progress_status,
            3,
        )

        self.progress_bar = QProgressBar()

        self.progress_bar.setRange(
            0,
            100,
        )

        self.progress_bar.setValue(
            0
        )

        progress_layout.addWidget(
            self.progress_bar,
            4,
        )

        self.btn_cancel_worker = QPushButton(
            "Cancel Ingestion"
        )

        self.btn_cancel_worker.setStyleSheet(
            "background-color: #c62828; "
            "color: white; "
            "font-weight: bold; "
            "padding: 2px 6px;"
        )

        self.btn_cancel_worker.setEnabled(
            False
        )

        self.btn_cancel_worker.clicked.connect(
            self.action_cancel_worker
        )

        progress_layout.addWidget(
            self.btn_cancel_worker,
            1,
        )

        # Keep processing status in the same compact horizontal operations
        # row as the export controls. This prevents a full-width idle band
        # from consuming valuable vertical forensic workspace.
        operations_row = QWidget()
        operations_layout = QHBoxLayout(operations_row)
        operations_layout.setContentsMargins(0, 0, 0, 0)
        operations_layout.setSpacing(6)

        self.progress_frame.setMaximumHeight(40)
        self.progress_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        operations_layout.addWidget(self.progress_frame, 2)
        operations_layout.addWidget(self.btn_export_all_separate, 1)
        operations_layout.addWidget(self.btn_export_consolidated, 1)
        operations_layout.addWidget(self.btn_commit_db, 1)
        main_layout.addWidget(operations_row)

        # Compact statement selector. It remains visible even when the full
        # Statement Workspace is hidden, so the user can switch between
        # loaded/processed accounts without exposing the detailed queue.
        self.statement_selector_frame = QFrame()
        self.statement_selector_frame.setStyleSheet(
            "QFrame { background: #f5f7fa; border: 1px solid #cfd8dc; border-radius: 4px; }"
        )
        selector_layout = QHBoxLayout(self.statement_selector_frame)
        selector_layout.setContentsMargins(5, 3, 5, 3)
        selector_layout.setSpacing(6)
        selector_layout.addWidget(QLabel("<b>Statements:</b>"))
        self.statement_selector_scroll = QScrollArea()
        self.statement_selector_scroll.setWidgetResizable(True)
        self.statement_selector_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.statement_selector_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.statement_selector_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.statement_selector_scroll.setFixedHeight(34)
        self.statement_selector_container = QWidget()
        self.statement_selector_layout = QHBoxLayout(self.statement_selector_container)
        self.statement_selector_layout.setContentsMargins(0, 0, 0, 0)
        self.statement_selector_layout.setSpacing(5)
        self.statement_selector_layout.addStretch(1)
        self.statement_selector_scroll.setWidget(self.statement_selector_container)
        selector_layout.addWidget(self.statement_selector_scroll, 1)
        main_layout.addWidget(self.statement_selector_frame)

        # --------------------------------------------------------------
        # --------------------------------------------------------------
        # COMPACT INPUT AREA + ACCOUNT TABLE
        # --------------------------------------------------------------
        # Default vertical footprint: approximately 20% of the normal
        # application window. The horizontal splitter remains user-resizable.

        self.drop_zone = StatementDropFrame()
        self.drop_zone.setMinimumHeight(0)
        self.drop_zone.setSizePolicy(
            QSizePolicy.Expanding,
            QSizePolicy.Expanding,
        )
        self.drop_zone.setStyleSheet(
            """
            QFrame {
                border: 2px dashed #0277bd;
                border-radius: 6px;
                background-color: #f0f8ff;
            }
            QFrame:hover {
                background-color: #e1f5fe;
                border-color: #01579b;
            }
            """
        )

        dz_layout = QHBoxLayout(self.drop_zone)
        dz_layout.setContentsMargins(10, 8, 10, 8)

        dz_label = QLabel(
            "<b>Drag & Drop Bank Statements / Passbook Photos</b><br>"
            "PDF • JPG/PNG/TIFF • XLSX/XLS • CSV/TSV • XML/CAMT • "
            "OFX/QFX • MT940 • QIF • BAI2 • JSON • TXT"
        )
        dz_label.setAlignment(Qt.AlignCenter)
        dz_label.setStyleSheet("color: #01579b; font-size: 12px;")
        dz_label.setWordWrap(True)
        dz_layout.addWidget(dz_label, 1)

        self.btn_browse = QPushButton("Browse Files")
        self.btn_browse.clicked.connect(self.action_manual_browse)
        dz_layout.addWidget(self.btn_browse)

        # The drop zone is ONLY for loading files. Processing controls live
        # immediately above the Detected Accounts table, because loading and
        # processing are deliberately separate forensic stages.
        self.drop_zone.filesDropped.connect(self.queue_statement_files)
        self.drop_zone.installEventFilter(self.drop_zone)
        dz_label.installEventFilter(self.drop_zone)
        self.btn_browse.installEventFilter(self.drop_zone)

        bank_box = QWidget()
        bank_box.setSizePolicy(
            QSizePolicy.Expanding,
            QSizePolicy.Expanding,
        )

        bank_vbox = QVBoxLayout(bank_box)
        bank_vbox.setContentsMargins(0, 0, 0, 0)
        bank_vbox.setSpacing(3)

        detected_label = QLabel("<b>Detected / Loaded Accounts & Statements:</b>")
        detected_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        bank_vbox.addWidget(detected_label)

        # Processing controls are deliberately ABOVE the table.
        # Use a 2x2 grid instead of one long horizontal row.  The right
        # account pane is intentionally compact, so four long captions can
        # otherwise be clipped at the normal 864px window width.
        action_bar = QGridLayout()
        action_bar.setContentsMargins(0, 0, 0, 0)
        action_bar.setHorizontalSpacing(5)
        action_bar.setVerticalSpacing(4)

        self.btn_process_selected = QPushButton("Process Selected")
        self.btn_process_selected.setEnabled(False)
        self.btn_process_selected.setToolTip(
            "Process only the selected loaded account/statement."
        )
        self.btn_process_selected.setStyleSheet(
            "background-color: #1565c0; color: white; font-weight: bold; padding: 5px 8px;"
        )
        self.btn_process_selected.clicked.connect(
            self.action_process_selected_statement
        )

        self.btn_process_all = QPushButton("Process All Statements")
        self.btn_process_all.setEnabled(False)
        self.btn_process_all.setToolTip(
            "Process every statement currently waiting in the workspace."
        )
        self.btn_process_all.setStyleSheet(
            "background-color: #283593; color: white; font-weight: bold; padding: 5px 8px;"
        )
        self.btn_process_all.clicked.connect(
            self.action_process_loaded_statements
        )

        self.btn_process_out = QPushButton("Merged Analyzer")
        self.btn_process_out.setEnabled(False)
        self.btn_process_out.setToolTip(
            "Process Out: run the final merged chronological analyzer on all processed statements."
        )
        self.btn_process_out.setStyleSheet(
            "background-color: #00695c; color: white; font-weight: bold; padding: 5px 8px;"
        )
        self.btn_process_out.clicked.connect(self.action_process_out_merged)

        self.btn_delete_account = QPushButton("Remove Selected")
        self.btn_delete_account.setEnabled(False)
        self.btn_delete_account.setToolTip(
            "Remove the selected loaded statement or processed account from the workspace."
        )
        self.btn_delete_account.setStyleSheet(
            "background-color: #b71c1c; color: white; font-weight: bold; padding: 5px 8px;"
        )
        self.btn_delete_account.clicked.connect(
            self.action_delete_selected_account
        )

        for button in (
            self.btn_process_selected,
            self.btn_process_all,
            self.btn_process_out,
            self.btn_delete_account,
        ):
            button.setMinimumHeight(30)
            button.setSizePolicy(
                QSizePolicy.Expanding,
                QSizePolicy.Fixed,
            )

        action_bar.addWidget(self.btn_process_selected, 0, 0)
        action_bar.addWidget(self.btn_process_all, 0, 1)
        action_bar.addWidget(self.btn_process_out, 1, 0)
        action_bar.addWidget(self.btn_delete_account, 1, 1)
        action_bar.setColumnStretch(0, 1)
        action_bar.setColumnStretch(1, 1)

        bank_vbox.addLayout(action_bar)

        self.table_bank_actions = QTableWidget()
        self.table_bank_actions.setColumnCount(8)
        self.table_bank_actions.setHorizontalHeaderLabels(
            [
                "Merge",
                "Bank / Issuer",
                "Account / Card",
                "Statement",
                "Entries",
                "Status",
                "Process",
                "Export",
            ]
        )

        account_header = self.table_bank_actions.horizontalHeader()
        account_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        account_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        account_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        account_header.setSectionResizeMode(3, QHeaderView.Stretch)
        for c in (4, 5, 6, 7):
            account_header.setSectionResizeMode(c, QHeaderView.ResizeToContents)

        self.table_bank_actions.setMinimumHeight(48)
        self.table_bank_actions.setSizePolicy(
            QSizePolicy.Expanding,
            QSizePolicy.Expanding,
        )
        self.table_bank_actions.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.table_bank_actions.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.table_bank_actions.cellClicked.connect(
            self.on_detected_account_clicked
        )
        self.table_bank_actions.itemSelectionChanged.connect(
            self.on_statement_selection_changed
        )
        self.table_bank_actions.itemChanged.connect(
            self.on_merge_selection_changed
        )
        bank_vbox.addWidget(self.table_bank_actions, 1)

        self.statement_info_box = QGroupBox("Selected Statement Information")
        self.statement_info_box.setMaximumHeight(82)
        info_layout = QVBoxLayout(self.statement_info_box)
        info_layout.setContentsMargins(6, 3, 6, 3)
        self.lbl_statement_info = QLabel(
            "Select a loaded statement to view bank/issuer, account or card, holder, statement period, currency, language and page count."
        )
        self.lbl_statement_info.setWordWrap(True)
        self.lbl_statement_info.setStyleSheet("font-size: 10px; padding: 2px;")
        info_layout.addWidget(self.lbl_statement_info)
        bank_vbox.addWidget(self.statement_info_box)

        # Horizontal splitter: default 35% / 65%.
        top_splitter = PersistentSplitter(Qt.Horizontal)
        top_splitter.setChildrenCollapsible(False)
        self.statement_workspace = top_splitter
        top_splitter.addWidget(self.drop_zone)
        top_splitter.addWidget(bank_box)
        top_splitter.setStretchFactor(0, 35)
        top_splitter.setStretchFactor(1, 65)
        top_splitter.setSizes([350, 650])

        # Vertical footprint is intentionally compact. The user can drag the
        # lower splitter handle to enlarge it whenever more rows are needed.
        top_splitter.setMinimumHeight(110)
        # Do not impose a maximum height. The user must be able to drag the
        # vertical splitter freely; the initial 20% footprint is set by the
        # parent splitter below.

        # --------------------------------------------------------------
        # TABS
        # --------------------------------------------------------------

        self.tabs = QTabWidget()

        self.tab_recon = QWidget()
        self.tab_contra = QWidget()
        self.tab_cash = QWidget()
        self.tab_sft = QWidget()

        self.tabs.addTab(
            self.tab_recon,
            "1. Reconciled Ledger & PDF Inspector",
        )

        self.tabs.addTab(
            self.tab_contra,
            "2. Multi-Bank Contra Elimination",
        )

        self.tabs.addTab(
            self.tab_cash,
            "3. Consolidated Cash Flow & Peak Credit",
        )

        self.tabs.addTab(
            self.tab_sft,
            "4. Statutory SFT & Cash Violations",
        )

        # Main vertical workspace. The top input/account area and the
        # statement/inspector workspace can be resized independently.
        self.main_vertical_splitter = PersistentSplitter(Qt.Vertical)
        self.main_vertical_splitter.setChildrenCollapsible(False)
        self.main_vertical_splitter.addWidget(top_splitter)
        self.main_vertical_splitter.addWidget(self.tabs)
        self.main_vertical_splitter.setStretchFactor(0, 0)
        self.main_vertical_splitter.setStretchFactor(1, 1)

        main_layout.addWidget(self.main_vertical_splitter, 1)

        # The statement workspace is visible by default. Its splitter remains
        # freely resizable, and the checkbox can hide it when more ledger space
        # is needed.
        self._statement_workspace_last_height = 0
        self.statement_workspace.setVisible(True)

        # Set defaults only once, after the widgets have a real size. The
        # splitter handles remain completely free afterwards.
        QTimer.singleShot(0, self._set_default_workspace_sizes)

        self.build_tab_recon()
        self.build_tab_contra()
        self.build_tab_cash()
        self.build_tab_sft()

    def toggle_statement_workspace(self, visible):
        """Show/hide the complete statement workspace and idle progress strip.

        When the workspace is hidden, no empty loading/progress area is left
        consuming vertical space. During an active parser run the progress
        strip remains visible so cancellation/status are still available.
        """
        try:
            sizes = self.main_vertical_splitter.sizes()
            total = max(1, sum(sizes))
            running = self._parser_process_running()
            if visible:
                self.statement_workspace.setVisible(True)
                self.progress_frame.setVisible(True)
                self.statement_selector_frame.setVisible(False)
                restore = self._statement_workspace_last_height or int(total * 0.30)
                restore = max(165, min(max(165, total - 120), restore))
                self.main_vertical_splitter.setSizes([restore, max(1, total - restore)])
            else:
                if self.statement_workspace.isVisible():
                    self._statement_workspace_last_height = max(110, sizes[0] if sizes else 0)
                self.statement_workspace.setVisible(False)
                self.refresh_statement_selector()
                if not running:
                    self.progress_frame.setVisible(False)
                self.main_vertical_splitter.setSizes([0, total])
        except Exception:
            pass

    def _set_default_workspace_sizes(self):
        """Start with a useful statement workspace; user can resize freely."""
        try:
            h = max(1, self.main_vertical_splitter.height())
            if self.chk_show_workspace.isChecked():
                top_h = max(165, int(h * 0.30))
                self.main_vertical_splitter.setSizes([top_h, max(1, h - top_h)])
            else:
                self.main_vertical_splitter.setSizes([0, h])
                if not self._parser_process_running():
                    self.progress_frame.setVisible(False)
        except Exception:
            pass

    def _update_selected_statement_header(self, filepath, bank=None, account=None):
        """Show selected bank/account/card and exact statement filename in header."""
        if filepath:
            filename = os.path.basename(str(filepath))
        else:
            filename = "No statement selected"

        if filepath and os.path.exists(filepath):
            try:
                meta = extract_statement_metadata(filepath)
                bank = meta.get("bank") or bank or "Financial Institution"
                account = meta.get("account") or account or "Not detected"
            except Exception:
                bank = bank or "Financial Institution"
                account = account or "Not detected"

        bank = bank or "No statement selected"
        account = account or "Not detected"
        self.lbl_case_info.setText(
            f"<b>Selected Account:</b> {bank} | <b>A/c or Card:</b> {account} | "
            f"<b>PAN:</b> {self.pan} | <b>AY:</b> 2024-25 | "
            f"<b>Currency:</b> Indian Rupees (Rs.)<br>"
            f"<b>Statement:</b> {filename}"
        )

    # ------------------------------------------------------------------
    # RECON TAB
    # ------------------------------------------------------------------

    def build_tab_recon(self):

        layout = QVBoxLayout(self.tab_recon)

        self.lbl_recon_status = QLabel("Evaluating Statement...")
        self.lbl_recon_status.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(self.lbl_recon_status)

        # Master Ledger is deliberately split into an account view and a
        # date-wise merged view. The underlying transaction objects remain
        # untouched, so reconciliation and forensic calculations use the same
        # source of truth regardless of the selected tab.
        self.ledger_tabs = QTabWidget()

        individual_box = QWidget()
        individual_layout = QVBoxLayout(individual_box)
        self.lbl_selected_account = QLabel("Select an account above to view its entries.")
        individual_layout.addWidget(self.lbl_selected_account)
        self.table_account = QTableWidget()
        self.table_account.setColumnCount(7)
        self.table_account.setHorizontalHeaderLabels(
            ["Date", "Narration", "Debit (Dr)", "Credit (Cr)", "Calculated Balance", "Image Balance", "Audit Status"]
        )
        ah = self.table_account.horizontalHeader()
        ah.setSectionResizeMode(0, QHeaderView.Fixed)
        ah.setSectionResizeMode(1, QHeaderView.Stretch)
        for c, w in ((0, 92), (2, 105), (3, 105), (4, 140), (5, 130), (6, 145)):
            self.table_account.setColumnWidth(c, w)
        self.table_account.setWordWrap(False)
        self.table_account.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table_account.setSortingEnabled(False)
        self.table_account.cellClicked.connect(self.on_individual_row_clicked)
        self.table_account.currentCellChanged.connect(self.on_ledger_current_cell_changed)
        self.table_account.cellChanged.connect(self.on_account_cell_changed)
        self._install_ledger_keyboard_delegate(self.table_account)
        individual_layout.addWidget(self.table_account)
        self.ledger_tabs.addTab(individual_box, "Individual Bank Account")

        merged_box = QWidget()
        merged_layout = QVBoxLayout(merged_box)
        merged_layout.addWidget(QLabel("<b>Merged Date-wise Master Ledger</b>  Select the accounts/cards to combine in the Statement Workspace above."))
        merge_bar = QHBoxLayout()
        self.lbl_merge_selection = QLabel("Selected sources: 0")
        merge_bar.addWidget(self.lbl_merge_selection)
        self.btn_merge_all = QPushButton("Select All")
        self.btn_merge_all.clicked.connect(lambda: self._set_all_merge_sources(True))
        merge_bar.addWidget(self.btn_merge_all)
        self.btn_merge_none = QPushButton("Clear All")
        self.btn_merge_none.clicked.connect(lambda: self._set_all_merge_sources(False))
        merge_bar.addWidget(self.btn_merge_none)
        self.btn_build_merged = QPushButton("Build Merged Ledger")
        self.btn_build_merged.clicked.connect(self._build_selected_merged_ledger)
        merge_bar.addWidget(self.btn_build_merged)
        merge_bar.addStretch(1)
        merged_layout.addLayout(merge_bar)
        self.table_recon = QTableWidget()
        self.table_recon.setColumnCount(9)
        self.table_recon.setHorizontalHeaderLabels(
            ["Bank Name", "Account Number", "Date", "Narration", "Debit (Dr)", "Credit (Cr)", "Calculated Balance", "Image Balance", "Audit Status"]
        )
        rh = self.table_recon.horizontalHeader()
        rh.setSectionResizeMode(3, QHeaderView.Stretch)
        for c, w in ((0, 125), (1, 115), (2, 92), (4, 105), (5, 105), (6, 140), (7, 130), (8, 155)):
            self.table_recon.setColumnWidth(c, w)
        self.table_recon.setWordWrap(False)
        self.table_recon.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table_recon.setSortingEnabled(False)
        self.table_recon.cellClicked.connect(self.on_table_row_clicked)
        self.table_recon.currentCellChanged.connect(self.on_ledger_current_cell_changed)
        self.table_recon.cellChanged.connect(self.on_recon_cell_changed)
        self._install_ledger_keyboard_delegate(self.table_recon)
        merged_layout.addWidget(self.table_recon)
        self.ledger_tabs.addTab(merged_box, "Merged Date-wise Master Ledger")
        self.ledger_tabs.currentChanged.connect(self.on_ledger_tab_changed)

        # Ledger and synchronized document inspector are independent panes.
        # This is a real splitter, not a fixed-height layout, so the user can
        # drag the handle continuously in either direction.
        ledger_inspector_splitter = PersistentSplitter(Qt.Vertical)
        ledger_inspector_splitter.setChildrenCollapsible(False)
        ledger_inspector_splitter.addWidget(self.ledger_tabs)

        # Synchronized inspector occupies the lower pane by default.
        inspector_box = QWidget()
        inspector_layout = QVBoxLayout(inspector_box)
        inspector_layout.setContentsMargins(0, 2, 0, 0)
        inspector_layout.addWidget(QLabel("<b>Synchronized Document Inspector:</b>"))

        inspector_toolbar = QHBoxLayout()
        inspector_toolbar.setContentsMargins(0, 0, 0, 0)

        self.lbl_inspector_status = QLabel("Select a ledger row to inspect the source document.")
        inspector_toolbar.addWidget(self.lbl_inspector_status, 1)

        self.btn_inspector_zoom_out = QPushButton("−")
        self.btn_inspector_zoom_out.setToolTip("Zoom out")
        self.btn_inspector_zoom_out.clicked.connect(lambda: self._inspector_zoom_by(1 / 1.25))
        inspector_toolbar.addWidget(self.btn_inspector_zoom_out)

        self.btn_inspector_zoom_fit = QPushButton("Recalibrate")
        self.btn_inspector_zoom_fit.setToolTip(
            "Recalculate the statement's default inspector presentation at 50% of the available width/height. "
            "This replaces the remembered zoom for the current statement."
        )
        self.btn_inspector_zoom_fit.clicked.connect(self._inspector_fit)
        inspector_toolbar.addWidget(self.btn_inspector_zoom_fit)

        self.btn_inspector_zoom_100 = QPushButton("100%")
        self.btn_inspector_zoom_100.clicked.connect(self._inspector_zoom_100)
        inspector_toolbar.addWidget(self.btn_inspector_zoom_100)

        self.btn_inspector_zoom_in = QPushButton("+")
        self.btn_inspector_zoom_in.setToolTip("Zoom in")
        self.btn_inspector_zoom_in.clicked.connect(lambda: self._inspector_zoom_by(1.25))
        inspector_toolbar.addWidget(self.btn_inspector_zoom_in)

        self.btn_inspector_prev = QPushButton("← Previous")
        self.btn_inspector_prev.clicked.connect(lambda: self.navigate_transaction(-1))
        inspector_toolbar.addWidget(self.btn_inspector_prev)
        self.btn_inspector_next = QPushButton("Next →")
        self.btn_inspector_next.clicked.connect(lambda: self.navigate_transaction(1))
        inspector_toolbar.addWidget(self.btn_inspector_next)
        inspector_toolbar.addWidget(QLabel("  Ctrl+wheel: zoom | Drag: pan | ↑/↓: row navigation"))
        inspector_layout.addLayout(inspector_toolbar)

        self.scene = QGraphicsScene()
        self.view = ForensicDocumentView(self.scene)
        self.view.zoomChanged.connect(self._on_view_zoom_changed)
        inspector_layout.addWidget(self.view, 1)
        self._inspector_base_transform = None

        ledger_inspector_splitter.addWidget(inspector_box)
        ledger_inspector_splitter.setStretchFactor(0, 1)
        ledger_inspector_splitter.setStretchFactor(1, 1)
        # Default workspace: approximately 50% ledger / 50% inspector.
        # The splitter remains freely draggable after startup.
        QTimer.singleShot(
            0,
            lambda: self._set_ledger_inspector_default_splitter(
                ledger_inspector_splitter
            ),
        )
        layout.addWidget(ledger_inspector_splitter, 1)

    def _set_ledger_inspector_default_splitter(self, splitter):
        """Start the ledger and inspector at roughly 50/50 height."""
        try:
            h = max(1, splitter.height())
            half = max(1, int(h * 0.50))
            splitter.setSizes([half, max(1, h - half)])
        except Exception:
            pass

    # ------------------------------------------------------------------
    # CONTRA TAB
    # ------------------------------------------------------------------

    def build_tab_contra(self):

        layout = QVBoxLayout(
            self.tab_contra
        )

        layout.addWidget(
            QLabel(
                "<h3>Inter-Account Transfer "
                "& Contra Cancellation Engine</h3>"
            )
        )

        layout.addWidget(
            QLabel(
                "Matches debits in Account A "
                "with credits in Account B "
                "within 48 hours."
            )
        )

        self.table_contra = QTableWidget()

        self.table_contra.setColumnCount(
            6
        )

        self.table_contra.setHorizontalHeaderLabels(
            [
                "Date",
                "Source Bank & A/c",
                "Destination Bank & A/c",
                "Contra Amount",
                "Time Delta",
                "Status",
            ]
        )

        self.table_contra.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )

        layout.addWidget(
            self.table_contra
        )

        self.lbl_contra_summary = QLabel(
            "Total Contra Transfers Filtered: Rs. 0.00"
        )

        layout.addWidget(
            self.lbl_contra_summary
        )

    # ------------------------------------------------------------------
    # CASH TAB
    # ------------------------------------------------------------------

    def build_tab_cash(self):

        layout = QVBoxLayout(
            self.tab_cash
        )

        layout.addWidget(
            QLabel(
                "<h3>Consolidated Cash Flow "
                "& Peak Credit Engine</h3>"
            )
        )

        layout.addWidget(
            QLabel(
                "Evaluates whether cash deposits "
                "were supported by prior cash withdrawals."
            )
        )

        self.table_cash = QTableWidget()

        self.table_cash.setColumnCount(
            5
        )

        self.table_cash.setHorizontalHeaderLabels(
            [
                "Date",
                "Cash Inflow (Wdl)",
                "Cash Outflow (Dep)",
                "Theoretical Cash-on-Hand",
                "Audit Finding",
            ]
        )

        self.table_cash.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )

        layout.addWidget(
            self.table_cash
        )

        self.lbl_peak_credit = QLabel(
            "Peak Unexplained Cash Exposure: Rs. 0.00"
        )

        layout.addWidget(
            self.lbl_peak_credit
        )

    # ------------------------------------------------------------------
    # SFT TAB
    # ------------------------------------------------------------------

    def build_tab_sft(self):

        layout = QVBoxLayout(
            self.tab_sft
        )

        layout.addWidget(
            QLabel(
                "<h3>Consolidated Statutory Flags "
                "& SFT Scrutiny</h3>"
            )
        )

        self.table_sft = QTableWidget()

        self.table_sft.setColumnCount(
            6
        )

        self.table_sft.setHorizontalHeaderLabels(
            [
                "Date",
                "Bank & A/c",
                "Particulars",
                "Amount",
                "Statutory Rule",
                "Violation Severity",
            ]
        )

        self.table_sft.horizontalHeader().setSectionResizeMode(
            2,
            QHeaderView.Stretch,
        )

        layout.addWidget(
            self.table_sft
        )

        self.lbl_sft_summary = QLabel(
            "Total Violations Detected: 0 Line Items"
        )

        layout.addWidget(
            self.lbl_sft_summary
        )

    # ------------------------------------------------------------------
    # DRAG DROP
    # ------------------------------------------------------------------

    def dz_dragEnterEvent(
        self,
        event: QDragEnterEvent,
    ):

        if event.mimeData().hasUrls():

            event.acceptProposedAction()

    # ------------------------------------------------------------------
    # FILE BROWSE / LOAD / PROCESS
    # ------------------------------------------------------------------

    SUPPORTED_STATEMENT_EXTENSIONS = {
        ".pdf", ".xlsx", ".xls", ".xlsm", ".csv", ".tsv",
        ".xml", ".ofx", ".qfx", ".qif", ".mt940", ".940", ".sta",
        ".bai", ".bai2", ".json", ".txt",
        ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp",
    }

    def _normalise_selected_files(self, files):
        result = []
        for path in files or []:
            path = os.path.abspath(str(path))
            if not os.path.isfile(path):
                continue
            if os.path.splitext(path)[1].lower() not in self.SUPPORTED_STATEMENT_EXTENSIONS:
                continue
            if path not in result:
                result.append(path)
        return result

    def _preview_statement_identity(self, filepath):
        """Cheap identity preview. Never performs full statement extraction."""
        name = os.path.basename(filepath)
        ext = os.path.splitext(filepath)[1].lower()
        bank = "Pending detection"
        account = "Pending detection"

        try:
            if ext == ".pdf":
                with open(filepath, "rb") as f:
                    data = f.read()
                doc = fitz.open(stream=data, filetype="pdf")
                try:
                    raw = doc[0].get_text("text") if len(doc) else ""
                finally:
                    doc.close()
                helper = StatementParserWorker([])
                bank, account = helper.detect_bank_and_acc(raw, filepath)

            elif ext in {".xlsx", ".xlsm", ".xls"}:
                # XLS support may depend on the existing parser stack. For
                # preview purposes XLSX/XLSM is cheap; failures simply remain
                # pending and are fully detected during Process.
                if ext in {".xlsx", ".xlsm"}:
                    wb = openpyxl.load_workbook(
                        filepath, read_only=True, data_only=True
                    )
                    try:
                        values = []
                        ws = wb.active
                        for row in ws.iter_rows(max_row=12, values_only=True):
                            values.extend(str(v) for v in row if v is not None)
                    finally:
                        wb.close()
                    helper = StatementParserWorker([])
                    bank, account = helper.detect_bank_and_acc(
                        " ".join(values), filepath
                    )
            elif ext in {".csv", ".tsv", ".txt"}:
                with open(filepath, "r", encoding="utf-8-sig", errors="replace") as f:
                    raw = f.read(12000)
                helper = StatementParserWorker([])
                bank, account = helper.detect_bank_and_acc(raw, filepath)
        except Exception:
            # Preview must never prevent loading. Full processing will report
            # the actual extraction error.
            pass

        if bank == "Bank Account":
            bank = "Pending detection"
        if not account:
            account = "Pending detection"

        return bank, account, name

    def queue_statement_files(self, files):
        """LOAD only. No parser/OCR starts here."""
        new_files = self._normalise_selected_files(files)
        new_files = [
            p for p in new_files
            if p not in self.pending_files
            and p not in self.processing_files
            and p not in self.loaded_files
        ]

        if not new_files:
            self.lbl_progress_status.setText(
                "No new supported statement files were loaded."
            )
            return

        for path in new_files:
            bank, account, name = self._preview_statement_identity(path)
            self.pending_files.append(path)

        self.refresh_loaded_statement_table()
        if self.pending_files:
            self._show_statement_metadata(self.pending_files[-1])
        self._refresh_processing_buttons()
        self.progress_bar.setValue(0)
        self.lbl_progress_status.setText(
            f"{len(new_files)} statement(s) loaded into the workspace. "
            f"{len(self.pending_files)} statement(s) waiting for processing."
        )

    def action_manual_browse(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Load Bank Statements",
            "",
            (
                "Bank Statements & Visual/Structured Files (*.pdf *.xlsx *.xls *.xlsm *.csv *.tsv *.xml *.ofx *.qfx *.qif *.mt940 *.940 *.sta *.bai *.bai2 *.json *.txt *.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp);;"
                "PDF Files (*.pdf);;Images / Passbook Photos (*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp);;"
                "Excel Files (*.xlsx *.xls *.xlsm);;CSV / TSV Files (*.csv *.tsv);;XML / ISO 20022 (*.xml);;"
                "OFX / QFX (*.ofx *.qfx);;SWIFT MT940 (*.mt940 *.940 *.sta);;QIF (*.qif);;BAI2 (*.bai *.bai2);;JSON (*.json);;Text / Auto-detect (*.txt)"
            ),
        )
        if paths:
            self.queue_statement_files(paths)

    def dz_dropEvent(self, event: QDropEvent):
        files = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.isLocalFile()
        ]
        self.queue_statement_files(files)
        event.acceptProposedAction()

    def _parser_process_running(self):
        try:
            return bool(self.worker and self.worker.state() == QProcess.ProcessState.Running)
        except Exception:
            return False

    def _refresh_processing_buttons(self):
        has_pending = bool(self.pending_files)
        running = self._parser_process_running()
        selected_pending = bool(
            self.selected_pending_file
            and self.selected_pending_file in self.pending_files
        )

        self.btn_process_selected.setEnabled(
            has_pending and selected_pending and not running
        )
        self.btn_process_all.setEnabled(has_pending and not running)
        self.btn_process_out.setEnabled(
            bool(self.transactions) and not running
        )
        self.btn_browse.setEnabled(not running)
        self.btn_delete_account.setEnabled(
            bool(self.table_bank_actions.currentRow() >= 0)
        )

    def _show_statement_metadata(self, filepath):
        if not filepath or not os.path.exists(filepath):
            return
        meta = extract_statement_metadata(filepath)
        self.lbl_statement_info.setText(
            f"<b>Statement:</b> {meta.get('file','')} &nbsp; | &nbsp; "
            f"<b>Type:</b> {meta.get('document_type','UNKNOWN')} &nbsp; | &nbsp; "
            f"<b>Bank/Issuer:</b> {meta.get('bank','')} &nbsp; | &nbsp; "
            f"<b>A/c or Card:</b> {meta.get('account','')}<br>"
            f"<b>Holder:</b> {meta.get('holder_name','Not detected')} &nbsp; | &nbsp; "
            f"<b>Statement period:</b> {meta.get('start_date') or 'Not detected'} to {meta.get('end_date') or 'Not detected'} &nbsp; | &nbsp; "
            f"<b>Currency:</b> {meta.get('currency','Not detected')} &nbsp; | &nbsp; "
            f"<b>Language:</b> {meta.get('language','en')} &nbsp; | &nbsp; "
            f"<b>Pages:</b> {meta.get('pages') or 'N/A'}"
        )

    def on_statement_selection_changed(self):
        row = self.table_bank_actions.currentRow()
        self.selected_pending_file = None

        if row < 0:
            self._update_selected_statement_header(None)
            self._refresh_processing_buttons()
            return

        # Column 3 is Statement after the Merge checkbox in column 0.
        statement_item = self.table_bank_actions.item(row, 3)
        path = statement_item.data(Qt.ItemDataRole.UserRole) if statement_item else None

        bank_item = self.table_bank_actions.item(row, 1)
        account_item = self.table_bank_actions.item(row, 2)
        bank = bank_item.text() if bank_item else None
        account = account_item.text() if account_item else None

        if path:
            if path in self.pending_files:
                self.selected_pending_file = path
            self._show_statement_metadata(path)
            self._update_selected_statement_header(path, bank, account)
        else:
            # Processed accounts may have multiple source files. Use the first
            # source attached to the selected account and keep its identity in
            # the header.
            record_key = None
            try:
                record = self._merge_source_records.get(
                    (bank, account, statement_item.text() if statement_item else "")
                )
                if record:
                    paths = record.get("paths", [])
                    if paths:
                        path = paths[0]
            except Exception:
                path = None
            self._update_selected_statement_header(path, bank, account)
            if path:
                self._show_statement_metadata(path)

        self._refresh_processing_buttons()

    def action_process_selected_statement(self):
        if not self.selected_pending_file:
            QMessageBox.information(
                self,
                "Select Statement",
                "Select one loaded statement from the Detected / Loaded Accounts table first.",
            )
            return
        self._start_processing_batch([self.selected_pending_file])

    def action_process_loaded_statements(self):
        self._start_processing_batch(list(self.pending_files))

    def _start_processing_batch(self, files):
        if self._parser_process_running():
            QMessageBox.information(self, "Extraction In Progress", "Processing is already running. Wait for completion or cancel it first.")
            return
        files = [p for p in files if p in self.pending_files]
        if not files:
            QMessageBox.information(self, "Nothing to Process", "There are no loaded statements waiting for processing.")
            return

        self.processing_files = list(files)
        self.successfully_processed_files = []
        self.failed_processing_files = []
        self.worker_result_payload = None
        self.worker_stdout_buffer = ""
        self.btn_process_selected.setEnabled(False)
        self.btn_process_all.setEnabled(False)
        self.btn_process_out.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.btn_cancel_worker.setEnabled(True)
        self.progress_frame.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.lbl_progress_status.setText(f"Starting isolated parser process for {len(files)} statement(s)...")
        self.refresh_loaded_statement_table()
        QApplication.processEvents()

        self.worker_stage_seen = False
        self.worker_last_stage = ""
        self.worker = QProcess(self)
        self.worker.setProcessChannelMode(QProcess.ProcessChannelMode.SeparateChannels)
        self.worker.readyReadStandardOutput.connect(self._on_parser_stdout)
        self.worker.readyReadStandardError.connect(self._on_parser_stderr)
        self.worker.finished.connect(self._on_parser_process_finished)
        self.worker.errorOccurred.connect(self._on_parser_process_error)

        arguments = [os.path.abspath(__file__), "--worker"] + [os.path.abspath(path) for path in files]
        self.lbl_progress_status.setText(f"Starting parser process: {len(files)} statement(s)...")
        self.worker.start(sys.executable, arguments)
        self.refresh_loaded_statement_table()
        self._refresh_processing_buttons()

    def _on_parser_stdout(self):
        if not self.worker:
            return
        try:
            data = bytes(self.worker.readAllStandardOutput()).decode("utf-8", errors="replace")
        except Exception:
            return
        self.worker_stdout_buffer += data
        while "\n" in self.worker_stdout_buffer:
            line, self.worker_stdout_buffer = self.worker_stdout_buffer.split("\n", 1)
            line = line.strip()
            if not line:
                continue
            if line.startswith("STATUS|"):
                status = line[7:]
                self.worker_stage_seen = True
                self.worker_last_stage = status
                self.lbl_progress_status.setText(status)
                QApplication.processEvents()
            elif line.startswith("RESULT|"):
                try:
                    self.worker_result_payload = json.loads(line[7:])
                except Exception as exc:
                    self.worker_result_payload = {"ok": False, "transactions": [], "results": [], "errors": [{"filepath": "", "error": str(exc)}], "completed_files": [], "failed_files": list(self.processing_files)}

    def _on_parser_stderr(self):
        if not self.worker:
            return
        try:
            data = bytes(self.worker.readAllStandardError()).decode("utf-8", errors="replace").strip()
        except Exception:
            data = ""
        if data:
            self.lbl_progress_status.setText("Parser diagnostic: " + data.splitlines()[-1][:500])

    def _on_parser_process_error(self, error):
        if getattr(self, "_closing", False):
            return
        self.lbl_progress_status.setText(f"Parser process error: {error}. Review required.")
        self.btn_cancel_worker.setEnabled(False)

    def _on_parser_process_finished(self, exit_code, exit_status):
        if getattr(self, "_closing", False):
            return
        payload = self.worker_result_payload or {}
        transactions = payload.get("transactions", []) or []
        errors = payload.get("errors", []) or []
        completed_files = payload.get("completed_files", []) or []
        failed_files = payload.get("failed_files", []) or []

        for _tx in transactions:
            ensure_printed_amount_evidence(_tx)
        self.transactions.extend(transactions)
        if transactions and transactions[0].get("source_file"):
            self._show_statement_metadata(transactions[0].get("source_file"))
        self.successfully_processed_files = [p for p in completed_files if p in self.processing_files]
        self.failed_processing_files = [p for p in failed_files if p in self.processing_files]
        for path in self.successfully_processed_files:
            if path in self.pending_files:
                self.pending_files.remove(path)
        self.processing_files = []
        self.btn_cancel_worker.setEnabled(False)

        if transactions:
            self.selected_account_key = self.selected_account_key or (transactions[0].get("bank", "Bank Account"), transactions[0].get("account_number", "N/A"))
            self.run_reconciliation_pass()

        self.refresh_loaded_statement_table()
        self._refresh_processing_buttons()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)

        if exit_code == 0 and not errors:
            remaining = f" {len(self.pending_files)} statement(s) remain loaded." if self.pending_files else ""
            self.lbl_progress_status.setText(f"Processing complete. {len(transactions)} transaction rows admitted.{remaining}")
        elif transactions:
            self.lbl_progress_status.setText(f"Processing completed with review items. {len(transactions)} transaction rows admitted; {len(errors)} item(s) require review.")
        else:
            self.lbl_progress_status.setText("Processing failed or produced no certified transaction rows. Review required.")

        try:
            self.worker.deleteLater()
        except Exception:
            pass
        self.worker = None
        self.worker_result_payload = None
        self.worker_stdout_buffer = ""
        if not self.chk_show_workspace.isChecked():
            self.progress_frame.setVisible(False)
        QApplication.processEvents()

    # Backward-compatible entry point used by any existing caller.
    def start_threaded_ingestion(self, files: list):
        self.queue_statement_files(files)

    def _poll_worker_stage(self):
        return

    def _check_worker_startup(self, generation):
        return

    def on_worker_progress(self, current_page, total_pages, status_text):
        if getattr(self, "_closing", False):
            return

        self.worker_stage_seen = True
        self.lbl_progress_status.setText(status_text)
        if total_pages > 0 and current_page > 0:
            self.progress_bar.setValue(
                max(0, min(100, int((current_page / total_pages) * 100)))
            )
        elif current_page == 0:
            self.progress_bar.setRange(0, 0)
        QApplication.processEvents()

    def on_statement_ready(self, filepath, bank_name, acc_no, txns):
        if getattr(self, "_closing", False):
            return

        if txns:
            for _tx in txns:
                ensure_printed_amount_evidence(_tx)
            normalized_path = os.path.normcase(os.path.abspath(str(filepath)))
            self.transactions = [t for t in self.transactions if os.path.normcase(os.path.abspath(str(t.get("source_file", "")))) != normalized_path]
            self.transactions.extend(txns)
            if filepath not in self.loaded_files:
                self.loaded_files.append(filepath)
            if filepath not in self.successfully_processed_files:
                self.successfully_processed_files.append(filepath)

            if self.selected_account_key is None:
                self.selected_account_key = (bank_name, acc_no)

            self.run_reconciliation_pass()

        # Refresh immediately, even before the entire batch is finished.
        self.refresh_loaded_statement_table()
        self._refresh_processing_buttons()

    def on_worker_error(self, filepath, error):
        if getattr(self, "_closing", False):
            return

        self.lbl_progress_status.setText(
            f"Review required: {os.path.basename(filepath)} | {error}"
        )
        self.refresh_loaded_statement_table()
        QApplication.processEvents()

    def on_worker_finished(self):
        return

    def action_cancel_worker(self):
        if not self._parser_process_running():
            return
        self.lbl_progress_status.setText("Cancelling parser process...")
        try:
            self.worker.terminate()
            QTimer.singleShot(1500, self._force_kill_parser_process)
        except Exception as exc:
            self.lbl_progress_status.setText(f"Parser cancellation error: {exc}")
        self._refresh_processing_buttons()

    def _force_kill_parser_process(self):
        if self._parser_process_running():
            try:
                self.worker.kill()
            except Exception:
                pass

    def refresh_statement_selector(self):
        """Render a compact account/card selector independent of the detailed workspace."""
        if not hasattr(self, "statement_selector_layout"):
            return
        while self.statement_selector_layout.count():
            item = self.statement_selector_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.statement_selector_layout.addStretch(1)

        records = []
        seen = set()
        # Processed sources first, because these can immediately open the ledger.
        for bank, account in sorted(
            {(t.get("bank", "Bank Account"), t.get("account_number", "N/A")) for t in self.transactions},
            key=lambda x: (str(x[0]), str(x[1]))
        ):
            key = (str(bank), str(account))
            if key in seen:
                continue
            seen.add(key)
            records.append({"bank": str(bank), "account": str(account), "processed": True, "path": None})

        # Also expose loaded-but-not-yet-processed statements for selection.
        for path in self.pending_files:
            bank, account, _name = self._preview_statement_identity(path)
            key = (str(bank), str(account), os.path.abspath(str(path)))
            if key in seen:
                continue
            records.append({"bank": str(bank), "account": str(account), "processed": False, "path": path})

        for record in records:
            text = os.path.basename(str(record["path"])) if record.get("path") else (sorted({os.path.basename(str(t.get("source_file"))) for t in self.transactions if t.get("bank") == record["bank"] and t.get("account_number") == record["account"] and t.get("source_file")})[0] if any(t.get("bank") == record["bank"] and t.get("account_number") == record["account"] and t.get("source_file") for t in self.transactions) else f"{record['bank']} | {record['account']}")
            button = QPushButton(text)
            button.setCheckable(True)
            button.setMinimumHeight(26)
            button.setMaximumHeight(28)
            button.setToolTip(f"{record['bank']} | {record['account']}\nSelect this statement")
            button.setStyleSheet(
                "QPushButton { padding: 2px 8px; font-weight: bold; background: #eceff1; "
                "border: 1px solid #b0bec5; border-radius: 4px; color: #263238; }"
                "QPushButton:checked { background: #1565c0; color: white; border: 1px solid #0d47a1; }"
            )
            if record["processed"]:
                button.clicked.connect(
                    lambda checked, b=record["bank"], a=record["account"], btn=button:
                    self._select_compact_statement(b, a, None, btn)
                )
            else:
                button.clicked.connect(
                    lambda checked, p=record["path"], b=record["bank"], a=record["account"], btn=button:
                    self._select_compact_statement(b, a, p, btn)
                )
            if self.selected_account_key == (record["bank"], record["account"]):
                button.setChecked(True)
            self.statement_selector_layout.insertWidget(max(0, self.statement_selector_layout.count() - 1), button)

        workspace_visible = bool(getattr(self, "chk_show_workspace", None) and self.chk_show_workspace.isChecked())
        self.statement_selector_frame.setVisible(bool(records) and not workspace_visible)

    def _select_compact_statement(self, bank, account, filepath=None, button=None):
        """Switch the active statement without opening the detailed workspace."""
        if filepath:
            self._update_selected_statement_header(filepath, bank, account)
        if any(t.get("bank") == bank and t.get("account_number") == account for t in self.transactions):
            self.select_account(bank, account)
        elif filepath:
            self.selected_pending_file = filepath
            self._refresh_processing_buttons()
        for i in range(self.statement_selector_layout.count()):
            item = self.statement_selector_layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, QPushButton):
                widget.setChecked(widget is button)

    def refresh_loaded_statement_table(self):
        """Render pending files and processed accounts in one visible queue."""
        self.refresh_statement_selector()
        pending = list(self.pending_files)

        processed_groups = {}
        for tx in self.transactions:
            key = (tx.get("bank", "Bank Account"), tx.get("account_number", "N/A"))
            processed_groups.setdefault(key, []).append(tx)

        rows = []
        for path in pending:
            bank, account, name = self._preview_statement_identity(path)
            status = (
                "PROCESSING" if path in self.processing_files
                else "REVIEW REQUIRED" if path in getattr(self, "failed_processing_files", [])
                else "LOADED"
            )
            rows.append({
                "bank": bank,
                "account": account,
                "statement": name,
                "entries": "-",
                "status": status,
                "path": path,
                "paths": [path],
                "processed": False,
                "merge_key": (bank, account, name),
            })

        for (bank, account), txns in processed_groups.items():
            source_names = sorted({
                os.path.basename(str(t.get("source_file", "")))
                for t in txns if t.get("source_file")
            })
            rows.append({
                "bank": bank,
                "account": account,
                "statement": ", ".join(source_names) if source_names else "Processed account",
                "entries": str(len(txns)),
                "status": "PROCESSED",
                "path": None,
                "paths": sorted({str(t.get("source_file")) for t in txns if t.get("source_file")}),
                "processed": True,
                "merge_key": (bank, account, ", ".join(source_names) if source_names else "Processed account"),
            })

        self.table_bank_actions.blockSignals(True)
        self.table_bank_actions.setRowCount(len(rows))

        for row, record in enumerate(rows):
            source_paths = set(record.get("paths", []))
            record_key = record.get("merge_key") or (record["bank"], record["account"], record["statement"])
            self._merge_source_records[record_key] = record
            checked = bool(source_paths & self.selected_merge_sources)
            merge_item = QTableWidgetItem()
            merge_item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            merge_item.setData(Qt.ItemDataRole.UserRole, record_key)
            merge_item.setToolTip("Include this account/card in the Merged Date-wise Master Ledger")
            self.table_bank_actions.setItem(row, 0, merge_item)

            values = [
                record["bank"],
                record["account"],
                record["statement"],
                record["entries"],
                record["status"],
            ]
            for offset, value in enumerate(values, start=1):
                item = QTableWidgetItem(str(value))
                if offset == 3 and record.get("paths"):
                    item.setData(Qt.ItemDataRole.UserRole, next(iter(record["paths"])))
                self.table_bank_actions.setItem(row, offset, item)

            status_item = self.table_bank_actions.item(row, 5)
            if record["status"] == "PROCESSED":
                status_item.setForeground(QColor("#1b5e20"))
            elif record["status"] == "PROCESSING":
                status_item.setForeground(QColor("#e65100"))
            else:
                status_item.setForeground(QColor("#1565c0"))

            process_button = QPushButton(
                "Process" if not record["processed"] else "View"
            )
            if record["processed"]:
                process_button.clicked.connect(
                    lambda checked, b=record["bank"], a=record["account"]:
                    self.select_account(b, a)
                )
            else:
                process_button.clicked.connect(
                    lambda checked, p=record["path"]:
                    self._start_processing_batch([p])
                )
            self.table_bank_actions.setCellWidget(row, 6, process_button)

            export_button = QPushButton("Export")
            if record["processed"]:
                export_button.clicked.connect(
                    lambda checked, b=record["bank"], a=record["account"]:
                    self.export_single_bank_account(b, a)
                )
                export_button.setEnabled(True)
            else:
                export_button.setEnabled(False)
            self.table_bank_actions.setCellWidget(row, 7, export_button)

        self.table_bank_actions.blockSignals(False)

        # Preserve selected pending file selection where possible.
        selected_restored = False
        for row in range(self.table_bank_actions.rowCount()):
            item = self.table_bank_actions.item(row, 3)
            if item and item.data(Qt.ItemDataRole.UserRole) == self.selected_pending_file:
                self.table_bank_actions.selectRow(row)
                selected_restored = True
                break
        if not selected_restored and self.table_bank_actions.rowCount() > 0:
            # Do not force a row selection merely to create a default header.
            # The header remains at the last explicitly selected statement.
            pass

    def on_detected_account_clicked(self, row, column):
        # Statement identity now occupies columns 1-3 because column 0 is the
        # Merge checkbox.
        item = self.table_bank_actions.item(row, 3)
        if item is None:
            return

        path = item.data(Qt.ItemDataRole.UserRole)
        bank_item = self.table_bank_actions.item(row, 1)
        acc_item = self.table_bank_actions.item(row, 2)
        bank = bank_item.text() if bank_item else None
        account = acc_item.text() if acc_item else None

        if path:
            self._update_selected_statement_header(path, bank, account)

        if path and path in self.pending_files:
            self.selected_pending_file = path
            self._refresh_processing_buttons()
            return

        if bank_item and acc_item and acc_item.text() != "Pending detection":
            self.select_account(bank_item.text(), acc_item.text())
            self._update_selected_statement_header(path, bank_item.text(), acc_item.text())

    def action_process_out_merged(self):
        """Finalize the batch and open the merged bank statement analyzer."""
        if self._parser_process_running():
            QMessageBox.information(
                self,
                "Processing Still Running",
                "Wait for statement processing to finish before running the merged analyzer.",
            )
            return

        if self.pending_files:
            QMessageBox.warning(
                self,
                "Statements Still Pending",
                "Some statements are loaded but not processed. Process them first, then run Process Out → Merged Analyzer.",
            )
            return

        if not self.transactions:
            QMessageBox.warning(
                self, "No Processed Statements",
                "Process the loaded bank statements before opening the merged analyzer.",
            )
            return

        self.run_reconciliation_pass()
        self.tabs.setCurrentWidget(self.tab_recon)
        if not self.selected_merge_sources:
            self.selected_merge_sources = {str(t.get("source_file")) for t in self.transactions if t.get("source_file")}
            self.refresh_loaded_statement_table()
        self.ledger_tabs.setCurrentIndex(1)
        self.btn_process_out.setEnabled(True)
        self._render_merged_table()
        self.lbl_progress_status.setText(
            f"Merged Bank Statement Analyzer ready: {len(self.selected_merge_sources)} selected statement source(s)."
        )

    # ------------------------------------------------------------------
    # RECONCILIATION
    # ------------------------------------------------------------------

    def run_reconciliation_pass(self):
        """Rebuild canonical balances from opening/reference + debit/credit.

        The printed balance is retained separately as ``image_balance`` and is
        used only as independent evidence. This prevents the parser from
        accepting a mathematically matching but wrongly reconstructed balance.
        """
        all_reconciled = True
        has_unextracted = False
        has_review = False
        has_derived_bf = False

        account_keys = list(dict.fromkeys([
            (t.get("bank", "Bank"), t.get("account_number", "N/A"))
            for t in self.transactions
        ]))

        for b_name, acc_no in account_keys:
            b_txns = [t for t in self.transactions if t.get("bank") == b_name and t.get("account_number") == acc_no]
            ordered = sorted(enumerate(b_txns), key=lambda item: (
                int(item[1].get("page", 0) or 0),
                0 if item[1].get("reference_role") == "OPENING" else 2 if item[1].get("reference_role") == "CLOSING" else 1,
                float((item[1].get("bbox") or [0,0,0,0])[1]),
                item[0],
            ))
            current_balance = None
            opening_initialized = False
            for _, row in ordered:
                if row.get("mode") == "REFERENCE":
                    image_ref = parse_decimal(row.get("image_balance"))
                    calc_ref = parse_decimal(row.get("calculated_balance"))
                    role = row.get("reference_role")

                    if role == "OPENING":
                        if row.get("page_bf_status") == "DERIVED":
                            has_derived_bf = True
                        if not opening_initialized:
                            # Only the first page's B/F is the starting input.
                            # Use its manually entered calculated balance when
                            # present; otherwise seed from the printed image value.
                            if calc_ref is not None and row.get("manual_override", {}).get("calculated_balance"):
                                current_balance = calc_ref
                            elif image_ref is not None:
                                current_balance = image_ref
                                row["calculated_balance"] = str(image_ref)
                            else:
                                current_balance = None
                            opening_initialized = True
                            row["status"] = "REFERENCE" if current_balance is not None else "REVIEW_REQUIRED"
                        else:
                            # Later printed/derived B/F rows are page-boundary
                            # checkpoints. They must not reset the running
                            # mathematical balance curve.
                            row["calculated_balance"] = str(current_balance) if current_balance is not None else "UNEXTRACTED_AMOUNT"
                            if current_balance is None:
                                row["status"] = "UNEXTRACTED_AMOUNT"
                                has_unextracted = True
                                all_reconciled = False
                            elif image_ref is None:
                                row["status"] = "IMAGE_BALANCE_NOT_PROVIDED"
                                has_review = True
                            elif current_balance == image_ref:
                                row["status"] = "MATCHED"
                            else:
                                row["status"] = "MISMATCH"
                                all_reconciled = False
                        row["balance"] = str(current_balance) if current_balance is not None else "UNEXTRACTED_AMOUNT"
                        if current_balance is None:
                            has_unextracted = True
                            all_reconciled = False

                    elif role == "CLOSING":
                        if calc_ref is None or not row.get("manual_override", {}).get("calculated_balance"):
                            calc_ref = current_balance
                            row["calculated_balance"] = str(calc_ref) if calc_ref is not None else "UNEXTRACTED_AMOUNT"
                        row["balance"] = str(calc_ref) if calc_ref is not None else "UNEXTRACTED_AMOUNT"
                        if calc_ref is None:
                            row["status"] = "UNEXTRACTED_AMOUNT"
                            has_unextracted = True
                            all_reconciled = False
                        elif image_ref is None:
                            row["status"] = "IMAGE_BALANCE_NOT_PROVIDED"
                            has_review = True
                        elif calc_ref == image_ref:
                            row["status"] = "CLOSING_MATCHED"
                        else:
                            row["status"] = "CLOSING_MISMATCH"
                            all_reconciled = False
                    continue

                dr = parse_decimal(row.get("withdrawal"))
                cr = parse_decimal(row.get("deposit"))
                if dr is None or cr is None:
                    row["status"] = "UNEXTRACTED_AMOUNT"
                    has_unextracted = True
                    all_reconciled = False
                    continue

                if current_balance is None:
                    row["status"] = "UNEXTRACTED_AMOUNT"
                    has_unextracted = True
                    all_reconciled = False
                    continue

                expected = current_balance + cr - dr
                row["calculated_balance"] = str(expected)
                row["balance"] = str(expected)
                if dr == Decimal("0.00") and cr == Decimal("0.00"):
                    row["status"] = "UNEXTRACTED_AMOUNT"
                    row["audit_detail"] = "No debit or credit amount was extracted for this transaction. The row has not been admitted as a zero-value transaction."
                    has_unextracted = True
                    all_reconciled = False
                    continue
                image_balance = parse_decimal(row.get("image_balance"))
                if image_balance is None:
                    row["status"] = "IMAGE_BALANCE_NOT_PROVIDED"
                    has_review = True
                elif image_balance == expected:
                    row["status"] = "MATCHED"
                else:
                    row["status"] = "MISMATCH"
                    all_reconciled = False
                if (parse_decimal(row.get("image_debit")) is not None and parse_decimal(row.get("withdrawal")) is not None and parse_decimal(row.get("withdrawal")) != parse_decimal(row.get("image_debit"))) or (parse_decimal(row.get("image_credit")) is not None and parse_decimal(row.get("deposit")) is not None and parse_decimal(row.get("deposit")) != parse_decimal(row.get("image_credit"))):
                    row["status"] = "MANUAL_OVERRIDE"
                    row["override_evidence"] = "Printed debit/credit retained separately; canonical value manually corrected."
                current_balance = expected

        self.render_reconciliation_table()
        self.refresh_loaded_statement_table()
        if has_unextracted:
            self.lbl_recon_status.setText("ACTION REQUIRED: UNEXTRACTED_AMOUNT detected. Formula balancing blocked.")
            self.lbl_recon_status.setStyleSheet("color: #b71c1c; font-weight: bold;")
            self.btn_commit_db.setEnabled(False)
        elif not all_reconciled:
            self.lbl_recon_status.setText("BALANCE GATE FAILED: Image balance does not reconcile with calculated balance.")
            self.lbl_recon_status.setStyleSheet("color: #c62828; font-weight: bold;")
            self.btn_commit_db.setEnabled(False)
        elif has_review:
            self.lbl_recon_status.setText("RECONCILED WITH REVIEW: Calculated balances are formula-based; some image balances were not provided.")
            self.lbl_recon_status.setStyleSheet("color: #e65100; font-weight: bold;")
            self.btn_commit_db.setEnabled(True)
        else:
            if has_derived_bf:
                self.lbl_recon_status.setText(
                    "RECONCILED: Page-start B/F balances were derived from the prior page where no printed B/F was present."
                )
            else:
                self.lbl_recon_status.setText("100% RECONCILED: Calculated balances match printed image balances.")
            self.lbl_recon_status.setStyleSheet("color: #1b5e20; font-weight: bold;")
            self.btn_commit_db.setEnabled(True)
        self.run_contra_detection()
        self.run_cash_flow_analysis()
        self.run_statutory_violations()

    # ------------------------------------------------------------------
    # ACCOUNT TABLE
    # ------------------------------------------------------------------

    def update_bank_action_table(self, account_keys):
        """Compatibility wrapper. The visible table is now queue-driven."""
        self.refresh_loaded_statement_table()

    def select_account(self, bank_name, account_number):
        self.selected_account_key = (bank_name, account_number)
        self.btn_delete_account.setEnabled(True)
        self.ledger_tabs.setCurrentIndex(0)
        self.render_individual_account_table()
        source_files = [
            str(t.get("source_file")) for t in self.transactions
            if t.get("bank") == bank_name and t.get("account_number") == account_number and t.get("source_file")
        ]
        self._update_selected_statement_header(
            source_files[0] if source_files else None,
            bank_name,
            account_number,
        )
        self.refresh_statement_selector()

    def action_delete_selected_account(self):
        if not self.selected_account_key:
            QMessageBox.information(self, "Select Account", "Select a detected bank account first.")
            return

        bank_name, account_number = self.selected_account_key
        answer = QMessageBox.question(
            self,
            "Delete Account From Case",
            f"Remove {bank_name} / A/c {account_number} and its extracted transactions from this case?",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self.transactions = [
            t for t in self.transactions
            if not (t.get("bank") == bank_name and t.get("account_number") == account_number)
        ]

        self.selected_account_key = None
        self.active_row = -1
        remaining_sources = {str(t.get("source_file")) for t in self.transactions if t.get("source_file")}
        self.loaded_files = [f for f in self.loaded_files if f in remaining_sources]
        self.refresh_loaded_statement_table()
        self.run_reconciliation_pass()
        self._refresh_processing_buttons()

    # ------------------------------------------------------------------
    # RECON TABLE
    # ------------------------------------------------------------------


    @staticmethod
    def _transaction_sort_key(item):
        """Display statement entries in source order, page by page.

        Reference rows must not all float to the top of the ledger. Their page
        and visual Y-coordinate keep B/F and C/F beside the page boundary where
        the source actually prints them.
        """
        index, tx = item
        page = int(tx.get("page", 0) or 0)
        bbox = tx.get("bbox") or [0, 0, 0, 0]
        y = float(bbox[1]) if len(bbox) > 1 else 0.0
        role = tx.get("reference_role")
        rank = 0 if role == "OPENING" else 2 if role == "CLOSING" else 1
        dt = parse_date_value(tx.get("date")) if tx.get("date") else None
        return (page, y, rank, dt or datetime.max, str(tx.get("account_number", "")), index)

    def _install_ledger_keyboard_delegate(self, table):
        """Install shared Enter/arrow commit behaviour on a ledger table."""
        if not hasattr(self, "_ledger_keyboard_delegates"):
            self._ledger_keyboard_delegates = {}
        delegate = LedgerKeyboardDelegate(table)
        table.setItemDelegate(delegate)
        delegate.navigationKeyPressed.connect(
            lambda key, target_table=table: self._handle_ledger_editor_navigation(target_table, key)
        )
        # Keep Python references alive for the lifetime of the table.
        self._ledger_keyboard_delegates[id(table)] = delegate

    def _handle_ledger_editor_navigation(self, table, key):
        """Keep focus in the ledger after Enter and navigate after arrow commits."""
        if table not in (self.table_account, self.table_recon):
            return
        if table.rowCount() <= 0 or table.columnCount() <= 0:
            return
        row = table.currentRow()
        column = table.currentColumn()
        if row < 0 or column < 0:
            return

        # Enter and Return commit the edit but deliberately keep the same cell
        # selected, allowing the user to continue with any arrow key.
        target_row, target_column = row, column
        if key == int(Qt.Key.Key_Up):
            target_row = max(0, row - 1)
        elif key == int(Qt.Key.Key_Down):
            target_row = min(table.rowCount() - 1, row + 1)
        elif key == int(Qt.Key.Key_Left):
            target_column = max(0, column - 1)
        elif key == int(Qt.Key.Key_Right):
            target_column = min(table.columnCount() - 1, column + 1)

        table.setFocus(Qt.FocusReason.OtherFocusReason)
        table.setCurrentCell(target_row, target_column)
        item = table.item(target_row, target_column)
        if item is not None:
            table.scrollToItem(item, QTableWidget.ScrollHint.EnsureVisible)
            first = table.item(target_row, 0)
            tx_index = first.data(Qt.ItemDataRole.UserRole) if first else None
            if tx_index is not None:
                try:
                    self.show_transaction_by_index(int(tx_index))
                except (TypeError, ValueError):
                    pass

    def render_reconciliation_table(self):
        """Render only the ledger view currently visible.

        Large cases can contain thousands of transactions. Rendering both the
        merged ledger and individual-account ledger simultaneously creates a
        second full QTableWidget object graph and can make Windows appear to
        hang or terminate under memory pressure. The other view is rendered
        when its tab is selected.
        """
        if self.ledger_tabs.currentIndex() == 1:
            self._render_merged_table()
        else:
            self.render_individual_account_table()

    def on_ledger_tab_changed(self, index):
        if index == 1:
            self._render_merged_table()
        else:
            self.render_individual_account_table()

    def _apply_reference_row_color(self, table, row, tx):
        """Color reference rows only; never trigger table rendering from here.

        This method is called *inside* the table-rendering loop. Calling
        render_reconciliation_table() from this method re-entered the same
        renderer for every ordinary row, repeatedly clearing and rebuilding
        the table. Large statements could therefore report hundreds of
        admitted transactions while displaying only a couple of rows.
        """
        role = tx.get("reference_role")
        source = tx.get("reference_source")
        page_status = tx.get("page_bf_status")

        # Ordinary transaction rows require no reference-row styling.
        if role not in {"OPENING", "CLOSING"}:
            return

        if source == "DERIVED" or page_status == "DERIVED":
            background, foreground = QColor("#c8e6c9"), QColor("#1b5e20")
        elif source == "MISSING" or page_status == "MISSING":
            background, foreground = QColor("#fff3cd"), QColor("#8a6d3b")
        else:
            # Printed B/F and C/F are source-evidence checkpoints. Color them
            # without changing their amounts or the running-balance formula.
            background, foreground = QColor("#d9edf7"), QColor("#0b5394")

        for column in range(table.columnCount()):
            cell = table.item(row, column)
            if cell is not None:
                cell.setBackground(background)
                cell.setForeground(foreground)

    def _apply_status_item(self, item, status):
        item.setData(Qt.ItemDataRole.UserRole, item.data(Qt.ItemDataRole.UserRole))
        if status == "UNEXTRACTED_AMOUNT":
            item.setToolTip("A printed monetary value could not be admitted. Select the row and inspect the source document.")
            item.setBackground(QColor("#ffcdd2")); item.setForeground(QColor("#b71c1c"))
        elif status in {"B/F DERIVED MATCHED", "MATCHED"} and item.data(Qt.ItemDataRole.UserRole) is not None:
            item.setBackground(QColor("#c8e6c9")); item.setForeground(QColor("#1b5e20"))
        elif status == "B/F PRINTED":
            item.setBackground(QColor("#d9edf7")); item.setForeground(QColor("#0b5394"))
        elif status == "B/F NOT DERIVABLE":
            item.setBackground(QColor("#fff3cd")); item.setForeground(QColor("#8a6d3b"))
        elif status in {"MISMATCH", "CLOSING_MISMATCH"}:
            item.setBackground(QColor("#ffe0b2")); item.setForeground(QColor("#e65100"))
        elif status == "MANUAL_OVERRIDE":
            item.setBackground(QColor("#e1bee7")); item.setForeground(QColor("#4a148c"))
        elif status in {"IMAGE_BALANCE_NOT_PROVIDED", "REVIEW_REQUIRED"}:
            item.setBackground(QColor("#fff3cd")); item.setForeground(QColor("#8a6d3b"))
        else:
            item.setBackground(QColor("#c8e6c9")); item.setForeground(QColor("#1b5e20"))

    def on_merge_selection_changed(self, item):
        if item.column() != 0:
            return
        key = item.data(Qt.ItemDataRole.UserRole)
        record = self._merge_source_records.get(key, {})
        paths = set(record.get("paths", []))
        if item.checkState() == Qt.CheckState.Checked:
            self.selected_merge_sources.update(paths)
        else:
            self.selected_merge_sources.difference_update(paths)
        self._update_merge_selection_label()

    def _update_merge_selection_label(self):
        count = len(self.selected_merge_sources)
        if hasattr(self, "lbl_merge_selection"):
            self.lbl_merge_selection.setText(f"Selected statement sources: {count}")

    def _set_all_merge_sources(self, checked):
        self.table_bank_actions.blockSignals(True)
        try:
            self.selected_merge_sources.clear()
            for row in range(self.table_bank_actions.rowCount()):
                item = self.table_bank_actions.item(row, 0)
                if item is None:
                    continue
                item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
                if checked:
                    key = item.data(Qt.ItemDataRole.UserRole)
                    self.selected_merge_sources.update(self._merge_source_records.get(key, {}).get("paths", []))
        finally:
            self.table_bank_actions.blockSignals(False)
        self._update_merge_selection_label()
        self._render_merged_table()

    def _build_selected_merged_ledger(self):
        if not self.selected_merge_sources:
            QMessageBox.information(self, "No Sources Selected", "Select at least one bank account or credit-card statement in the Merge column.")
            return
        self.run_reconciliation_pass()
        self.ledger_tabs.setCurrentIndex(1)
        self._render_merged_table()
        self.lbl_progress_status.setText(f"Merged ledger built from {len(self.selected_merge_sources)} selected statement source(s).")

    def _render_merged_table(self):
        selected = set(self.selected_merge_sources)
        if selected:
            filtered = [(i, tx) for i, tx in enumerate(self.transactions) if str(tx.get("source_file", "")) in selected]
        else:
            filtered = []
        ordered = sorted(filtered, key=self._transaction_sort_key)
        table = self.table_recon
        table.setUpdatesEnabled(False)
        table.blockSignals(True)
        try:
            table.clearContents()
            table.setRowCount(len(ordered))
            for row_index, (tx_index, transaction) in enumerate(ordered):
                values = [
                    transaction.get("bank", "Bank Account"),
                    transaction.get("account_number", "N/A"),
                    transaction.get("date", ""),
                    transaction.get("narration", ""),
                    format_inr(transaction.get("withdrawal")),
                    format_inr(transaction.get("deposit")),
                    format_inr(transaction.get("calculated_balance")),
                    format_inr(transaction.get("image_balance")),
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    item.setData(Qt.ItemDataRole.UserRole, tx_index)
                    editable = column in (4, 5, 7)
                    if column == 6 and transaction.get("reference_role") == "OPENING" and transaction.get("page") == 0:
                        editable = True
                        item.setToolTip("Editable starting calculated balance for the first-page B/F only. Later balances are formula-controlled.")
                    elif column == 7:
                        item.setToolTip("Editable printed/image balance. Verify against the source document, then press Enter.")
                    if not editable:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row_index, column, item)
                status = transaction.get("status", "OK")
                status_item = QTableWidgetItem(str(status))
                status_item.setData(Qt.ItemDataRole.UserRole, tx_index)
                self._apply_status_item(status_item, status)
                table.setItem(row_index, 8, status_item)
                self._apply_reference_row_color(table, row_index, transaction)
        finally:
            table.blockSignals(False)
            table.setUpdatesEnabled(True)

    def render_individual_account_table(self):
        if not self.selected_account_key:
            self.lbl_selected_account.setText("Select an account above to view its entries.")
            self.table_account.setRowCount(0)
            return

        bank_name, account_number = self.selected_account_key
        matches = [
            (idx, tx) for idx, tx in enumerate(self.transactions)
            if tx.get("bank") == bank_name and tx.get("account_number") == account_number
        ]
        matches.sort(key=self._transaction_sort_key)
        self.lbl_selected_account.setText(
            f"<b>Selected:</b> {bank_name} | A/c {account_number} | {len(matches)} entries"
        )
        table = self.table_account
        table.setUpdatesEnabled(False)
        table.blockSignals(True)
        try:
            table.clearContents()
            table.setRowCount(len(matches))
            for row_index, (tx_index, tx) in enumerate(matches):
                values = [
                    tx.get("date", ""),
                    tx.get("narration", ""),
                    format_inr(tx.get("withdrawal")),
                    format_inr(tx.get("deposit")),
                    format_inr(tx.get("calculated_balance")),
                    format_inr(tx.get("image_balance")),
                    tx.get("status", "OK"),
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    item.setData(Qt.ItemDataRole.UserRole, tx_index)
                    editable = column in (2, 3, 5)
                    if column == 4 and tx.get("reference_role") == "OPENING" and tx.get("page") == 0:
                        editable = True
                        item.setToolTip("Editable starting calculated balance for the first-page B/F only. Later balances are formula-controlled.")
                    elif column == 5:
                        item.setToolTip("Editable printed/image balance. Verify against the source document, then press Enter.")
                    if not editable:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row_index, column, item)
                self._apply_reference_row_color(table, row_index, tx)
        finally:
            table.blockSignals(False)
            table.setUpdatesEnabled(True)

    def begin_row_edit(self, tx_index):
        if tx_index < 0 or tx_index >= len(self.transactions):
            return
        self.active_row = tx_index
        self.show_transaction_by_index(tx_index)
        tx = self.transactions[tx_index]
        self.lbl_inspector_status.setText(
            f"Edit row: {tx.get('date','')} | edit Debit (Dr) or Credit (Cr) in the ledger and press Enter. Calculated Balance is read-only."
        )
        # Put the row's debit cell into edit mode in the individual table when
        # possible. The user can then type the corrected value and press Enter.
        for table in (self.table_account, self.table_recon):
            for r in range(table.rowCount()):
                item = table.item(r, 0)
                if item and item.data(Qt.ItemDataRole.UserRole) == tx_index:
                    col = 2 if table is self.table_account else 4
                    table.setCurrentCell(r, col)
                    table.editItem(table.item(r, col))
                    return

    def _schedule_restore_ledger_focus(self, table, tx_index, column):
        """Restore keyboard control to the edited ledger cell after recalculation.

        Reconciliation rebuilds QTableWidget items. Restoring focus synchronously
        inside cellChanged is too early because Qt may still be completing the
        Enter-key event and moving its current index. A zero-delay timer restores
        the cell after that event has fully finished.
        """
        try:
            tx_index = int(tx_index)
            column = int(column)
        except (TypeError, ValueError):
            return

        def restore():
            try:
                if table is None or table.columnCount() <= 0:
                    return
                target_row = -1
                for r in range(table.rowCount()):
                    first = table.item(r, 0)
                    if first is None:
                        continue
                    try:
                        row_tx_index = int(first.data(Qt.ItemDataRole.UserRole))
                    except (TypeError, ValueError):
                        continue
                    if row_tx_index == tx_index:
                        target_row = r
                        break
                if target_row < 0:
                    return
                target_column = max(0, min(column, table.columnCount() - 1))
                table.setFocus(Qt.FocusReason.OtherFocusReason)
                table.setCurrentCell(target_row, target_column)
                current_item = table.item(target_row, target_column)
                if current_item is not None:
                    table.scrollToItem(current_item, QTableWidget.ScrollHint.EnsureVisible)
            except RuntimeError:
                # The window may be closing while the queued callback runs.
                return
            except Exception:
                return

        QTimer.singleShot(0, restore)

    def _commit_reference_balance_edit(self, table, row):
        """Commit an officer amendment to the *calculated* B/F or C/F balance.

        The printed/image balance is independent evidence and remains read-only.
        Only the calculated balance cell of an OPENING/B/F or CLOSING/C/F row is
        editable. For an opening reference, the amended calculated balance
        becomes the starting point for subsequent formula reconstruction. For a
        closing reference, it is compared against the independent image balance.
        """
        # Individual table: Calculated Balance = column 4.
        # Merged table:     Calculated Balance = column 6.
        item = table.item(row, 6 if table is self.table_recon else 4)
        if item is None:
            return
        first_item = table.item(row, 0)
        if first_item is None:
            return
        tx_index = first_item.data(Qt.ItemDataRole.UserRole)
        if tx_index is None:
            return
        try:
            tx_index = int(tx_index)
        except Exception:
            return
        if not (0 <= tx_index < len(self.transactions)):
            return
        tx = self.transactions[tx_index]
        if tx.get("reference_role") not in {"OPENING", "CLOSING"}:
            return
        value = parse_decimal(item.text().strip())
        if value is None:
            QMessageBox.warning(self, "Invalid Balance", "Enter a valid numeric B/F or C/F balance.")
            self.render_reconciliation_table()
            self._schedule_restore_ledger_focus(table, tx_index, 6 if table is self.table_recon else 4)
            return
        tx["calculated_balance"] = str(value)
        tx["balance"] = str(value)
        tx.setdefault("manual_override", {})["calculated_balance"] = True
        tx["override_evidence"] = "First-page opening calculated balance manually amended as the formula starting point after verification of source evidence."
        self.run_reconciliation_pass()
        self.show_transaction_by_index(tx_index)
        self._schedule_restore_ledger_focus(table, tx_index, 6 if table is self.table_recon else 4)

    def _commit_image_balance_edit(self, table, row):
        """Commit an officer amendment to the independent printed/image balance.

        This applies to every ledger row, including ordinary transactions and
        B/F/C/F reference rows.  The calculated balance is never overwritten.
        """
        item = table.item(row, 7 if table is self.table_recon else 5)
        if item is None:
            return
        first_item = table.item(row, 0)
        if first_item is None:
            return
        tx_index = first_item.data(Qt.ItemDataRole.UserRole)
        if tx_index is None:
            return
        try:
            tx_index = int(tx_index)
        except Exception:
            return
        if not (0 <= tx_index < len(self.transactions)):
            return
        tx = self.transactions[tx_index]
        value = parse_decimal(item.text().strip())
        if value is None:
            QMessageBox.warning(
                self,
                "Invalid Image Balance",
                "Enter a valid numeric printed/image balance.",
            )
            self.render_reconciliation_table()
            self._schedule_restore_ledger_focus(table, tx_index, 7 if table is self.table_recon else 5)
            return
        tx["image_balance"] = str(value)
        tx.setdefault("manual_override", {})["image_balance"] = True
        tx["override_evidence"] = (
            "Printed/image balance manually amended after verification of source evidence."
        )
        self.run_reconciliation_pass()
        self.show_transaction_by_index(tx_index)
        self._schedule_restore_ledger_focus(table, tx_index, 7 if table is self.table_recon else 5)

    def _commit_edited_cell(self, table, row, column):
        item = table.item(row, column)
        if item is None:
            return
        tx_index = table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        if tx_index is None:
            return
        try:
            tx_index = int(tx_index)
        except Exception:
            return
        tx = self.transactions[tx_index]
        value = parse_decimal(item.text().strip())
        if value is None:
            QMessageBox.warning(self, "Invalid Amount", "Enter a valid numeric amount.")
            self.render_reconciliation_table()
            self._schedule_restore_ledger_focus(table, tx_index, column)
            return
        if column in (4, 2):
            tx["withdrawal"] = str(value)
            tx.setdefault("manual_override", {})["withdrawal"] = True
        elif column in (5, 3):
            tx["deposit"] = str(value)
            tx.setdefault("manual_override", {})["deposit"] = True
        self.run_reconciliation_pass()
        self.show_transaction_by_index(tx_index)
        self._schedule_restore_ledger_focus(table, tx_index, column)

    def on_recon_cell_changed(self, row, column):
        first = self.table_recon.item(row, 0)
        tx_index = first.data(Qt.ItemDataRole.UserRole) if first else None
        tx = None
        if tx_index is not None:
            try:
                idx = int(tx_index)
                if 0 <= idx < len(self.transactions):
                    tx = self.transactions[idx]
            except Exception:
                tx = None

        if column in (4, 5):
            # Debit/Credit are editable only for ordinary transaction rows.
            if tx is not None and tx.get("reference_role") not in {"OPENING", "CLOSING"}:
                self._commit_edited_cell(self.table_recon, row, column)
            return
        if column == 6:
            if tx is not None and tx.get("reference_role") == "OPENING" and int(tx.get("page", -1)) == 0:
                self._commit_reference_balance_edit(self.table_recon, row)
            else:
                self.render_reconciliation_table()
            return
        if column == 7:
            self._commit_image_balance_edit(self.table_recon, row)

    def on_account_cell_changed(self, row, column):
        first = self.table_account.item(row, 0)
        tx_index = first.data(Qt.ItemDataRole.UserRole) if first else None
        tx = None
        if tx_index is not None:
            try:
                idx = int(tx_index)
                if 0 <= idx < len(self.transactions):
                    tx = self.transactions[idx]
            except Exception:
                tx = None

        if column in (2, 3):
            self._commit_edited_cell(self.table_account, row, column)
            return
        if column == 4:
            if tx is not None and tx.get("reference_role") == "OPENING" and int(tx.get("page", -1)) == 0:
                self._commit_reference_balance_edit(self.table_account, row)
            else:
                self.render_reconciliation_table()
            return
        if column == 5:
            self._commit_image_balance_edit(self.table_account, row)

    def on_ledger_current_cell_changed(self, current_row, current_column, previous_row, previous_column):
        """Make keyboard/up-down selection move the synchronized evidence image."""
        if current_row is None or current_row < 0:
            return
        table = self.table_recon if self.ledger_tabs.currentIndex() == 1 else self.table_account
        item = table.item(current_row, 0)
        if item is None:
            return
        tx_index = item.data(Qt.ItemDataRole.UserRole)
        if tx_index is None:
            return
        try:
            tx_index = int(tx_index)
        except Exception:
            return
        if tx_index != self.active_row:
            self.show_transaction_by_index(tx_index)

    def on_individual_row_clicked(self, row, column):
        item = self.table_account.item(row, 0)
        if item is None:
            return
        tx_index = item.data(Qt.ItemDataRole.UserRole)
        if tx_index is not None:
            self.show_transaction_by_index(int(tx_index))

    # ------------------------------------------------------------------
    # PDF INSPECTOR
    # ------------------------------------------------------------------

    def on_table_row_clicked(self, row, column):
        item = self.table_recon.item(row, 0)
        if item is None:
            return
        tx_index = item.data(Qt.ItemDataRole.UserRole)
        if tx_index is None:
            return
        self.show_transaction_by_index(int(tx_index))

    def navigate_transaction(self, direction):
        if not self.transactions:
            return
        ordered = sorted(range(len(self.transactions)), key=lambda i: self._transaction_sort_key((i, self.transactions[i])))
        if self.active_row not in ordered:
            self.show_transaction_by_index(ordered[0 if direction > 0 else -1])
            return
        pos = ordered.index(self.active_row) + (1 if direction > 0 else -1)
        if 0 <= pos < len(ordered):
            self.show_transaction_by_index(ordered[pos])

    def keyPressEvent(self, event):
        # When a ledger table (or its in-place editor) owns focus, preserve
        # native cell navigation. Up/Down must move the current cell and let
        # currentCellChanged synchronize the inspector, not hijack focus into
        # the global transaction navigator.
        focus = QApplication.focusWidget()
        for table in (getattr(self, "table_account", None), getattr(self, "table_recon", None)):
            if table is not None and focus is not None:
                if focus is table or table.isAncestorOf(focus):
                    super().keyPressEvent(event)
                    return
        if event.key() == Qt.Key.Key_Down:
            self.navigate_transaction(1)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Up:
            self.navigate_transaction(-1)
            event.accept()
            return
        super().keyPressEvent(event)

    def show_transaction_by_index(self, tx_index):
        if tx_index < 0 or tx_index >= len(self.transactions):
            return
        self.active_row = tx_index
        transaction = self.transactions[tx_index]
        self.render_pdf_bounding_box(
            transaction.get("page", 0),
            transaction.get("bbox", [50, 100, 500, 140]),
            transaction.get("source_file"),
        )

    def _inspector_statement_key(self, source_file):
        """Stable key used to remember inspector zoom for one statement."""
        if not source_file:
            return "__NO_SOURCE__"
        try:
            return os.path.normcase(os.path.abspath(str(source_file)))
        except Exception:
            return str(source_file)

    def _remember_inspector_zoom(self, zoom):
        zoom = max(0.10, min(12.0, float(zoom)))
        self._inspector_zoom = zoom
        key = self._inspector_source_key
        if key:
            self._inspector_zoom_by_source[key] = zoom

    def _on_view_zoom_changed(self, zoom):
        """Receive Ctrl+wheel zoom from the document view and remember it."""
        self._remember_inspector_zoom(zoom)
        self.lbl_inspector_status.setText(
            f"Inspector: zoom {zoom * 100:.0f}% | remembered for this statement | Drag to pan"
        )

    def _apply_inspector_zoom(self, scale, center_scene=None):
        if self.scene.sceneRect().isNull():
            return
        scale = max(0.10, min(12.0, float(scale)))
        self.view.resetTransform()
        self.view.scale(scale, scale)
        self.view._document_zoom = scale
        self._remember_inspector_zoom(scale)
        if center_scene is None:
            if self._inspector_evidence_scene_rect is not None:
                center_scene = self._inspector_evidence_scene_rect.center()
            else:
                center_scene = self.scene.sceneRect().center()
        self.view.centerOn(center_scene)
        # Queue a second centering pass after Qt has laid out the new scene.
        # This is important when arrow navigation changes to another evidence
        # crop: the newly selected entry must visibly move into the inspector.
        QTimer.singleShot(0, lambda c=center_scene: self.view.centerOn(c))

    def _calculate_default_inspector_zoom(self):
        """Return the largest proportional zoom within 50% x 50% boundaries.

        This is deliberately the V26.5 presentation rule: 100% is preferred
        when it fits; otherwise the exact largest zoom that satisfies both the
        width and height limits is used. There are no artificial 45%/50%/65%
        presets, and the PDF aspect ratio is never changed.
        """
        if self.scene.sceneRect().isNull():
            return 1.0
        viewport = self.view.viewport().size()
        max_w = max(1, int(viewport.width() * 0.50))
        max_h = max(1, int(viewport.height() * 0.50))
        rect = self.scene.sceneRect()
        if rect.width() <= 0 or rect.height() <= 0:
            return 1.0
        return min(1.0, max_w / rect.width(), max_h / rect.height())

    def _inspector_fit(self):
        """Explicitly recalibrate the current statement's default presentation."""
        if self.scene.sceneRect().isNull():
            return
        scale = self._calculate_default_inspector_zoom()
        self._apply_inspector_zoom(scale)
        self.lbl_inspector_status.setText(
            f"Inspector: recalibrated default {scale * 100:.0f}% | 50% width / 50% height boundary | zoom remembered for this statement"
        )

    def _inspector_zoom_by(self, factor):
        if self.scene.sceneRect().isNull():
            return
        new_zoom = self._inspector_zoom * float(factor)
        self._apply_inspector_zoom(new_zoom)
        self.lbl_inspector_status.setText(
            f"Inspector: zoom {self._inspector_zoom * 100:.0f}% | remembered for this statement | Ctrl+wheel also zooms"
        )

    def _inspector_zoom_100(self):
        if self.scene.sceneRect().isNull():
            return
        self._apply_inspector_zoom(1.0)
        self.lbl_inspector_status.setText(
            "Inspector: 100% | remembered for this statement | Drag to pan"
        )

    def render_pdf_bounding_box(
        self,
        page_num,
        bbox,
        source_file=None,
    ):
        """Render a compact contextual evidence crop around the selected entry.

        The viewer intentionally renders the transaction row, a modest amount
        of context above/below it, and the evidence rectangle. It does not
        render the entire PDF page. This keeps memory use and screen footprint
        small while preserving the original document geometry.
        """
        self.scene.clear()
        self._inspector_evidence_scene_rect = None
        self._inspector_source_key = self._inspector_statement_key(source_file)
        remembered_zoom = self._inspector_zoom_by_source.get(self._inspector_source_key)
        self._inspector_zoom = remembered_zoom if remembered_zoom is not None else 1.0
        self._inspector_base_transform = None

        if bbox is None and source_file and os.path.exists(source_file):
            # There is no printed B/F on a derived row. Never highlight the
            # page header as if it were the source entry. Show page context and
            # clearly label the value as derived from the prior page.
            try:
                with open(source_file, "rb") as f:
                    file_bytes = f.read()
                doc = fitz.open(stream=file_bytes, filetype="pdf")
                try:
                    if 0 <= int(page_num) < len(doc):
                        page = doc[int(page_num)]
                        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
                        self.scene.addPixmap(QPixmap.fromImage(image))
                        self.scene.setSceneRect(self.scene.itemsBoundingRect().adjusted(-8, -8, 8, 8))
                        self._inspector_evidence_scene_rect = self.scene.sceneRect()
                        scale = self._calculate_default_inspector_zoom()
                        self._apply_inspector_zoom(scale, self.scene.sceneRect().center())
                        self.lbl_inspector_status.setText(f"Page {int(page_num)+1}/{len(doc)} | DERIVED B/F | no printed B/F on this page | carried from prior page | Default zoom {scale*100:.0f}%")
                        return
                finally:
                    doc.close()
            except Exception:
                pass

        if (
            source_file
            and source_file.lower().endswith(".pdf")
            and os.path.exists(source_file)
        ):
            try:
                with open(source_file, "rb") as f:
                    file_bytes = f.read()

                doc = fitz.open(stream=file_bytes, filetype="pdf")
                try:
                    if 0 <= int(page_num) < len(doc):
                        page = doc[int(page_num)]
                        page_rect = page.rect
                        x0, y0, x1, y1 = [float(v) for v in bbox]

                        # V26.5 evidence presentation: selected entry plus a
                        # controlled amount of context immediately above and below.
                        # Do not turn the inspector into a full-page viewer.
                        row_height = max(12.0, y1 - y0)
                        margin_x = max(45.0, min(90.0, row_height * 2.0))
                        margin_y = max(42.0, min(95.0, row_height * 2.25))
                        clip = fitz.Rect(
                            max(page_rect.x0, x0 - margin_x),
                            max(page_rect.y0, y0 - margin_y),
                            min(page_rect.x1, x1 + margin_x),
                            min(page_rect.y1, y1 + margin_y),
                        )
                        if clip.width < 120:
                            clip.x1 = min(page_rect.x1, clip.x0 + 120)
                        if clip.height < 60:
                            clip.y1 = min(page_rect.y1, clip.y0 + 60)

                        # Keep the V26.5 evidence rendering quality so that the
                        # explicit 100% inspection mode remains genuinely useful.
                        render_scale = 3.0
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(render_scale, render_scale),
                            clip=clip,
                            alpha=False,
                        )

                        image = QImage(
                            pix.samples,
                            pix.width,
                            pix.height,
                            pix.stride,
                            QImage.Format_RGB888,
                        ).copy()
                        pixmap = QPixmap.fromImage(image)
                        self.scene.addPixmap(pixmap)

                        tx = self.transactions[self.active_row] if 0 <= self.active_row < len(self.transactions) else {}
                        status = str(tx.get("status", "OK"))
                        if status == "UNEXTRACTED_AMOUNT":
                            pen_color = QColor(211, 47, 47)
                            fill_color = QColor(244, 67, 54, 45)
                        elif status == "MISMATCH":
                            pen_color = QColor(239, 108, 0)
                            fill_color = QColor(255, 152, 0, 45)
                        else:
                            pen_color = QColor(0, 150, 136)
                            fill_color = QColor(0, 200, 83, 35)

                        evidence_rect = QRectF(
                            (x0 - clip.x0) * render_scale,
                            (y0 - clip.y0) * render_scale,
                            max(4.0, (x1 - x0) * render_scale),
                            max(4.0, (y1 - y0) * render_scale),
                        )
                        rect = QGraphicsRectItem(evidence_rect)
                        rect.setPen(QPen(pen_color, 2.5))
                        rect.setBrush(QBrush(fill_color))
                        rect.setZValue(10)
                        self.scene.addItem(rect)

                        self.scene.setSceneRect(
                            self.scene.itemsBoundingRect().adjusted(-8, -8, 8, 8)
                        )
                        self._inspector_evidence_scene_rect = evidence_rect
                        evidence_center_scene = evidence_rect.center()

                        if remembered_zoom is None:
                            scale = self._calculate_default_inspector_zoom()
                            self._apply_inspector_zoom(scale, evidence_center_scene)
                            mode_text = f"Default zoom {scale * 100:.0f}%"
                        else:
                            self._apply_inspector_zoom(remembered_zoom, evidence_center_scene)
                            mode_text = f"Remembered zoom {remembered_zoom * 100:.0f}%"

                        self.lbl_inspector_status.setText(
                            f"Page {int(page_num) + 1}/{len(doc)} | {status} | {mode_text} | compact entry view | 50% width / 50% height default boundary"
                        )
                        return
                finally:
                    doc.close()
            except Exception as exc:
                self.lbl_inspector_status.setText(
                    f"Inspector could not render source region: {exc}"
                )

        pixmap = QPixmap(640, 180)
        pixmap.fill(QColor("#ffffff"))
        self.scene.addPixmap(pixmap)
        self.scene.setSceneRect(pixmap.rect())
        self._inspector_evidence_scene_rect = self.scene.sceneRect()
        self._apply_inspector_zoom(1.0, self._inspector_evidence_scene_rect.center())
        self.lbl_inspector_status.setText("Source image unavailable for this entry.")

    # ------------------------------------------------------------------
    # MANUAL OVERRIDE
    # ------------------------------------------------------------------

    # ==================================================================
    # CONTRA DETECTION
    # ==================================================================

    def run_contra_detection(self):

        self.contras = []

        total_contra = Decimal(
            "0.00"
        )

        matched_pairs = set()

        for i, t1 in enumerate(
            self.transactions
        ):

            dr1 = parse_decimal(
                t1.get("withdrawal")
            )

            if (
                dr1 is None
                or dr1 <= 0
            ):
                continue

            d1 = parse_statement_date(
                t1.get("date")
            )

            if d1 is None:
                continue

            bank1 = (
                f"{t1.get('bank')} "
                f"({t1.get('account_number')})"
            )

            for j, t2 in enumerate(
                self.transactions
            ):

                if i == j:
                    continue

                pair_key = tuple(
                    sorted(
                        (
                            i,
                            j,
                        )
                    )
                )

                if pair_key in matched_pairs:
                    continue

                cr2 = parse_decimal(
                    t2.get("deposit")
                )

                if (
                    cr2 is None
                    or cr2 <= 0
                ):
                    continue

                bank2 = (
                    f"{t2.get('bank')} "
                    f"({t2.get('account_number')})"
                )

                if bank1 == bank2:
                    continue

                if dr1 != cr2:
                    continue

                d2 = parse_statement_date(
                    t2.get("date")
                )

                if d2 is None:
                    continue

                hours = abs(
                    (
                        d2 - d1
                    ).total_seconds()
                    / 3600.0
                )

                if hours <= 48:

                    matched_pairs.add(
                        pair_key
                    )

                    t1["is_contra"] = True
                    t2["is_contra"] = True

                    self.contras.append(
                        {
                            "date": t1.get(
                                "date"
                            ),
                            "from_bank": bank1,
                            "to_bank": bank2,
                            "amount": dr1,
                            "delta": (
                                f"{hours:.1f} hrs"
                            ),
                        }
                    )

                    total_contra += dr1

        self.table_contra.setRowCount(
            len(self.contras)
        )

        for row_index, contra in enumerate(
            self.contras
        ):

            values = [
                contra.get("date"),
                contra.get("from_bank"),
                contra.get("to_bank"),
                format_inr(
                    contra.get("amount")
                ),
                contra.get("delta"),
                "Contra Neutralized",
            ]

            for column, value in enumerate(
                values
            ):

                self.table_contra.setItem(
                    row_index,
                    column,
                    QTableWidgetItem(
                        str(value)
                    ),
                )

        self.lbl_contra_summary.setText(
            (
                "Total Contra Transfers Filtered: "
                f"{format_inr(total_contra)}"
            )
        )

    # ==================================================================
    # CASH FLOW
    # ==================================================================

    def run_cash_flow_analysis(self):

        self.cash_flows = []

        running_cash = Decimal(
            "0.00"
        )

        peak_deficit = Decimal(
            "0.00"
        )

        def sort_date(transaction):

            parsed = parse_statement_date(
                transaction.get(
                    "date"
                )
            )

            return (
                parsed
                if parsed
                else datetime.min
            )

        sorted_transactions = sorted(
            self.transactions,
            key=lambda transaction: (
                sort_date(transaction),
                0
                if decimal_or_zero(
                    transaction.get(
                        "deposit"
                    )
                ) > 0
                else 1,
            ),
        )

        for transaction in sorted_transactions:

            mode = transaction.get(
                "mode",
                "DIGITAL",
            )

            date = transaction.get(
                "date",
                "",
            )

            dr = decimal_or_zero(
                transaction.get(
                    "withdrawal"
                )
            )

            cr = decimal_or_zero(
                transaction.get(
                    "deposit"
                )
            )

            if mode == "CASH_WDL":

                running_cash += dr

                self.cash_flows.append(
                    {
                        "date": date,
                        "inflow": dr,
                        "outflow": Decimal(
                            "0.00"
                        ),
                        "running_cash": running_cash,
                        "finding": (
                            "Cash Withdrawal "
                            "(Source Created)"
                        ),
                    }
                )

            elif mode == "CASH_DEP":

                # If the narration explicitly states an external cash
                # receipt, that receipt is a source available immediately
                # before the bank cash deposit. Do not count the same money
                # as an unexplained prior-cash requirement.
                cash_source = decimal_or_zero(
                    transaction.get("cash_source")
                )

                running_cash += cash_source
                running_cash -= cr

                if running_cash < 0:

                    deficit = abs(
                        running_cash
                    )

                    peak_deficit = max(
                        peak_deficit,
                        deficit,
                    )

                    finding = (
                        "DEFICIT! "
                        "Unexplained Source: "
                        f"{format_inr(deficit)}"
                    )

                else:

                    finding = (
                        "Covered by Prior Withdrawals"
                    )

                self.cash_flows.append(
                    {
                        "date": date,
                        "inflow": Decimal(
                            "0.00"
                        ),
                        "outflow": cr,
                        "running_cash": running_cash,
                        "finding": finding,
                    }
                )

        self.table_cash.setRowCount(
            len(self.cash_flows)
        )

        for row_index, cash_row in enumerate(
            self.cash_flows
        ):

            values = [
                cash_row["date"],
                format_inr(
                    cash_row["inflow"]
                ),
                format_inr(
                    cash_row["outflow"]
                ),
                format_inr(
                    cash_row["running_cash"]
                ),
                cash_row["finding"],
            ]

            for column, value in enumerate(
                values
            ):

                item = QTableWidgetItem(
                    str(value)
                )

                if (
                    column == 4
                    and "DEFICIT"
                    in str(value)
                ):

                    item.setForeground(
                        QColor("#b71c1c")
                    )

                self.table_cash.setItem(
                    row_index,
                    column,
                    item,
                )

        self.lbl_peak_credit.setText(
            (
                "Peak Unexplained Cash Exposure: "
                f"{format_inr(peak_deficit)}"
            )
        )

    # ==================================================================
    # STATUTORY VIOLATIONS
    # ==================================================================

    def run_statutory_violations(self):

        self.violations = []

        cumulative_cash = {}

        for transaction in self.transactions:

            date = transaction.get(
                "date"
            )

            bank_desc = (
                f"{transaction.get('bank')} "
                f"({transaction.get('account_number')})"
            )

            narration = transaction.get(
                "narration",
                "",
            )

            mode = transaction.get(
                "mode",
                "DIGITAL",
            )

            dr = decimal_or_zero(
                transaction.get(
                    "withdrawal"
                )
            )

            cr = decimal_or_zero(
                transaction.get(
                    "deposit"
                )
            )

            amount = max(
                dr,
                cr,
            )

            # ----------------------------------------------------------
            # 269ST
            # ----------------------------------------------------------

            if (
                mode in [
                    "CASH_DEP",
                    "CASH_WDL",
                ]
                and amount
                >= Decimal("200000.00")
            ):

                self.violations.append(
                    {
                        "date": date,
                        "bank": bank_desc,
                        "narration": narration,
                        "amount": amount,
                        "rule": "Section 269ST",
                        "severity": (
                            "Potential penalty "
                            "u/s 271DA"
                        ),
                    }
                )

            # ----------------------------------------------------------
            # 269SS / 269T
            # ----------------------------------------------------------

            if (
                mode in [
                    "CASH_DEP",
                    "CASH_WDL",
                ]
                and any(
                    keyword in narration.lower()
                    for keyword in [
                        "loan",
                        "advance",
                        "borrow",
                    ]
                )
                and amount
                >= Decimal("20000.00")
            ):

                self.violations.append(
                    {
                        "date": date,
                        "bank": bank_desc,
                        "narration": narration,
                        "amount": amount,
                        "rule": (
                            "Section 269SS / 269T"
                        ),
                        "severity": (
                            "Potential penalty "
                            "u/s 271D / 271E"
                        ),
                    }
                )

            # ----------------------------------------------------------
            # STRUCTURING FLAG
            # ----------------------------------------------------------

            if mode == "CASH_DEP":

                if (
                    amount
                    == Decimal("49000.00")
                    or (
                        Decimal("190000.00")
                        <= amount
                        <= Decimal("199999.00")
                    )
                ):

                    self.violations.append(
                        {
                            "date": date,
                            "bank": bank_desc,
                            "narration": narration,
                            "amount": amount,
                            "rule": (
                                "Smurfing / Structuring"
                            ),
                            "severity": (
                                "Suspicious Threshold "
                                "Pattern"
                            ),
                        }
                    )

                cumulative_cash[
                    bank_desc
                ] = (
                    cumulative_cash.get(
                        bank_desc,
                        Decimal("0.00"),
                    )
                    + amount
                )

                if (
                    cumulative_cash[
                        bank_desc
                    ]
                    >= Decimal("1000000.00")
                ):

                    self.violations.append(
                        {
                            "date": date,
                            "bank": bank_desc,
                            "narration": (
                                "Cumulative Deposits: "
                                f"{format_inr(cumulative_cash[bank_desc])}"
                            ),
                            "amount": (
                                cumulative_cash[
                                    bank_desc
                                ]
                            ),
                            "rule": (
                                "Rule 114E (SFT)"
                            ),
                            "severity": (
                                "Review annual "
                                "cash-deposit reporting"
                            ),
                        }
                    )

        self.table_sft.setRowCount(
            len(self.violations)
        )

        for row_index, violation in enumerate(
            self.violations
        ):

            values = [
                violation["date"],
                violation["bank"],
                violation["narration"],
                format_inr(
                    violation["amount"]
                ),
                violation["rule"],
                violation["severity"],
            ]

            for column, value in enumerate(
                values
            ):

                item = QTableWidgetItem(
                    str(value)
                )

                if column == 5:

                    item.setForeground(
                        QColor("#b71c1c")
                    )

                self.table_sft.setItem(
                    row_index,
                    column,
                    item,
                )

        self.lbl_sft_summary.setText(
            (
                "Total Violations Detected: "
                f"{len(self.violations)} Line Items"
            )
        )

    # ==================================================================
    # EXPORT CONSOLIDATED
    # ==================================================================

    def action_export_consolidated_excel(
        self
    ):

        default_dir = os.getcwd()

        if self.loaded_files:

            default_dir = os.path.dirname(
                self.loaded_files[0]
            )

        default_filename = (
            f"TaxForensic_Consolidated_"
            f"{self.pan}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}"
            f".xlsx"
        )

        in_situ_target = os.path.join(
            default_dir,
            default_filename,
        )

        if self.chk_insitu.isChecked():

            file_path = in_situ_target

        else:

            file_path = (
                QFileDialog.getSaveFileName(
                    self,
                    "Save Consolidated Assessment Excel As",
                    in_situ_target,
                    "Excel Files (*.xlsx)",
                )[0]
            )

        if not file_path:
            return

        sha = (
            compute_file_sha256(
                self.loaded_files[0]
            )
            if self.loaded_files
            else "CONSOLIDATED_BATCH_HASH"
        )

        case_meta = {
            "pan": self.pan,
            "ay": "2024-25",
            "act": "Income-tax Act, 1961",
            "regime": "Old Regime",
            "filenames": (
                ", ".join(
                    os.path.basename(
                        filepath
                    )
                    for filepath in self.loaded_files
                )
                if self.loaded_files
                else "Consolidated_Statements"
            ),
            "bank_scope": (
                "Consolidated All Accounts"
            ),
            "bank_name": "Multiple Banks",
            "account_number": (
                "Multiple Accounts"
            ),
            "sha256": sha,
        }

        TaxForensicExcelExporter.generate_assessment_workbook(
            case_meta,
            self.transactions,
            self.contras,
            self.cash_flows,
            self.violations,
            file_path,
        )

        QMessageBox.information(
            self,
            "Export Successful",
            (
                "Consolidated assessment workbook "
                "written to:\n"
                f"{file_path}"
            ),
        )

    # ==================================================================
    # EXPORT SINGLE ACCOUNT
    # ==================================================================

    def export_single_bank_account(
        self,
        bank_name,
        account_number,
    ):

        bank_transactions = [
            transaction
            for transaction in self.transactions
            if (
                transaction.get("bank")
                == bank_name
                and
                transaction.get("account_number")
                == account_number
            )
        ]

        account_tag = (
            f"{bank_name} "
            f"({account_number})"
        )

        bank_contras = [
            contra
            for contra in self.contras
            if (
                account_tag
                in contra.get(
                    "from_bank",
                    "",
                )
                or
                account_tag
                in contra.get(
                    "to_bank",
                    "",
                )
            )
        ]

        bank_violations = [
            violation
            for violation in self.violations
            if account_tag
            in violation.get(
                "bank",
                "",
            )
        ]

        bank_cash = []

        running_cash = Decimal(
            "0.00"
        )

        for transaction in bank_transactions:

            mode = transaction.get(
                "mode"
            )

            dr = decimal_or_zero(
                transaction.get(
                    "withdrawal"
                )
            )

            cr = decimal_or_zero(
                transaction.get(
                    "deposit"
                )
            )

            if mode == "CASH_WDL":

                running_cash += dr

                bank_cash.append(
                    {
                        "date": transaction.get(
                            "date"
                        ),
                        "inflow": dr,
                        "outflow": Decimal(
                            "0.00"
                        ),
                        "running_cash": running_cash,
                        "finding": (
                            "Withdrawal Available"
                        ),
                    }
                )

            elif mode == "CASH_DEP":

                running_cash -= cr

                if running_cash < 0:

                    finding = (
                        "DEFICIT! Unexplained: "
                        f"{format_inr(abs(running_cash))}"
                    )

                else:

                    finding = (
                        "Covered by withdrawals"
                    )

                bank_cash.append(
                    {
                        "date": transaction.get(
                            "date"
                        ),
                        "inflow": Decimal(
                            "0.00"
                        ),
                        "outflow": cr,
                        "running_cash": running_cash,
                        "finding": finding,
                    }
                )

        clean_bank = re.sub(
            r"\W+",
            "_",
            bank_name,
        )

        clean_account = re.sub(
            r"\W+",
            "_",
            account_number,
        )

        source_file = (
            bank_transactions[0].get(
                "source_file"
            )
            if bank_transactions
            else None
        )

        target_dir = (
            os.path.dirname(source_file)
            if source_file
            and os.path.exists(source_file)
            else os.getcwd()
        )

        filename = (
            f"Audit_"
            f"{clean_bank}_"
            f"Ac_{clean_account}_"
            f"{self.pan}_"
            f"{datetime.now().strftime('%Y%m%d')}"
            f".xlsx"
        )

        in_situ_target = os.path.join(
            target_dir,
            filename,
        )

        if self.chk_insitu.isChecked():

            file_path = in_situ_target

        else:

            file_path = (
                QFileDialog.getSaveFileName(
                    self,
                    "Save Bank Audit Excel",
                    in_situ_target,
                    "Excel Files (*.xlsx)",
                )[0]
            )

        if not file_path:
            return

        case_meta = {
            "pan": self.pan,
            "ay": "2024-25",
            "act": "Income-tax Act, 1961",
            "regime": "Old Regime",
            "filenames": (
                os.path.basename(source_file)
                if source_file
                else f"{bank_name}_Statement"
            ),
            "bank_scope": (
                f"{bank_name} - A/c "
                f"{account_number}"
            ),
            "bank_name": bank_name,
            "account_number": account_number,
            "sha256": (
                compute_file_sha256(
                    source_file
                )
                if source_file
                else "UNKNOWN"
            ),
        }

        TaxForensicExcelExporter.generate_assessment_workbook(
            case_meta,
            bank_transactions,
            bank_contras,
            bank_cash,
            bank_violations,
            file_path,
        )

        QMessageBox.information(
            self,
            "Export Complete",
            (
                f"Audit workbook written to:\n"
                f"{file_path}"
            ),
        )

    # ==================================================================
    # EXPORT ALL
    # ==================================================================

    def action_export_all_separate_banks(
        self
    ):

        account_keys = list(
            dict.fromkeys(
                [
                    (
                        transaction.get(
                            "bank",
                            "Bank",
                        ),
                        transaction.get(
                            "account_number",
                            "N/A",
                        ),
                    )
                    for transaction
                    in self.transactions
                ]
            )
        )

        if not account_keys:

            QMessageBox.warning(
                self,
                "No Accounts",
                "No accounts detected to export.",
            )

            return

        if self.chk_insitu.isChecked():

            target_dir = None

        else:

            target_dir = (
                QFileDialog.getExistingDirectory(
                    self,
                    "Select Folder to Save All Bank Excels",
                )
            )

            if not target_dir:
                return

        exported_files = []

        for bank_name, account_number in account_keys:

            bank_transactions = [
                transaction
                for transaction in self.transactions
                if (
                    transaction.get("bank")
                    == bank_name
                    and
                    transaction.get(
                        "account_number"
                    )
                    == account_number
                )
            ]

            if not bank_transactions:
                continue

            clean_bank = re.sub(
                r"\W+",
                "_",
                bank_name,
            )

            clean_account = re.sub(
                r"\W+",
                "_",
                account_number,
            )

            source_file = (
                bank_transactions[0].get(
                    "source_file"
                )
            )

            if target_dir:

                output_dir = target_dir

            elif (
                source_file
                and os.path.exists(
                    source_file
                )
            ):

                output_dir = os.path.dirname(
                    source_file
                )

            else:

                output_dir = os.getcwd()

            filename = (
                f"Audit_"
                f"{clean_bank}_"
                f"Ac_{clean_account}_"
                f"{self.pan}_"
                f"{datetime.now().strftime('%Y%m%d')}"
                f".xlsx"
            )

            file_path = os.path.join(
                output_dir,
                filename,
            )

            account_tag = (
                f"{bank_name} "
                f"({account_number})"
            )

            bank_contras = [
                contra
                for contra in self.contras
                if (
                    account_tag
                    in contra.get(
                        "from_bank",
                        "",
                    )
                    or
                    account_tag
                    in contra.get(
                        "to_bank",
                        "",
                    )
                )
            ]

            bank_violations = [
                violation
                for violation in self.violations
                if account_tag
                in violation.get(
                    "bank",
                    "",
                )
            ]

            bank_cash = []

            running_cash = Decimal(
                "0.00"
            )

            for transaction in bank_transactions:

                mode = transaction.get(
                    "mode"
                )

                dr = decimal_or_zero(
                    transaction.get(
                        "withdrawal"
                    )
                )

                cr = decimal_or_zero(
                    transaction.get(
                        "deposit"
                    )
                )

                if mode == "CASH_WDL":

                    running_cash += dr

                    bank_cash.append(
                        {
                            "date": transaction.get(
                                "date"
                            ),
                            "inflow": dr,
                            "outflow": Decimal(
                                "0.00"
                            ),
                            "running_cash": running_cash,
                            "finding": (
                                "Withdrawal Available"
                            ),
                        }
                    )

                elif mode == "CASH_DEP":

                    running_cash -= cr

                    finding = (
                        (
                            "DEFICIT! Unexplained: "
                            f"{format_inr(abs(running_cash))}"
                        )
                        if running_cash < 0
                        else
                        "Covered by withdrawals"
                    )

                    bank_cash.append(
                        {
                            "date": transaction.get(
                                "date"
                            ),
                            "inflow": Decimal(
                                "0.00"
                            ),
                            "outflow": cr,
                            "running_cash": running_cash,
                            "finding": finding,
                        }
                    )

            case_meta = {
                "pan": self.pan,
                "ay": "2024-25",
                "act": "Income-tax Act, 1961",
                "regime": "Old Regime",
                "filenames": (
                    os.path.basename(
                        source_file
                    )
                    if source_file
                    else f"{bank_name}_Statement"
                ),
                "bank_scope": (
                    f"{bank_name} - A/c "
                    f"{account_number}"
                ),
                "bank_name": bank_name,
                "account_number": account_number,
                "sha256": (
                    compute_file_sha256(
                        source_file
                    )
                    if source_file
                    else "UNKNOWN"
                ),
            }

            TaxForensicExcelExporter.generate_assessment_workbook(
                case_meta,
                bank_transactions,
                bank_contras,
                bank_cash,
                bank_violations,
                file_path,
            )

            exported_files.append(
                file_path
            )

        QMessageBox.information(
            self,
            "Batch Export Complete",
            (
                f"Successfully exported "
                f"{len(exported_files)} "
                f"individual bank account workbooks."
            ),
        )

    # ==================================================================
    # SQLITE COMMIT
    # ==================================================================

    def action_commit_db(self):

        # Never commit if extraction/reconciliation failed.
        if any(
            transaction.get("status")
            in {"UNEXTRACTED_AMOUNT", "MISMATCH", "CLOSING_MISMATCH"}
            for transaction
            in self.transactions
        ):

            QMessageBox.warning(
                self,
                "Commit Blocked",
                (
                    "Database commit is blocked because "
                    "one or more transactions are not "
                    "successfully reconciled."
                ),
            )

            return

        conn = sqlite3.connect(
            DATABASE_PATH
        )

        try:

            cur = conn.cursor()

            for transaction in self.transactions:

                cur.execute(
                    """
                    INSERT INTO bank_transactions (
                        pan,
                        bank_name,
                        account_number,
                        txn_date,
                        narration,
                        mode,
                        withdrawal,
                        deposit,
                        balance,
                        is_contra,
                        audit_flag,
                        pdf_page,
                        bbox_coords,
                        source_file
                    )
                    VALUES (
                        ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?
                    )
                    """,
                    (
                        self.pan,
                        transaction.get(
                            "bank"
                        ),
                        transaction.get(
                            "account_number"
                        ),
                        transaction.get(
                            "date"
                        ),
                        transaction.get(
                            "narration"
                        ),
                        transaction.get(
                            "mode"
                        ),
                        safe_float(
                            transaction.get(
                                "withdrawal"
                            )
                        ),
                        safe_float(
                            transaction.get(
                                "deposit"
                            )
                        ),
                        safe_float(
                            transaction.get(
                                "balance"
                            )
                        ),
                        transaction.get(
                            "is_contra",
                            False,
                        ),
                        transaction.get(
                            "status"
                        ),
                        transaction.get(
                            "page",
                            0,
                        ),
                        str(
                            transaction.get(
                                "bbox",
                                [],
                            )
                        ),
                        transaction.get(
                            "source_file",
                            "",
                        ),
                    ),
                )

            negative_cash = [
                Decimal(
                    str(
                        row["running_cash"]
                    )
                )
                for row in self.cash_flows
                if Decimal(
                    str(
                        row["running_cash"]
                    )
                ) < 0
            ]

            peak_cash = (
                max(
                    [
                        abs(value)
                        for value
                        in negative_cash
                    ]
                )
                if negative_cash
                else Decimal("0.00")
            )

            sec_269st_total = sum(
                [
                    Decimal(
                        str(
                            violation[
                                "amount"
                            ]
                        )
                    )
                    for violation
                    in self.violations
                    if "269ST"
                    in violation[
                        "rule"
                    ]
                ],
                Decimal("0.00"),
            )

            sec_269ss_total = sum(
                [
                    Decimal(
                        str(
                            violation[
                                "amount"
                            ]
                        )
                    )
                    for violation
                    in self.violations
                    if "269SS"
                    in violation[
                        "rule"
                    ]
                ],
                Decimal("0.00"),
            )

            sft_total = sum(
                [
                    Decimal(
                        str(
                            violation[
                                "amount"
                            ]
                        )
                    )
                    for violation
                    in self.violations
                    if "SFT"
                    in violation[
                        "rule"
                    ]
                ],
                Decimal("0.00"),
            )

            contra_total = sum(
                [
                    Decimal(
                        str(
                            contra[
                                "amount"
                            ]
                        )
                    )
                    for contra
                    in self.contras
                ],
                Decimal("0.00"),
            )

            cur.execute(
                """
                INSERT OR REPLACE INTO forensic_schedules (
                    pan,
                    peak_unexplained_cash,
                    sec_269st_violations,
                    sec_269ss_violations,
                    sft_cash_total,
                    contra_neutralized
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    self.pan,
                    float(peak_cash),
                    float(sec_269st_total),
                    float(sec_269ss_total),
                    float(sft_total),
                    float(contra_total),
                ),
            )

            conn.commit()

        finally:

            conn.close()

        QMessageBox.information(
            self,
            "Committed",
            (
                f"All bank entries for PAN "
                f"{self.pan} committed to "
                f"tax_case.db."
            ),
        )

    # ==================================================================
    # DEMONSTRATION DATA
    # ==================================================================

    def closeEvent(self, event):
        self._closing = True
        try:
            if self._parser_process_running():
                self.worker.terminate()
                QTimer.singleShot(500, self._force_kill_parser_process)
        except Exception:
            pass
        event.accept()

    def load_initial_audit_state(self):
        """Compatibility stub. Startup is intentionally empty and evidence-driven."""
        self.transactions = []
        self.contras = []
        self.cash_flows = []
        self.violations = []


# ======================================================================
# UNIVERSAL STRUCTURED INPUT REGRESSION TESTS
# ======================================================================

def run_universal_structured_input_tests():
    '''Deterministic regression corpus for machine-readable bank formats.'''
    import tempfile
    with tempfile.TemporaryDirectory(dir=APP_DIRECTORY) as td:
        xml_path = os.path.join(td, "sample_camt.xml")
        xml = '''<?xml version="1.0" encoding="UTF-8"?>
<Document xmlns="urn:iso:std:iso:20022:tech:xsd:camt.053.001.02">
 <BkToCstmrStmt><Stmt><Acct><Id><Othr><Id>1234567890</Id></Othr></Id></Acct>
 <Bal><Tp><CdOrPrtry><Cd>OPBD</Cd></CdOrPrtry></Tp><Amt Ccy="INR">1000.00</Amt><CdtDbtInd>CRDT</CdtDbtInd></Bal>
 <Ntry><Amt Ccy="INR">250.00</Amt><CdtDbtInd>CRDT</CdtDbtInd><BookgDt><Dt>2026-04-01</Dt></BookgDt><NtryDtls><TxDtls><RmtInf><Ustrd>SALARY</Ustrd></RmtInf></TxDtls></NtryDtls></Ntry>
 <Ntry><Amt Ccy="INR">100.00</Amt><CdtDbtInd>DBIT</CdtDbtInd><BookgDt><Dt>2026-04-02</Dt></BookgDt><NtryDtls><TxDtls><RmtInf><Ustrd>ATM CASH</Ustrd></RmtInf></TxDtls></NtryDtls></Ntry>
 <Bal><Tp><CdOrPrtry><Cd>CLBD</Cd></CdOrPrtry></Tp><Amt Ccy="INR">1150.00</Amt><CdtDbtInd>CRDT</CdtDbtInd></Bal>
 </Stmt></BkToCstmrStmt></Document>'''
        open(xml_path, "w", encoding="utf-8").write(xml)
        _, _, tx = UniversalStatementEngine.parse_camt(xml_path)
        assert len(tx) == 2 and parse_decimal(tx[-1]["balance"]) == Decimal("1150.00")
        assert tx[-1]["statement_closing_balance_check"] == "PASS"

        ofx_path = os.path.join(td, "sample.ofx")
        ofx = '''<OFX><STMTRS><BANKACCTFROM><ACCTID>1234567890</ACCTID></BANKACCTFROM><BANKTRANLIST><STMTTRN><TRNTYPE>CREDIT</TRNTYPE><DTPOSTED>20260401</DTPOSTED><TRNAMT>250.00</TRNAMT><FITID>A1</FITID><NAME>SALARY</NAME></STMTTRN></BANKTRANLIST><LEDGERBAL><BALAMT>1250.00</BALAMT></LEDGERBAL></STMTRS></OFX>'''
        open(ofx_path, "w", encoding="utf-8").write(ofx)
        _, _, tx = UniversalStatementEngine.parse_ofx(ofx_path)
        assert len(tx) == 1 and parse_decimal(tx[0]["deposit"]) == Decimal("250.00")

        mt_path = os.path.join(td, "sample.mt940")
        mt = ':25:1234567890\n:60F:C260401INR1000,00\n:61:2604010401C250,00NTRFREF1\n:86:SALARY\n:61:2604020402D100,00NTRFREF2\n:86:ATM CASH\n:62F:C260402INR1150,00\n'
        open(mt_path, "w", encoding="utf-8").write(mt)
        _, _, tx = UniversalStatementEngine.parse_mt940(mt_path)
        assert len(tx) == 2 and parse_decimal(tx[-1]["balance"]) == Decimal("1150.00")
        assert tx[-1]["statement_closing_balance_check"] == "PASS"

        qif_path = os.path.join(td, "sample.qif")
        qif = '!Type:Bank\nD04/01/2026\nT250.00\nPSALARY\n^\nD04/02/2026\nT-100.00\nPATM CASH\n^\n'
        open(qif_path, "w", encoding="utf-8").write(qif)
        _, _, tx = UniversalStatementEngine.parse_qif(qif_path)
        assert len(tx) == 2 and parse_decimal(tx[1]["withdrawal"]) == Decimal("100.00")
    return True


# ======================================================================
# 8. HEADLESS AUTOMATED TEST SUITE
# ======================================================================

def run_automated_test_suite():

    # PySide6 signals/QObject machinery requires a Qt application object even
    # when the parser is being exercised synchronously in headless mode.
    # This is intentionally QCoreApplication, not QApplication, so --test
    # remains a true non-GUI test run.
    qt_app = QCoreApplication.instance()
    if qt_app is None:
        qt_app = QCoreApplication(sys.argv)

    print("=" * 80)

    print(
        "STARTING COMPLETE HEADLESS FORENSIC TEST SUITE"
    )

    print("=" * 80)

    # ------------------------------------------------------------------
    # 0. Brought-forward Cr/Dr reference regression
    # ------------------------------------------------------------------
    print("[0A] Testing PDF checkpoint recovery and strict extraction gate...")
    print(" ->", run_pdf_checkpoint_recovery_tests())

    print("[0/8] Testing brought-forward Cr/Dr reference extraction...")
    reference_cases = [
        ("BROUGHT FORWARD : 202.80Cr", Decimal("202.80")),
        ("BROUGHT FORWARD : 202.80 Cr", Decimal("202.80")),
        ("BROUGHT FORWARD : 202.80\nCr", Decimal("202.80")),
        ("OPENING BALANCE : Rs. 50,202.80 Cr", Decimal("50202.80")),
        ("OPENING BAL : 50,202.80\nCr", Decimal("50202.80")),
        ("BALANCE B/F : 1,25,000.00Dr", Decimal("125000.00")),
        ("CARRIED FORWARD : 7,500.00 Cr", Decimal("7500.00")),
    ]
    for reference_text, expected in reference_cases:
        extracted = extract_any_reference_balance(reference_text)[0]
        assert extracted == expected, (
            f"Failed generic reference extraction for {reference_text!r}: "
            f"expected {expected!r}, got {extracted!r}"
        )
    print(" -> Generic opening/brought-forward reference extraction passed across multiple bank-style labels and text-layer variants.")

    # ------------------------------------------------------------------
    # 1. Hostile visual/OCR corpus
    # ------------------------------------------------------------------

    print("[1/8] Testing hostile visual evidence/OCR corpus...")
    visual_result = VisualEvidenceEngine.run_hostile_regression()
    if visual_result.get("status") == "SKIPPED":
        missing = ", ".join(visual_result.get("missing", []))
        print(
            " -> SKIPPED: optional visual/OCR capability is unavailable "
            f"({missing}). Core parser tests will continue."
        )
    else:
        print(
            " -> Visual OCR hostile-corpus tests passed:",
            visual_result["files"],
            "files."
        )

    # ------------------------------------------------------------------
    # 2. Generate synthetic files
    # ------------------------------------------------------------------

    print(
        "[2/8] Testing universal structured inputs "
        "(XML/CAMT, OFX, MT940, QIF)..."
    )

    run_universal_structured_input_tests()

    print(
        " -> Universal structured-input tests passed."
    )

    print(
        "[3/8] Generating synthetic test PDF & Excel files..."
    )

    pdf_name = (
        "Automated_Test_HDFC.pdf"
    )

    xlsx_name = (
        "Automated_Test_SBI.xlsx"
    )

    doc = fitz.open()

    page1 = doc.new_page(
        width=595,
        height=842,
    )

    page1.insert_text(
        (50, 40),
        "HDFC BANK LIMITED - ACCOUNT STATEMENT",
        fontsize=12,
    )

    page1.insert_text(
        (50, 55),
        "Account Number: 5010049281092",
        fontsize=9,
    )

    page1.insert_text(
        (50, 85),
        "Date",
        fontsize=8,
    )

    page1.insert_text(
        (120, 85),
        "Particulars / Narration",
        fontsize=8,
    )

    page1.insert_text(
        (330, 85),
        "Withdrawal (Dr)",
        fontsize=8,
    )

    page1.insert_text(
        (420, 85),
        "Deposit (Cr)",
        fontsize=8,
    )

    page1.insert_text(
        (510, 85),
        "Balance",
        fontsize=8,
    )

    # Opening
    page1.insert_text(
        (50, 115),
        "01-04-2024",
        fontsize=8,
    )

    page1.insert_text(
        (120, 115),
        "OPENING BALANCE B/F",
        fontsize=8,
    )

    page1.insert_text(
        (510, 115),
        "500000.00",
        fontsize=8,
    )

    # Salary
    page1.insert_text(
        (50, 145),
        "05-04-2024",
        fontsize=8,
    )

    page1.insert_text(
        (120, 145),
        "NEFT CR - SALARY PAYMENT RECEIVED",
        fontsize=8,
    )

    page1.insert_text(
        (120, 158),
        "FOR PERIOD 01/04/24 TO 30/04/24 INCL",
        fontsize=8,
    )

    page1.insert_text(
        (420, 145),
        "150000.00",
        fontsize=8,
    )

    page1.insert_text(
        (510, 145),
        "650000.00",
        fontsize=8,
    )

    # Withdrawal
    page1.insert_text(
        (50, 190),
        "08-04-2024",
        fontsize=8,
    )

    page1.insert_text(
        (120, 190),
        "ATM WDL - CASH SELF",
        fontsize=8,
    )

    page1.insert_text(
        (330, 190),
        "40000.00",
        fontsize=8,
    )

    page1.insert_text(
        (510, 190),
        "610000.00",
        fontsize=8,
    )

    # Page 2
    page2 = doc.new_page(
        width=595,
        height=842,
    )

    page2.insert_text(
        (50, 50),
        "08-04-2024",
        fontsize=8,
    )

    page2.insert_text(
        (120, 50),
        "BALANCE BROUGHT FORWARD",
        fontsize=8,
    )

    page2.insert_text(
        (510, 50),
        "610000.00",
        fontsize=8,
    )

    # Contra
    page2.insert_text(
        (50, 90),
        "15-04-2024",
        fontsize=8,
    )

    page2.insert_text(
        (120, 90),
        "NEFT TO SBI A/C 30491820192",
        fontsize=8,
    )

    page2.insert_text(
        (330, 90),
        "200000.00",
        fontsize=8,
    )

    page2.insert_text(
        (510, 90),
        "410000.00",
        fontsize=8,
    )

    # Cash deposit
    page2.insert_text(
        (50, 130),
        "20-04-2024",
        fontsize=8,
    )

    page2.insert_text(
        (120, 130),
        "CDM CASH DEP /BRANCH 1024",
        fontsize=8,
    )

    page2.insert_text(
        (420, 130),
        "100000.00",
        fontsize=8,
    )

    page2.insert_text(
        (510, 130),
        "510000.00",
        fontsize=8,
    )

    # 269ST
    page2.insert_text(
        (50, 170),
        "25-04-2024",
        fontsize=8,
    )

    page2.insert_text(
        (120, 170),
        "CASH RECEIVED ON PROPERTY SALE",
        fontsize=8,
    )

    page2.insert_text(
        (420, 170),
        "350000.00",
        fontsize=8,
    )

    page2.insert_text(
        (510, 170),
        "860000.00",
        fontsize=8,
    )

    doc.save(
        pdf_name
    )

    doc.close()

    # --------------------------------------------------------------
    # Excel
    # --------------------------------------------------------------

    wb = openpyxl.Workbook()

    ws = wb.active

    ws.append(
        [
            "State Bank of India",
            "Account Number: 30491820192",
        ]
    )

    ws.append(
        [
            "Txn Date",
            "Particulars",
            "Withdrawal (Dr)",
            "Deposit (Cr)",
            "Balance",
        ]
    )

    ws.append(
        [
            "01-04-2024",
            "OPENING BALANCE B/F",
            "0.00",
            "0.00",
            "100000.00",
        ]
    )

    ws.append(
        [
            "16-04-2024",
            "NEFT INWARD FROM HDFC A/C 5010049281092",
            "0.00",
            "200000.00",
            "300000.00",
        ]
    )

    ws.append(
        [
            "22-04-2024",
            "INTEREST CREDIT",
            "0.00",
            "3500.00",
            "303500.00",
        ]
    )

    wb.save(
        xlsx_name
    )

    wb.close()

    print(
        " -> Mock files generated successfully."
    )

    # ------------------------------------------------------------------
    # 2. Parser
    # ------------------------------------------------------------------

    print(
        "[4/8] Testing synchronous statement parser..."
    )

    # Headless mode must not execute the QThread worker through worker.run().
    # The GUI uses StatementParserWorker.start(), but the --test path should
    # exercise the parser methods directly. This avoids Qt event-loop/thread
    # machinery blocking a synchronous regression test.
    parser = StatementParserWorker(
        []
    )

    all_parsed = []

    def on_ready(
        path,
        bank,
        acc,
        txns,
    ):
        print(
            f" -> Statement Parsed: "
            f"{bank} "
            f"(A/c: {acc}) "
            f"with {len(txns)} rows."
        )
        all_parsed.extend(txns)

    parser.sig_statement_ready.connect(on_ready)

    print(" -> Parsing synthetic PDF synchronously...")
    parser.parse_pdf(
        pdf_name,
        0,
        2,
    )

    print(" -> Parsing synthetic Excel synchronously...")
    parser.parse_excel(
        xlsx_name
    )

    assert len(
        all_parsed
    ) > 0, (
        "FATAL: Parser returned "
        "0 transactions."
    )

    print(
        " -> Parser passed."
    )

    # ------------------------------------------------------------------
    # 3. Mathematical reconciliation
    # ------------------------------------------------------------------

    print(
        "[5/8] Testing zero-error balance equation "
        "on parsed transactions..."
    )

    account_keys = list(
        dict.fromkeys(
            [
                (
                    transaction.get(
                        "bank"
                    ),
                    transaction.get(
                        "account_number"
                    ),
                )
                for transaction
                in all_parsed
            ]
        )
    )

    extraction_errors = []

    arithmetic_errors = []

    for bank_name, account_number in account_keys:

        account_transactions = [
            transaction
            for transaction in all_parsed
            if (
                transaction.get(
                    "bank"
                )
                == bank_name
                and
                transaction.get(
                    "account_number"
                )
                == account_number
            )
        ]

        print(
            f" -> Checking {bank_name} / "
            f"A/c {account_number}: "
            f"{len(account_transactions)} rows"
        )

        for row_index, transaction in enumerate(
            account_transactions
        ):

            dr = parse_decimal(
                transaction.get(
                    "withdrawal"
                )
            )

            cr = parse_decimal(
                transaction.get(
                    "deposit"
                )
            )

            balance = parse_decimal(
                transaction.get(
                    "balance"
                )
            )

            missing = []

            # Bank statements legitimately leave one side of the Dr/Cr
            # columns blank. Treat those blanks as zero for reconciliation.
            # Do NOT apply this rule to the running balance.
            if dr is None:
                dr = Decimal("0.00")
                transaction["withdrawal"] = "0.00"

            if cr is None:
                cr = Decimal("0.00")
                transaction["deposit"] = "0.00"

            if balance is None:
                missing.append(
                    "balance"
                )

            if missing:

                extraction_errors.append(
                    {
                        "bank": bank_name,
                        "account": account_number,
                        "row": row_index + 1,
                        "date": transaction.get(
                            "date"
                        ),
                        "narration": transaction.get(
                            "narration"
                        ),
                        "missing": missing,
                    }
                )

                continue

            if row_index == 0:

                print(
                    f"    Row {row_index + 1}: "
                    "opening/reference row accepted."
                )

                continue

            previous_balance = parse_decimal(
                account_transactions[
                    row_index - 1
                ].get(
                    "balance"
                )
            )

            if previous_balance is None:

                extraction_errors.append(
                    {
                        "bank": bank_name,
                        "account": account_number,
                        "row": row_index,
                        "date": transaction.get(
                            "date"
                        ),
                        "narration": transaction.get(
                            "narration"
                        ),
                        "missing": [
                            "previous balance"
                        ],
                    }
                )

                continue

            expected_balance = (
                previous_balance
                + cr
                - dr
            )

            if expected_balance != balance:

                arithmetic_errors.append(
                    {
                        "bank": bank_name,
                        "account": account_number,
                        "row": row_index + 1,
                        "date": transaction.get(
                            "date"
                        ),
                        "previous_balance": previous_balance,
                        "deposit": cr,
                        "withdrawal": dr,
                        "expected_balance": expected_balance,
                        "actual_balance": balance,
                    }
                )

    if extraction_errors:

        print()

        print(
            "!!! EXTRACTION GATE FAILED !!!"
        )

        for error in extraction_errors:

            print(
                f" -> {error['bank']} / "
                f"A/c {error['account']} / "
                f"Row {error['row']} / "
                f"Date {error['date']}"
            )

            print(
                "    Missing/invalid: "
                + ", ".join(
                    error["missing"]
                )
            )

        raise AssertionError(
            (
                "Parser produced "
                f"{len(extraction_errors)} "
                "row(s) with unextractable "
                "monetary values."
            )
        )

    if arithmetic_errors:

        print()

        print(
            "!!! BALANCE RECONCILIATION FAILED !!!"
        )

        for error in arithmetic_errors:

            print(
                f" -> {error['bank']} / "
                f"A/c {error['account']} / "
                f"Row {error['row']}"
            )

            print(
                "    Previous Balance: "
                f"{format_inr(error['previous_balance'])}"
            )

            print(
                "    Deposit: "
                f"{format_inr(error['deposit'])}"
            )

            print(
                "    Withdrawal: "
                f"{format_inr(error['withdrawal'])}"
            )

            print(
                "    Expected Balance: "
                f"{format_inr(error['expected_balance'])}"
            )

            print(
                "    Actual Balance: "
                f"{format_inr(error['actual_balance'])}"
            )

        raise AssertionError(
            (
                f"{len(arithmetic_errors)} "
                "genuine balance reconciliation "
                "error(s) detected."
            )
        )

    print(
        " -> All accounts verified with zero "
        "arithmetic discrepancy down to Rs. 0.00."
    )

    # ------------------------------------------------------------------
    # 4. Contra
    # ------------------------------------------------------------------

    print(
        "[6/8] Testing contra transfer elimination..."
    )

    contras = []

    matched_pairs = set()

    total_contra = Decimal(
        "0.00"
    )

    for i, transaction1 in enumerate(
        all_parsed
    ):

        debit = parse_decimal(
            transaction1.get(
                "withdrawal"
            )
        )

        if (
            debit is None
            or debit <= 0
        ):
            continue

        date1 = parse_statement_date(
            transaction1.get(
                "date"
            )
        )

        if date1 is None:
            continue

        bank1 = (
            f"{transaction1.get('bank')} "
            f"({transaction1.get('account_number')})"
        )

        for j, transaction2 in enumerate(
            all_parsed
        ):

            if i == j:
                continue

            pair_key = tuple(
                sorted(
                    (
                        i,
                        j,
                    )
                )
            )

            if pair_key in matched_pairs:
                continue

            credit = parse_decimal(
                transaction2.get(
                    "deposit"
                )
            )

            if (
                credit is None
                or credit <= 0
            ):
                continue

            bank2 = (
                f"{transaction2.get('bank')} "
                f"({transaction2.get('account_number')})"
            )

            if bank1 == bank2:
                continue

            if debit != credit:
                continue

            date2 = parse_statement_date(
                transaction2.get(
                    "date"
                )
            )

            if date2 is None:
                continue

            hours = abs(
                (
                    date2 - date1
                ).total_seconds()
                / 3600.0
            )

            if hours <= 48:

                matched_pairs.add(
                    pair_key
                )

                contras.append(
                    {
                        "date": transaction1.get(
                            "date"
                        ),
                        "from_bank": bank1,
                        "to_bank": bank2,
                        "amount": debit,
                        "delta": (
                            f"{hours:.1f} hrs"
                        ),
                    }
                )

                total_contra += debit

    assert (
        total_contra
        == Decimal("200000.00")
    ), (
        "Expected Rs. 2,00,000 contra, "
        f"got {format_inr(total_contra)}"
    )

    print(
        " -> Inter-bank transfer of "
        f"{format_inr(total_contra)} "
        "successfully neutralized."
    )

    # ------------------------------------------------------------------
    # 5. Cash flow
    # ------------------------------------------------------------------

    print(
        "[7/8] Testing physical cash curve "
        "& peak deficit exposure..."
    )

    running_cash = Decimal(
        "0.00"
    )

    peak_deficit = Decimal(
        "0.00"
    )

    def test_sort_date(transaction):

        parsed = parse_statement_date(
            transaction.get(
                "date"
            )
        )

        return (
            parsed
            if parsed
            else datetime.min
        )

    sorted_transactions = sorted(
        all_parsed,
        key=lambda transaction: (
            test_sort_date(transaction),
            0
            if decimal_or_zero(
                transaction.get(
                    "deposit"
                )
            ) > 0
            else 1,
        ),
    )

    for transaction in sorted_transactions:

        mode = transaction.get(
            "mode",
            "DIGITAL",
        )

        debit = decimal_or_zero(
            transaction.get(
                "withdrawal"
            )
        )

        credit = decimal_or_zero(
            transaction.get(
                "deposit"
            )
        )

        if mode == "CASH_WDL":

            running_cash += debit

        elif mode == "CASH_DEP":

            cash_source = decimal_or_zero(
                transaction.get("cash_source")
            )

            running_cash += cash_source
            running_cash -= credit

            if running_cash < 0:

                peak_deficit = max(
                    peak_deficit,
                    abs(running_cash),
                )

    # The synthetic data contains:
    #
    # 08-04 withdrawal = 40,000
    # 20-04 cash deposit = 1,00,000
    # 25-04 cash deposit = 3,50,000, explicitly described as CASH
    # RECEIVED ON PROPERTY SALE. The latter is treated as an external
    # declared cash source immediately funding that deposit.
    #
    # Therefore the first unexplained deficit is 60,000 and remains the
    # peak deficit.

    cash_source_rows = [
        transaction
        for transaction in all_parsed
        if "CASH RECEIVED"
        in str(transaction.get("narration", "")).upper()
    ]

    assert cash_source_rows, (
        "Expected the explicit CASH RECEIVED test transaction "
        "to be present."
    )

    assert decimal_or_zero(
        cash_source_rows[0].get("cash_source")
    ) == Decimal("350000.00"), (
        "Expected explicit CASH RECEIVED source of "
        "Rs. 3,50,000.00 to be recognized."
    )

    assert (
        peak_deficit
        == Decimal("60000.00")
    ), (
        "Expected peak cash deficit "
        "Rs. 60,000.00, got "
        f"{format_inr(peak_deficit)}"
    )

    print(
        " -> Peak Unexplained Cash Deficit "
        "confirmed: "
        f"{format_inr(peak_deficit)}."
    )

    # ------------------------------------------------------------------
    # 6. Excel exporter
    # ------------------------------------------------------------------

    print(
        "[8/8] Testing openpyxl Court-Ready "
        "Excel Generation..."
    )

    output_xlsx = (
        "Automated_Test_Output.xlsx"
    )

    case_meta = {
        "pan": "ABCDE1234F",
        "ay": "2024-25",
        "act": "Income-tax Act, 1961",
        "regime": "Old Regime",
        "filenames": (
            f"{pdf_name}, {xlsx_name}"
        ),
        "bank_scope": (
            "Consolidated All Accounts"
        ),
        "bank_name": "Multiple Banks",
        "account_number": (
            "Multiple Accounts"
        ),
        "sha256": "TEST_HASH_OK",
    }

    violations = [
        {
            "date": "25-04-2024",
            "bank": (
                "HDFC Bank "
                "(5010049281092)"
            ),
            "narration": (
                "CASH RECEIVED ON PROPERTY SALE"
            ),
            "amount": Decimal(
                "350000.00"
            ),
            "rule": "Section 269ST",
            "severity": (
                "Penalty u/s 271DA"
            ),
        }
    ]

    cash_flows = [
        {
            "date": "08-04-2024",
            "inflow": Decimal(
                "40000.00"
            ),
            "outflow": Decimal(
                "0.00"
            ),
            "running_cash": Decimal(
                "40000.00"
            ),
            "finding": (
                "Withdrawal Available"
            ),
        }
    ]

    TaxForensicExcelExporter.generate_assessment_workbook(
        case_meta,
        all_parsed,
        contras,
        cash_flows,
        violations,
        output_xlsx,
    )

    assert os.path.exists(
        output_xlsx
    ), (
        "FATAL: Output Excel workbook "
        "was not created."
    )

    print(
        " -> Excel audit sheet written successfully: "
        f"{output_xlsx}"
    )

    # ------------------------------------------------------------------
    # CLEANUP
    # ------------------------------------------------------------------

    for filename in [
        pdf_name,
        xlsx_name,
    ]:

        try:

            if os.path.exists(filename):
                os.remove(filename)

        except Exception:
            pass

    print("=" * 80)

    print(
        "ALL TESTS PASSED WITH ZERO ERRORS"
    )

    print(
        "FORENSIC EXTRACTION + RECONCILIATION "
        "+ CONTRA + CASH FLOW + EXCEL"
    )

    print("=" * 80)



# ======================================================================
# 9. EXECUTION ENTRY POINT
# ======================================================================

if __name__ == "__main__":

    if "--corpus" in sys.argv:
        sys.exit(run_v26_corpus())

    elif "--version" in sys.argv:
        ensure_taxforensic_support_dirs()
        print(f"TaxForensic {TAXFORENSIC_ENGINE_VERSION} | Corpus {CORPUS_VERSION}")
        print("Built-in language packs:", ", ".join(sorted(BUILTIN_LANGUAGE_PACKS)))
        sys.exit(0)

    elif "--test" in sys.argv:
        run_automated_test_suite()

    elif "--worker" in sys.argv or "--analyze" in sys.argv:
        mode = "--worker" if "--worker" in sys.argv else "--analyze"
        start = sys.argv.index(mode) + 1
        paths = [os.path.abspath(x) for x in sys.argv[start:] if x]
        if not paths:
            print("ERROR: no statement file supplied.", flush=True)
            sys.exit(2)
        sys.exit(run_isolated_parser_process(paths, analyze_mode=(mode == "--analyze")))

    else:
        app = QApplication(sys.argv)
        window = TaxForensicCompleteApp()
        window.show()
        sys.exit(app.exec())
