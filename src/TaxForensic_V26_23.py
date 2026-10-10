import sys
import re
import os
import hashlib
import sqlite3
import csv
import json
import xml.etree.ElementTree as ET
import html


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

TAXFORENSIC_ENGINE_VERSION = "V26.23"
CORPUS_VERSION = "1"

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