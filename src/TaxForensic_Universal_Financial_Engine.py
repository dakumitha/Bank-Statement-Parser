"""TaxForensic universal financial-document normalization engine.

Deterministic, evidence-preserving parsers for structured transaction exports.
The engine never guesses an amount from narration and never merges unlike
asset units (for example, INR and BTC) into one balance.
"""
from __future__ import annotations
import csv, json, os, re, hashlib, unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

ENGINE_VERSION = "UF1.0"

# Canonical fields are mapped from normalized header labels. Extend these
# vocabularies using reviewed fixtures rather than bank-name conditionals.
ALIASES = {
 "date": {"date", "txn date", "transaction date", "trade date", "order date", "execution date", "posting date", "booking date", "value date", "settlement date", "timestamp", "datetime", "time", "fecha", "date de transaction", "交易日期", "交易时间", "تاريخ المعاملة", "तिथि", "दिनांक"},
 "settlement_date": {"settlement date", "settle date", "pay in date", "pay out date", "value date", "fecha de liquidacion"},
 "description": {"description", "narration", "particulars", "remarks", "details", "transaction details", "transaction description", "merchant", "merchant name", "payee", "beneficiary", "name", "symbol description", "memo", "摘要", "交易说明", "الوصف", "विवरण"},
 "reference": {"reference", "reference no", "reference number", "txn id", "transaction id", "transaction reference", "utr", "rrn", "order id", "trade id", "deal id", "contract note no", "contract note number", "exchange order no", "bank reference", "ref", "交易编号", "رقم المرجع"},
 "debit": {"debit", "withdrawal", "withdrawals", "money out", "outflow", "paid out", "dr", "debit amount", "amount debited", "spent", "payment amount", "支出", "借方", "مدين"},
 "credit": {"credit", "deposit", "deposits", "money in", "inflow", "paid in", "cr", "credit amount", "amount credited", "received", "收入", "贷方", "دائن"},
 "amount": {"amount", "transaction amount", "net amount", "total amount", "value", "gross amount", "net total", "trade value", "consideration", "total", "amount inr", "amount (inr)", "amount (usd)"},
 "balance": {"balance", "running balance", "ledger balance", "available balance", "closing balance", "balance after transaction", "solde", "余额", "الرصيد"},
 "currency": {"currency", "ccy", "currency code", "quote currency", "settlement currency", "fiat currency"},
 "symbol": {"symbol", "ticker", "scrip", "scrip name", "security", "security name", "instrument", "instrument name", "trading symbol", "stock", "token", "coin", "asset", "crypto asset", "currency pair", "market"},
 "isin": {"isin", "isin code"},
 "exchange": {"exchange", "exchange name", "segment", "market"},
 "side": {"side", "buy/sell", "buy sell", "transaction type", "trade type", "type", "action", "direction", "order type", "debit credit", "cr/dr"},
 "quantity": {"quantity", "qty", "trade quantity", "traded quantity", "filled quantity", "units", "units traded", "volume", "asset quantity", "coin amount", "crypto amount"},
 "price": {"price", "rate", "trade price", "average price", "avg price", "unit price", "execution price", "price per unit", "market price"},
 "gross": {"gross amount", "gross value", "trade value", "turnover", "consideration", "gross consideration"},
 "fees": {"fees", "fee", "brokerage", "charges", "transaction charges", "total charges", "commission", "trading fee"},
 "taxes": {"tax", "taxes", "gst", "stt", "stamp duty", "sebi charges", "exchange charges", "tds", "withholding tax"},
 "asset_quantity": {"asset quantity", "crypto quantity", "token quantity", "quantity", "qty", "amount of asset", "volume"},
 "asset": {"asset", "asset symbol", "coin", "token", "currency", "crypto", "crypto currency", "digital asset", "vda", "symbol"},
 "tx_hash": {"tx hash", "transaction hash", "hash", "txid", "transaction id", "blockchain transaction"},
 "network": {"network", "chain", "blockchain"},
 "wallet_from": {"from", "from address", "sender", "source wallet", "from wallet"},
 "wallet_to": {"to", "to address", "recipient", "destination wallet", "to wallet"},
 "open_interest": {"open interest", "oi"},
 "realized_pnl": {"realized pnl", "realised pnl", "realized profit", "realised profit", "profit loss", "p&l", "net pnl"},
}

DATE_FORMATS = ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%m/%d/%Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S")


def norm_label(v: Any) -> str:
    s = unicodedata.normalize("NFKC", str(v or "")).replace("\ufeff", "").replace("_", " ")
    s = s.casefold().strip()
    s = re.sub(r"[\u00a0\u2007\u202f\s]+", " ", s)
    s = re.sub(r"[^\w%/ -]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()


def parse_amount(v: Any) -> Decimal | None:
    if v is None or isinstance(v, bool): return None
    if isinstance(v, Decimal): return v
    if isinstance(v, (int, float)):
        try: return Decimal(str(v))
        except InvalidOperation: return None
    s = unicodedata.normalize("NFKC", str(v)).strip()
    if not s or s in {"-", "—", "–", "N/A", "NA", "NULL", "NONE"}: return None
    negative = (s.startswith("(") and s.endswith(")")) or bool(re.search(r"\bDR\.?\s*$", s, re.I))
    s = re.sub(r"\b(?:INR|USD|EUR|GBP|AED|JPY|CNY|HKD|SGD|AUD|CAD|CHF|DR|CR)\b", "", s, flags=re.I)
    s = re.sub(r"[₹$€£¥₽₹₿%\s]", "", s)
    s = s.replace("(", "").replace(")", "").replace("'", "")
    # Handle common decimal conventions, without guessing ambiguous 1,234.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s:
        tail = s.rsplit(",", 1)[1]
        if len(tail) in (1, 2): s = s.replace(",", ".")
        else: s = s.replace(",", "")
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", s): return None
    try:
        x = Decimal(s)
        if negative and x > 0: x = -x
        return x
    except InvalidOperation: return None


def normalize_date(v: Any) -> str | None:
    if v is None: return None
    if isinstance(v, datetime): return v.date().isoformat()
    s = str(v).strip()
    for fmt in DATE_FORMATS:
        try: return datetime.strptime(s[:19], fmt).date().isoformat()
        except ValueError: pass
    # ISO timestamps with timezone suffix: parse date portion only if ISO-like.
    if re.match(r"^\d{4}-\d{2}-\d{2}T", s): return s[:10]
    return None


def map_headers(headers: list[Any]) -> dict[str, int]:
    normalized = [norm_label(h) for h in headers]
    result: dict[str, int] = {}
    # Exact normalized match first, then only safe full-phrase containment.
    for field, aliases in ALIASES.items():
        choices = {norm_label(a) for a in aliases}
        for i, h in enumerate(normalized):
            if h in choices:
                result.setdefault(field, i)
        if field not in result:
            matches = []
            for i, h in enumerate(normalized):
                if h and any(len(a) >= 5 and (h.startswith(a + " ") or h.endswith(" " + a)) for a in choices): matches.append(i)
            if len(matches) == 1: result[field] = matches[0]
    return result


def detect_document_type(headers: list[Any], sample_rows: list[list[Any]] | None = None, filename: str = "") -> str:
    h = " ".join(norm_label(x) for x in headers)
    s = (h + " " + norm_label(str(filename).replace("_", " ").replace("-", " ")) + " " + " ".join(norm_label(c) for row in (sample_rows or [])[:5] for c in row)).lower()
    score = {"BANK": 0, "CREDIT_CARD": 0, "BROKER": 0, "VDA": 0}
    for term in ("running balance", "account number", "withdrawal", "deposit", "opening balance", "closing balance", "utr", "upi"): score["BANK"] += term in s
    if "balance" in h and "date" in h and ("description" in h or "narration" in h or "particulars" in h): score["BANK"] += 2
    for term in ("credit card", "minimum due", "minimum payment", "credit limit", "available credit", "statement balance", "merchant"): score["CREDIT_CARD"] += term in s
    for term in ("contract note", "brokerage", "isin", "exchange", "trade quantity", "trading symbol", "settlement date", "stt", "sebi charges", "buy/sell", "realized pnl", "realised pnl"): score["BROKER"] += term in s
    for term in ("tx hash", "transaction hash", "blockchain", "wallet", "network", "crypto", "token quantity", "digital asset", "vda", "coin"): score["VDA"] += term in s
    winner = max(score, key=score.get)
    return winner if score[winner] > 0 else "UNKNOWN"


def _get(row: list[Any], mapping: dict[str, int], key: str) -> Any:
    i = mapping.get(key)
    return row[i] if i is not None and i < len(row) else None


def _string(v: Any) -> str | None:
    if v is None: return None
    s = str(v).strip()
    return s or None


def _signed_amount(row: list[Any], mapping: dict[str, int], dtype: str, desc: str = "", currency: str | None = None, asset: str | None = None) -> tuple[Decimal | None, Decimal | None, list[str]]:
    flags: list[str] = []
    debit = parse_amount(_get(row, mapping, "debit")) if "debit" in mapping else None
    credit = parse_amount(_get(row, mapping, "credit")) if "credit" in mapping else None
    amount = parse_amount(_get(row, mapping, "amount")) if "amount" in mapping else None
    side = norm_label(_get(row, mapping, "side"))
    if debit is not None or credit is not None:
        debit = debit or Decimal("0")
        credit = credit or Decimal("0")
        if debit < 0 or credit < 0: flags.append("NEGATIVE_VALUE_IN_DEBIT_CREDIT_COLUMN")
        if debit > 0 and credit > 0: flags.append("BOTH_DEBIT_AND_CREDIT_POSITIVE")
        if debit == 0 and credit == 0 and (amount is not None and amount != 0): flags.append("AMOUNT_PRESENT_BUT_DIRECTION_COLUMNS_EMPTY")
        return abs(debit), abs(credit), flags
    if amount is None:
        flags.append("AMOUNT_NOT_EXTRACTED")
        return None, None, flags
    if dtype == "VDA" and currency and asset and currency.strip().upper() == asset.strip().upper():
        flags.append("ASSET_MOVEMENT_NOT_CASH_FLOW")
        return Decimal("0"), Decimal("0"), flags
    if dtype != "BANK":
        flags.append("CASH_DIRECTION_NOT_INFERRED_FROM_UNMARKED_AMOUNT")
        return Decimal("0"), Decimal("0"), flags
    if amount < 0: return abs(amount), Decimal("0"), flags
    if any(t in side for t in ("debit", "withdraw", "sell", "sold", "payment", "fee", "charge", "outflow", "send", "spent")):
        return abs(amount), Decimal("0"), flags
    if any(t in side for t in ("credit", "deposit", "buy", "bought", "refund", "receive", "inflow", "salary")):
        return Decimal("0"), abs(amount), flags
    if dtype == "BANK":
        flags.append("DIRECTION_AMBIGUOUS")
        return Decimal("0"), Decimal("0"), flags
    # For card/broker/VDA, a signed/net amount is not automatically cash debit/credit.
    return Decimal("0"), Decimal("0"), ["CASH_DIRECTION_NOT_INFERRED_FROM_UNMARKED_AMOUNT"]


def normalize_rows(headers: list[Any], rows: list[list[Any]], source_file: str, sheet: str | None = None, source_row_start: int = 2) -> dict[str, Any]:
    mapping = map_headers(headers)
    dtype = detect_document_type(headers, rows, os.path.basename(source_file))
    txns: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=source_row_start):
        if not row or not any(str(x or "").strip() for x in row): continue
        date_raw = _get(row, mapping, "date")
        date_iso = normalize_date(date_raw)
        desc = _string(_get(row, mapping, "description"))
        ref = _string(_get(row, mapping, "reference"))
        qty = parse_amount(_get(row, mapping, "quantity")) if "quantity" in mapping else None
        amount = parse_amount(_get(row, mapping, "amount")) if "amount" in mapping else None
        asset_qty = parse_amount(_get(row, mapping, "asset_quantity")) if "asset_quantity" in mapping else qty
        # Some VDA exports label the token quantity simply Amount. Only use it
        # as asset quantity when the currency column explicitly matches the asset.
        if dtype == "VDA" and asset_qty is None and amount is not None:
            _asset_hint = _string(_get(row, mapping, "asset")) or _string(_get(row, mapping, "symbol"))
            _currency_hint = _string(_get(row, mapping, "currency"))
            _side_hint = norm_label(_get(row, mapping, "side"))
            if _asset_hint and _currency_hint and _asset_hint.upper() == _currency_hint.upper() and any(t in _side_hint for t in ("deposit", "withdraw", "transfer", "send", "receive")):
                asset_qty = abs(amount)
        price = parse_amount(_get(row, mapping, "price")) if "price" in mapping else None
        gross = parse_amount(_get(row, mapping, "gross")) if "gross" in mapping else None
        fees = parse_amount(_get(row, mapping, "fees")) if "fees" in mapping else None
        taxes = parse_amount(_get(row, mapping, "taxes")) if "taxes" in mapping else None
        side = _string(_get(row, mapping, "side"))
        _currency_for_direction = _string(_get(row, mapping, "currency"))
        _asset_for_direction = _string(_get(row, mapping, "asset")) or _string(_get(row, mapping, "symbol"))
        debit, credit, flags = _signed_amount(row, mapping, dtype, desc or "", _currency_for_direction, _asset_for_direction)
        balance = parse_amount(_get(row, mapping, "balance")) if "balance" in mapping else None
        asset = _string(_get(row, mapping, "asset")) or _string(_get(row, mapping, "symbol"))
        symbol = _string(_get(row, mapping, "symbol"))
        if dtype in {"BROKER", "VDA"} and qty is not None and price is not None and gross is not None:
            product = abs(qty * price)
            if abs(product - abs(gross)) > Decimal("0.02"):
                flags.append("QUANTITY_PRICE_GROSS_MISMATCH")
        if not date_raw: flags.append("DATE_MISSING")
        elif not date_iso: flags.append("DATE_NOT_NORMALIZED")
        if not desc and not symbol and not ref: flags.append("IDENTITY_FIELDS_MISSING")
        if dtype == "BANK" and (debit is None or credit is None): flags.append("BANK_AMOUNT_INCOMPLETE")
        if dtype == "BROKER" and qty is None and amount is None and gross is None: flags.append("TRADE_QUANTITY_AND_VALUE_MISSING")
        if dtype == "VDA" and asset_qty is None and amount is None: flags.append("ASSET_QUANTITY_AND_FIAT_VALUE_MISSING")
        if dtype == "CREDIT_CARD" and amount is None and debit in (None, Decimal("0")) and credit in (None, Decimal("0")): flags.append("CARD_AMOUNT_MISSING")
        fingerprint_data = "|".join(str(x or "") for x in (date_iso, desc, ref, symbol, side, qty, price, amount, debit, credit))
        fingerprint = hashlib.sha256(fingerprint_data.encode("utf-8")).hexdigest()
        status = "OK" if not flags else "REVIEW_REQUIRED"
        txns.append({
            "document_type": dtype, "source_file": str(source_file), "source_sheet": sheet,
            "source_row": index, "source_page": None, "date": date_iso, "date_original": _string(date_raw),
            "description": desc, "narration": desc, "reference": ref, "symbol": symbol,
            "isin": _string(_get(row, mapping, "isin")), "exchange": _string(_get(row, mapping, "exchange")),
            "side": side, "quantity": str(qty) if qty is not None else None,
            "unit_price": str(price) if price is not None else None, "gross_amount": str(gross) if gross is not None else None,
            "fees": str(fees) if fees is not None else None, "taxes": str(taxes) if taxes is not None else None,
            "currency": _string(_get(row, mapping, "currency")), "asset": asset,
            "asset_quantity": str(asset_qty) if asset_qty is not None else None,
            "network": _string(_get(row, mapping, "network")), "tx_hash": _string(_get(row, mapping, "tx_hash")),
            "wallet_from": _string(_get(row, mapping, "wallet_from")), "wallet_to": _string(_get(row, mapping, "wallet_to")),
            "realized_pnl": str(parse_amount(_get(row, mapping, "realized_pnl"))) if parse_amount(_get(row, mapping, "realized_pnl")) is not None else None,
            "amount": str(amount) if amount is not None else None,
            "withdrawal": str(debit) if debit is not None else None, "deposit": str(credit) if credit is not None else None,
            "balance": str(balance) if balance is not None else None, "image_balance": str(balance) if balance is not None else None,
            "calculated_balance": None, "source_evidence": json.dumps({"headers": [str(x) for x in headers], "row": [str(x) if x is not None else None for x in row]}, ensure_ascii=False),
            "fingerprint": fingerprint, "review_flags": sorted(set(flags)), "status": status,
        })
    # Duplicates are flagged; no rows are deleted or merged automatically.
    counts: dict[str, int] = {}
    for tx in txns: counts[tx["fingerprint"]] = counts.get(tx["fingerprint"], 0) + 1
    for tx in txns:
        if counts[tx["fingerprint"]] > 1:
            tx["review_flags"].append("POSSIBLE_DUPLICATE")
            tx["review_flags"] = sorted(set(tx["review_flags"]))
            tx["status"] = "REVIEW_REQUIRED"
    return {"engine_version": ENGINE_VERSION, "document_type": dtype, "source_file": str(source_file), "headers": [str(x) for x in headers], "header_mapping": mapping, "transactions": txns, "transaction_count": len(txns), "review_count": sum(t["status"] != "OK" for t in txns)}


def read_structured_file(path: str) -> dict[str, Any]:
    ext = Path(path).suffix.lower()
    if ext in {".csv", ".tsv", ".txt"}:
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
            sample = f.read(8192); f.seek(0)
            try: dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
            except csv.Error: dialect = csv.excel_tab if ext == ".tsv" else csv.excel
            data = list(csv.reader(f, dialect))
        if not data: return normalize_rows([], [], path)
        # Find likely header among first 30 rows; avoid treating title/metadata as column names.
        best_i, best_score = 0, -1
        for i, row in enumerate(data[:30]):
            mapping = map_headers(row)
            score = len(mapping) + (3 if "date" in mapping else 0) + (2 if any(k in mapping for k in ("amount", "debit", "credit", "quantity")) else 0)
            if score > best_score: best_i, best_score = i, score
        if best_score < 2: return normalize_rows(data[0], data[1:], path)
        return normalize_rows(data[best_i], data[best_i+1:], path, source_row_start=best_i + 2)
    if ext in {".xlsx", ".xlsm"}:
        try: import openpyxl
        except ImportError as e: raise RuntimeError("openpyxl is required to read XLSX/XLSM") from e
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            results = []
            for ws in wb.worksheets:
                data = [list(row) for row in ws.iter_rows(values_only=True)]
                best_i, best_score = 0, -1
                for i, row in enumerate(data[:30]):
                    score = len(map_headers(row))
                    if score > best_score: best_i, best_score = i, score
                if best_score >= 2:
                    result = normalize_rows(data[best_i], data[best_i+1:], path, ws.title, best_i + 2)
                    results.append(result)
            if not results: return normalize_rows([], [], path)
            tx = [t for r in results for t in r["transactions"]]
            dtypes = {r["document_type"] for r in results}
            return {"engine_version": ENGINE_VERSION, "document_type": dtypes.pop() if len(dtypes)==1 else "MIXED", "source_file": path, "transactions": tx, "transaction_count": len(tx), "review_count": sum(t["status"]!="OK" for t in tx), "sheets": [r.get("source_sheet") for r in results]}
        finally: wb.close()
    if ext == ".json":
        with open(path, encoding="utf-8-sig") as f: payload = json.load(f)
        if isinstance(payload, dict):
            rows = next((payload[k] for k in ("transactions", "trades", "data", "records", "ledger") if isinstance(payload.get(k), list)), None)
            if rows is None: rows = [payload]
        elif isinstance(payload, list): rows = payload
        else: rows = []
        if not rows: return normalize_rows([], [], path)
        headers = list(dict.fromkeys(str(k) for r in rows if isinstance(r, dict) for k in r.keys()))
        return normalize_rows(headers, [[r.get(h) for h in headers] for r in rows if isinstance(r, dict)], path)
    raise ValueError(f"Unsupported structured format: {ext}")


def analyze_ledger(result: dict[str, Any]) -> dict[str, Any]:
    """Produce separate, unit-safe summaries; never sum unrelated assets."""
    txns = result.get("transactions", [])
    summary: dict[str, Any] = {"document_type": result.get("document_type", "UNKNOWN"), "transaction_count": len(txns), "review_count": sum(t.get("status") != "OK" for t in txns), "possible_duplicates": sum("POSSIBLE_DUPLICATE" in t.get("review_flags", []) for t in txns), "by_currency": {}, "by_asset": {}, "by_symbol": {}, "flags": {}}
    for tx in txns:
        for flag in tx.get("review_flags", []): summary["flags"][flag] = summary["flags"].get(flag, 0) + 1
        cur = tx.get("currency") or "UNSPECIFIED"
        bucket = summary["by_currency"].setdefault(cur, {"debit": "0", "credit": "0", "gross": "0", "fees": "0", "taxes": "0"})
        for k, dest in (("withdrawal", "debit"), ("deposit", "credit"), ("gross_amount", "gross"), ("fees", "fees"), ("taxes", "taxes")):
            value = parse_amount(tx.get(k))
            if value is not None: bucket[dest] = str(Decimal(bucket[dest]) + value)
        asset = tx.get("asset")
        q = parse_amount(tx.get("asset_quantity"))
        if asset and q is not None:
            summary["by_asset"].setdefault(asset, "0")
            summary["by_asset"][asset] = str(Decimal(summary["by_asset"][asset]) + q)
        symbol = tx.get("symbol")
        pnl = parse_amount(tx.get("realized_pnl"))
        if symbol and pnl is not None:
            summary["by_symbol"].setdefault(symbol, "0")
            summary["by_symbol"][symbol] = str(Decimal(summary["by_symbol"][symbol]) + pnl)
    return summary


def to_application_transactions(result: dict[str, Any]) -> list[dict[str, Any]]:
    """Adapt the richer ledger records to the existing TaxForensic GUI contract.

    Extra fields are retained, so downstream analysis can use the richer data.
    Missing balances stay missing; no synthetic balance is created here.
    """
    txns = []
    for item in result.get("transactions", []):
        tx = dict(item)
        tx.setdefault("bank", "Financial Document")
        tx.setdefault("account_number", "N/A")
        tx.setdefault("mode", "DIGITAL")
        tx.setdefault("page", 0)
        tx.setdefault("bbox", [0, 0, 0, 0])
        tx.setdefault("balance", "UNEXTRACTED_AMOUNT")
        tx.setdefault("image_balance", "UNEXTRACTED_AMOUNT")
        tx.setdefault("source_format", "UNIVERSAL_STRUCTURED")
        tx.setdefault("extraction_method", "HEADER_MAPPED_STRUCTURED")
        tx["status"] = "OK" if not tx.get("review_flags") else "REVIEW_REQUIRED"
        txns.append(tx)
    return txns
