import csv, json, os, sys, tempfile, unittest
from decimal import Decimal
_HERE = os.path.dirname(os.path.abspath(__file__))
_SRC = os.path.join(os.path.dirname(_HERE), "src")
sys.path.insert(0, _SRC if os.path.isdir(_SRC) else _HERE)
from TaxForensic_Universal_Financial_Engine import read_structured_file, analyze_ledger, parse_amount

class UniversalFinancialEngineTests(unittest.TestCase):
    def write_csv(self, td, name, rows):
        p = os.path.join(td, name)
        with open(p, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)
        return p

    def test_bank_csv_header_mapping_preserves_debit_credit_and_balance(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "bank.csv", [
                ["Transaction Date", "Narration", "Debit", "Credit", "Running Balance"],
                ["02-04-2024", "UPI/DR/416894450831/Payment", "123.60", "", "6880.25"],
                ["03-04-2024", "SALARY", "", "50000.00", "56880.25"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "BANK")
            self.assertEqual(r["transaction_count"], 2)
            self.assertEqual(Decimal(r["transactions"][0]["withdrawal"]), Decimal("123.60"))
            self.assertEqual(Decimal(r["transactions"][0]["balance"]), Decimal("6880.25"))
            self.assertEqual(Decimal(r["transactions"][1]["deposit"]), Decimal("50000.00"))

    def test_credit_card_export_is_classified_and_amount_not_assumed_directionless(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "credit_card_statement.csv", [
                ["Transaction Date", "Merchant", "Description", "Amount", "Currency", "Transaction Type"],
                ["2026-01-02", "SHOP", "Grocery purchase", "1200.00", "INR", "Purchase"],
                ["2026-01-04", "SHOP", "Refund", "-200.00", "INR", "Refund"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "CREDIT_CARD")
            self.assertEqual(r["transaction_count"], 2)
            self.assertIn("CASH_DIRECTION_NOT_INFERRED_FROM_UNMARKED_AMOUNT", r["transactions"][0]["review_flags"])
            self.assertEqual(Decimal(r["transactions"][1]["amount"]), Decimal("-200.00"))

    def test_broker_trade_checks_quantity_price_and_gross(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "contract_note.csv", [
                ["Trade Date", "Trading Symbol", "ISIN", "Exchange", "Buy/Sell", "Quantity", "Trade Price", "Gross Amount", "Brokerage", "STT", "Order ID"],
                ["2026-04-01", "ABC", "INE000A01000", "NSE", "BUY", "10", "100.00", "1000.00", "20.00", "1.00", "ORD-1"],
                ["2026-04-02", "XYZ", "INE000B01000", "NSE", "SELL", "5", "200.00", "900.00", "10.00", "1.00", "ORD-2"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "BROKER")
            self.assertEqual(r["transaction_count"], 2)
            self.assertNotIn("QUANTITY_PRICE_GROSS_MISMATCH", r["transactions"][0]["review_flags"])
            self.assertIn("QUANTITY_PRICE_GROSS_MISMATCH", r["transactions"][1]["review_flags"])
            self.assertEqual(r["transactions"][0]["isin"], "INE000A01000")

    def test_vda_export_keeps_token_quantity_separate_from_fiat(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "crypto_transactions.csv", [
                ["Timestamp", "Type", "Asset", "Amount", "Currency", "Network", "Transaction Hash", "From Address", "To Address"],
                ["2026-01-02T10:00:00", "Withdrawal", "BTC", "0.015", "BTC", "Bitcoin", "hash-1", "wallet-a", "wallet-b"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "VDA")
            self.assertEqual(r["transaction_count"], 1)
            self.assertEqual(r["transactions"][0]["asset"], "BTC")
            self.assertEqual(Decimal(r["transactions"][0]["asset_quantity"]), Decimal("0.015"))
            self.assertEqual(r["transactions"][0]["tx_hash"], "hash-1")
            self.assertEqual(r["transactions"][0]["withdrawal"], "0")
            self.assertIn("ASSET_MOVEMENT_NOT_CASH_FLOW", r["transactions"][0]["review_flags"])
            self.assertIsNone(r["transactions"][0]["balance"])

    def test_ambiguous_bank_amount_is_flagged_not_guessed(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "bank.csv", [
                ["Date", "Description", "Amount", "Balance"],
                ["2026-01-02", "TRANSFER", "100.00", "900.00"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "BANK")
            self.assertIn("DIRECTION_AMBIGUOUS", r["transactions"][0]["review_flags"])

    def test_duplicate_rows_are_flagged_not_deleted(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "bank.csv", [
                ["Date", "Narration", "Debit", "Credit", "Balance"],
                ["2026-01-02", "UPI/DR/123456", "10", "", "90"],
                ["2026-01-02", "UPI/DR/123456", "10", "", "90"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["transaction_count"], 2)
            self.assertTrue(all("POSSIBLE_DUPLICATE" in t["review_flags"] for t in r["transactions"]))


    def test_metadata_before_header_preserves_actual_source_row(self):
        with tempfile.TemporaryDirectory() as td:
            p = self.write_csv(td, "bank.csv", [
                ["Bank statement export"],
                ["Account", "masked"],
                ["Transaction Date", "Narration", "Debit", "Credit", "Balance"],
                ["2026-01-02", "UPI/DR/123456", "10", "", "90"],
            ])
            r = read_structured_file(p)
            self.assertEqual(r["transactions"][0]["source_row"], 4)

    def test_json_broker_rows_are_normalized(self):
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "trades.json")
            with open(p, "w", encoding="utf-8") as f:
                json.dump({"trades": [{"trade_date": "2026-04-01", "trading_symbol": "ABC", "isin": "INE000A01000", "buy_sell": "BUY", "quantity": "2", "trade_price": "50", "gross_amount": "100"}]}, f)
            r = read_structured_file(p)
            self.assertEqual(r["document_type"], "BROKER")
            self.assertEqual(r["transaction_count"], 1)
            self.assertEqual(r["transactions"][0]["symbol"], "ABC")

    def test_amount_locale_handling(self):
        self.assertEqual(parse_amount("₹1,23,456.78"), Decimal("123456.78"))
        self.assertEqual(parse_amount("1.234,56 EUR"), Decimal("1234.56"))
        self.assertEqual(parse_amount("(125.50)"), Decimal("-125.50"))

if __name__ == "__main__":
    unittest.main(verbosity=2)
