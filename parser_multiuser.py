import os
import re
import json
import imaplib
import email
from email.header import decode_header
from datetime import datetime
from typing import List, Dict, Optional, Any


# =========================================================
# PDF TEXT EXTRACTION
# =========================================================
def extract_text_from_pdf(file_path: str) -> str:
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("Install pdfplumber using: pip install pdfplumber")

    extracted_parts: List[str] = []

    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_parts.append(text)

                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            if row:
                                row_text = " ".join(str(cell).strip() for cell in row if cell)
                                if row_text:
                                    extracted_parts.append(row_text)
    except Exception as e:
        raise Exception(f"Failed to read PDF '{file_path}': {e}")

    final_text = "\n".join(extracted_parts).strip()

    if not final_text:
        raise ValueError(f"No readable text found in PDF '{file_path}' (possibly scanned PDF).")

    return final_text


# =========================================================
# IMAP EMAIL FETCHING
# =========================================================
def decode_mime_words(value: str) -> str:
    if not value:
        return ""

    decoded_parts = decode_header(value)
    result: List[str] = []

    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(encoding or "utf-8", errors="ignore"))
        else:
            result.append(part)

    return "".join(result).strip()


def sanitize_filename(filename: str) -> str:
    filename = filename.strip().replace("\\", "_").replace("/", "_")
    filename = re.sub(r"[^A-Za-z0-9._ -]", "_", filename)
    return filename or f"attachment_{int(datetime.now().timestamp())}.pdf"


def infer_imap_server(email_address: str) -> str:
    domain = email_address.split("@")[-1].lower()
    common_servers = {
        "gmail.com": "imap.gmail.com",
        "yahoo.com": "imap.mail.yahoo.com",
        "outlook.com": "outlook.office365.com",
        "hotmail.com": "outlook.office365.com",
        "live.com": "outlook.office365.com",
    }
    return common_servers.get(domain, f"imap.{domain}")


def fetch_pdf_from_email_imap(
    email_address: str,
    app_password: str,
    save_dir: str = "downloads/email_attachments",
    subject_keywords: Optional[List[str]] = None,
    sender_filters: Optional[List[str]] = None,
    max_emails_to_scan: int = 50,
    already_downloaded: Optional[set] = None,
) -> Optional[Dict[str, str]]:
    if subject_keywords is None:
        subject_keywords = [
            "statement",
            "credit card",
            "e-statement",
            "card statement",
            "monthly statement",
            "bill",
        ]

    if sender_filters is None:
        sender_filters = []

    if already_downloaded is None:
        already_downloaded = set()

    os.makedirs(save_dir, exist_ok=True)
    imap_server = infer_imap_server(email_address)
    mail = None

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, app_password)
        mail.select("INBOX")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            raise Exception("Failed to search inbox.")

        message_ids = messages[0].split()
        if not message_ids:
            return None

        recent_ids = message_ids[-max_emails_to_scan:]

        for msg_id in reversed(recent_ids):
            msg_key = msg_id.decode(errors="ignore")
            if msg_key in already_downloaded:
                continue

            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            for response_part in msg_data:
                if not isinstance(response_part, tuple):
                    continue

                msg = email.message_from_bytes(response_part[1])
                subject = decode_mime_words(msg.get("Subject", ""))
                sender = decode_mime_words(msg.get("From", ""))

                subject_lower = subject.lower()
                sender_lower = sender.lower()

                subject_ok = any(keyword in subject_lower for keyword in subject_keywords)
                sender_ok = True if not sender_filters else any(s in sender_lower for s in sender_filters)

                for part in msg.walk():
                    content_disposition = part.get("Content-Disposition", "")
                    if "attachment" not in content_disposition.lower():
                        continue

                    filename = part.get_filename()
                    if not filename:
                        continue

                    filename = decode_mime_words(filename)
                    filename_lower = filename.lower()
                    is_pdf = filename_lower.endswith(".pdf")
                    looks_relevant = (
                        subject_ok
                        or "statement" in filename_lower
                        or "card" in filename_lower
                        or "bill" in filename_lower
                        or "estatement" in filename_lower
                    )

                    if not is_pdf or not looks_relevant or not sender_ok:
                        continue

                    safe_name = sanitize_filename(filename)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    save_path = os.path.join(save_dir, f"{timestamp}_{safe_name}")

                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue

                    with open(save_path, "wb") as f:
                        f.write(payload)

                    already_downloaded.add(msg_key)

                    return {
                        "email_subject": subject,
                        "email_sender": sender,
                        "attachment_filename": safe_name,
                        "saved_pdf_path": save_path,
                    }

        return None

    except imaplib.IMAP4.error as e:
        raise Exception(
            f"IMAP login failed: {e}. If using Gmail, use an App Password instead of your normal password."
        )
    except Exception as e:
        raise Exception(f"Email fetch failed: {e}")
    finally:
        try:
            if mail:
                mail.logout()
        except Exception:
            pass


# =========================================================
# STATEMENT PARSER
# =========================================================
class StatementParser:
    def __init__(self, text: str):
        self.raw_text = text
        self.cleaned_text = self._clean_text(text)

    def _clean_text(self, text: str) -> str:
        text = text.replace("₹", " ").replace("Rs.", " ").replace("INR", " ")
        text = re.sub(r"[^\x00-\x7F]+", " ", text)
        text = re.sub(r"[ \t]+", " ", text)
        return text.strip()

    def _parse_date(self, date_str: str, default_year: Optional[int] = None) -> Optional[str]:
        date_str = date_str.strip()

        formats = [
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%d %b %Y",
            "%d %B %Y",
            "%d-%b-%y",
            "%d-%b-%Y",
            "%d/%m/%y",
            "%d-%m-%y",
            "%d %b",
            "%d %B",
        ]

        for fmt in formats:
            try:
                dt = datetime.strptime(date_str, fmt)

                if "%Y" not in fmt and "%y" not in fmt:
                    if default_year is not None:
                        dt = dt.replace(year=default_year)
                    else:
                        return None

                return dt.strftime("%Y-%m-%d")
            except ValueError:
                continue

        return None

    def _clean_amount(self, amount_str: str) -> Optional[float]:
        try:
            amount_str = amount_str.replace(",", "").replace(" ", "")
            return float(amount_str)
        except Exception:
            return None

    def _search_amount_after_label(self, labels: List[str]) -> Optional[float]:
        for label in labels:
            pattern = rf"{label}.*?([+-]?[0-9,]+(?:\.\d{{2}})?)"
            match = re.search(pattern, self.cleaned_text, re.IGNORECASE | re.DOTALL)
            if match:
                value = self._clean_amount(match.group(1))
                if value is not None:
                    return value
        return None

    def _search_date_after_label(self, labels: List[str]) -> Optional[str]:
        patterns = [
            r"(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
            r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        ]
        for label in labels:
            for date_pattern in patterns:
                pattern = rf"{label}.*?{date_pattern}"
                match = re.search(pattern, self.cleaned_text, re.IGNORECASE | re.DOTALL)
                if match:
                    parsed = self._parse_date(match.group(1))
                    if parsed:
                        return parsed
        return None

    def get_statement_year(self) -> Optional[int]:
        statement_date = self.extract_statement_date()
        if statement_date:
            try:
                return datetime.strptime(statement_date, "%Y-%m-%d").year
            except ValueError:
                pass

        period = self.extract_statement_period()
        if period.get("start"):
            try:
                return datetime.strptime(period["start"], "%Y-%m-%d").year
            except ValueError:
                pass

        return None

    # -------------------------------
    # HEADER / SUMMARY EXTRACTION
    # -------------------------------
    def extract_bank_name(self) -> Optional[str]:
        lines = [line.strip() for line in self.raw_text.splitlines() if line.strip()]
        for line in lines[:5]:
            if "bank" in line.lower():
                return line.strip()
        return None

    def extract_customer_name(self) -> Optional[str]:
        patterns = [
            r"Customer\s+Name[:\s]*([A-Za-z .'-]+)",
            r"Name[:\s]*([A-Za-z .'-]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, self.cleaned_text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return None

    def extract_card_number(self) -> Optional[str]:
        patterns = [
            r"Card\s+Number[:\s]*([Xx*0-9 ]{12,25})",
            r"Card\s+No[:\s]*([Xx*0-9 ]{12,25})",
        ]
        for pattern in patterns:
            match = re.search(pattern, self.cleaned_text, re.IGNORECASE)
            if match:
                return re.sub(r"\s+", " ", match.group(1)).strip()
        return None

    def extract_statement_date(self) -> Optional[str]:
        return self._search_date_after_label([r"Statement\s+Date"])

    def extract_statement_period(self) -> Dict[str, Optional[str]]:
        patterns = [
            r"Statement\s+Period[:\s]*(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\s*[-–]\s*(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
            r"Statement\s+Period[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*[-–]\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        ]
        for pattern in patterns:
            match = re.search(pattern, self.cleaned_text, re.IGNORECASE)
            if match:
                return {
                    "start": self._parse_date(match.group(1)),
                    "end": self._parse_date(match.group(2)),
                }
        return {"start": None, "end": None}

    def extract_total_due(self) -> Optional[float]:
        return self._search_amount_after_label([
            r"TOTAL\s+AMOUNT\s+DUE",
            r"AMOUNT\s+DUE",
            r"TOTAL\s+DUE",
        ])

    def extract_due_date(self) -> Optional[str]:
        return self._search_date_after_label([
            r"PAYMENT\s+DUE\s+DATE",
            r"DUE\s+DATE",
        ])

    def extract_previous_balance(self) -> Optional[float]:
        return self._search_amount_after_label([r"Previous\s+Balance"])

    def extract_payments_credits(self) -> Optional[float]:
        return self._search_amount_after_label([r"Payments\s*/\s*Credits", r"Payments\s+Credits"])

    def extract_credit_limit(self) -> Optional[float]:
        return self._search_amount_after_label([r"Credit\s+Limit"])

    def extract_available_credit(self) -> Optional[float]:
        return self._search_amount_after_label([r"Available\s+Credit"])

    def extract_retail_purchases(self) -> Optional[float]:
        return self._search_amount_after_label([r"Retail\s+Purchases"])

    def extract_finance_charges(self) -> Optional[float]:
        return self._search_amount_after_label([r"Finance\s+Charges"])

    def extract_fees_taxes(self) -> Optional[float]:
        return self._search_amount_after_label([r"Fees\s*&\s*Taxes", r"Fees\s+Taxes"])

    def extract_minimum_due(self) -> Optional[float]:
        return self._search_amount_after_label([r"Minimum\s+Amount\s+Due", r"Minimum\s+Due"])

    # -------------------------------
    # CATEGORIZATION
    # -------------------------------
    def categorize_merchant(self, merchant: str) -> str:
        merchant_lower = merchant.lower()
        category_rules = {
            "Food": ["swiggy", "zomato", "dominos", "easydiner", "starbucks"],
            "Travel": ["uber", "ola", "makemytrip", "irctc", "fastag"],
            "Shopping": ["amazon", "flipkart", "myntra", "ajio", "nykaa", "decathlon"],
            "Groceries": ["dmart", "reliance smart", "bigbasket", "blinkit", "zepto", "zeptonow", "lulu"],
            "Bills & Recharge": ["jio", "airtel", "cred"],
            "Health": ["apollo", "medplus", "1mg"],
            "Entertainment": ["netflix", "spotify", "bookmyshow", "pvr"],
            "Fuel": ["shell", "indian oil", "petrol pump", "petrol", "fuel"],
        }
        for category, keywords in category_rules.items():
            if any(keyword in merchant_lower for keyword in keywords):
                return category
        return "Others"

    def is_noise_line(self, line: str) -> bool:
        noise_keywords = [
            "previous balance",
            "credit limit",
            "available credit",
            "minimum amount due",
            "minimum due",
            "total amount due",
            "payment due date",
            "statement date",
            "statement period",
            "account summary",
            "customer name",
            "card number",
            "fees",
            "finance charges",
            "retail purchases",
            "payments / credits",
            "three-month spend overview",
            "one-month spend overview",
            "transaction count",
            "spend amount",
            "sample statement",
            "nova bank",
            "zenith bank",
            "credit card statement",
            "txn date",
            "merchant description",
            "date merchant amount",
            "amount (rs.)",
            "transactions -",
            "continued",
        ]
        line_lower = line.lower()
        return any(keyword in line_lower for keyword in noise_keywords)

    # -------------------------------
    # TRANSACTION EXTRACTION
    # -------------------------------
    def extract_transactions(self) -> List[Dict[str, Any]]:
        transactions: List[Dict[str, Any]] = []
        seen = set()
        lines = self.raw_text.split("\n")
        default_year = self.get_statement_year()

        date_pattern = (
            r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
            r"|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}"
            r"|\d{1,2}-[A-Za-z]{3}-\d{2,4}"
            r"|\d{1,2}\s+[A-Za-z]{3,9})"
        )
        amount_pattern = r"([+-]?\d[\d,]*\.?\d{0,2})\s*$"

        for line in lines:
            line = line.strip()
            if not line or self.is_noise_line(line):
                continue

            date_match = re.search(date_pattern, line)
            amount_match = re.search(amount_pattern, line)

            if not date_match or not amount_match:
                continue

            raw_date = date_match.group(1)
            parsed_date = self._parse_date(raw_date, default_year=default_year)
            amount = self._clean_amount(amount_match.group(1))

            if not parsed_date or amount is None:
                continue

            merchant_start = date_match.end()
            merchant_end = amount_match.start()
            merchant = line[merchant_start:merchant_end].strip()
            merchant = re.sub(r"[^A-Za-z0-9&.*\-/ ]", "", merchant)
            merchant = re.sub(r"\s+", " ", merchant).strip()

            if not merchant:
                continue

            txn_key = (parsed_date, merchant.lower(), amount)
            if txn_key in seen:
                continue
            seen.add(txn_key)

            transactions.append(
                {
                    "date": parsed_date,
                    "merchant": merchant.title(),
                    "amount": amount,
                    "category": self.categorize_merchant(merchant),
                }
            )

        transactions.sort(key=lambda x: x["date"])
        return transactions

    # -------------------------------
    # ANALYTICS
    # -------------------------------
    def build_summary(self, transactions: List[Dict[str, Any]]) -> Dict[str, Any]:
        if not transactions:
            return {
                "total_transactions": 0,
                "total_spent": 0.0,
                "average_transaction": 0.0,
                "highest_transaction": None,
                "lowest_transaction": None,
            }

        amounts = [txn["amount"] for txn in transactions]
        return {
            "total_transactions": len(transactions),
            "total_spent": round(sum(amounts), 2),
            "average_transaction": round(sum(amounts) / len(amounts), 2),
            "highest_transaction": round(max(amounts), 2),
            "lowest_transaction": round(min(amounts), 2),
        }

    def build_monthly_spend(self, transactions: List[Dict[str, Any]]) -> Dict[str, float]:
        monthly_spend: Dict[str, float] = {}
        for txn in transactions:
            month_key = txn["date"][:7]
            monthly_spend[month_key] = round(monthly_spend.get(month_key, 0.0) + txn["amount"], 2)
        return dict(sorted(monthly_spend.items()))

    def build_category_spend(self, transactions: List[Dict[str, Any]]) -> Dict[str, float]:
        category_spend: Dict[str, float] = {}
        for txn in transactions:
            category = txn["category"]
            category_spend[category] = round(category_spend.get(category, 0.0) + txn["amount"], 2)
        return dict(sorted(category_spend.items(), key=lambda x: x[0]))

    def flag_suspicious_transactions(self, transactions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        suspicious: List[Dict[str, Any]] = []
        if not transactions:
            return suspicious

        avg = sum(txn["amount"] for txn in transactions) / len(transactions)
        threshold = max(10000.0, avg * 3)

        for txn in transactions:
            reasons = []
            if txn["amount"] >= threshold:
                reasons.append("High amount")
            if txn["category"] == "Others" and txn["amount"] > 5000:
                reasons.append("Unusual merchant/category")
            if reasons:
                suspicious.append(
                    {
                        "date": txn["date"],
                        "merchant": txn["merchant"],
                        "amount": txn["amount"],
                        "category": txn["category"],
                        "reason": ", ".join(reasons),
                    }
                )
        return suspicious

    def build_derived(self, summary_fields: Dict[str, Any], txn_summary: Dict[str, Any]) -> Dict[str, Any]:
        credit_limit = summary_fields.get("credit_limit")
        total_due = summary_fields.get("total_due")
        utilization_percent = None
        is_over_limit = None

        if isinstance(credit_limit, (int, float)) and credit_limit > 0 and isinstance(total_due, (int, float)):
            utilization_percent = round((total_due / credit_limit) * 100, 2)
            is_over_limit = total_due > credit_limit

        return {
            "total_transactions": txn_summary.get("total_transactions", 0),
            "total_spent": txn_summary.get("total_spent", 0.0),
            "utilization_percent": utilization_percent,
            "is_over_limit": is_over_limit,
        }

    # -------------------------------
    # MAIN PARSE
    # -------------------------------
    def parse(self) -> Dict[str, Any]:
        period = self.extract_statement_period()
        transactions = self.extract_transactions()
        txn_summary = self.build_summary(transactions)

        metadata = {
            "bank_name": self.extract_bank_name(),
            "customer_name": self.extract_customer_name(),
            "card_number_masked": self.extract_card_number(),
            "statement_date": self.extract_statement_date(),
            "statement_period_start": period.get("start"),
            "statement_period_end": period.get("end"),
        }

        summary = {
            "total_due": self.extract_total_due(),
            "due_date": self.extract_due_date(),
            "minimum_due": self.extract_minimum_due(),
            "previous_balance": self.extract_previous_balance(),
            "payments_credits": self.extract_payments_credits(),
            "credit_limit": self.extract_credit_limit(),
            "available_credit": self.extract_available_credit(),
            "retail_purchases": self.extract_retail_purchases(),
            "finance_charges": self.extract_finance_charges(),
            "fees_taxes": self.extract_fees_taxes(),
        }

        return {
            "metadata": metadata,
            "summary": summary,
            "transactions": transactions,
            "analytics": {
                "summary": txn_summary,
                "monthly_spend": self.build_monthly_spend(transactions),
                "category_spend": self.build_category_spend(transactions),
                "suspicious_transactions": self.flag_suspicious_transactions(transactions),
            },
            "derived": self.build_derived(summary, txn_summary),
        }


# =========================================================
# PARSING HELPERS
# =========================================================
def parse_single_pdf(file_path: str) -> Dict[str, Any]:
    text = extract_text_from_pdf(file_path)
    parser = StatementParser(text)
    result = parser.parse()
    result["source_file"] = file_path
    return result


def parse_multiple_pdfs(file_paths: List[str]) -> Dict[str, Any]:
    accounts: List[Dict[str, Any]] = []
    errors: List[Dict[str, str]] = []

    for file_path in file_paths:
        file_path = file_path.strip()
        if not file_path:
            continue

        try:
            parsed = parse_single_pdf(file_path)
            accounts.append(parsed)
        except Exception as e:
            errors.append({
                "source_file": file_path,
                "error": str(e)
            })

    return {
        "accounts": accounts,
        "errors": errors,
    }


def parse_single_text(text: str) -> Dict[str, Any]:
    parser = StatementParser(text)
    parsed = parser.parse()
    parsed["source_file"] = "raw_text_input"
    return parsed


# =========================================================
# OUTPUT HELPERS
# =========================================================
def save_as_json(data: Dict[str, Any], filename: str = "output.json") -> None:
    with open(filename, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=2)
    print(f"✅ Saved as {filename}")


def save_as_pdf(data: Dict[str, Any], filename: str = "output_report.pdf") -> None:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import (
            SimpleDocTemplate,
            Paragraph,
            Spacer,
            Table,
            TableStyle,
            PageBreak,
        )
    except ImportError:
        raise ImportError("Install reportlab using: pip install reportlab")

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CustomTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#173A63"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionHeading",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            textColor=colors.HexColor("#173A63"),
            spaceBefore=8,
            spaceAfter=6,
        )
    )

    story = []
    story.append(Paragraph("Credit Card Statement Analysis Report", styles["CustomTitle"]))
    story.append(Spacer(1, 8))

    for idx, account in enumerate(data.get("accounts", []), start=1):
        metadata = account.get("metadata", {})
        summary = account.get("summary", {})
        derived = account.get("derived", {})
        analytics = account.get("analytics", {})
        txn_summary = analytics.get("summary", {})

        story.append(Paragraph(f"Account {idx}", styles["SectionHeading"]))

        info_box = [
            ["Source File", account.get("source_file", "Not available")],
            ["Bank Name", metadata.get("bank_name") or "Not found"],
            ["Customer Name", metadata.get("customer_name") or "Not found"],
            ["Card Number", metadata.get("card_number_masked") or "Not found"],
            ["Statement Date", metadata.get("statement_date") or "Not found"],
            [
                "Statement Period",
                f"{metadata.get('statement_period_start') or 'NA'} to {metadata.get('statement_period_end') or 'NA'}",
            ],
            ["Total Due", f"Rs. {summary.get('total_due', 0):,.2f}" if summary.get("total_due") is not None else "Not found"],
            ["Due Date", summary.get("due_date") or "Not found"],
            ["Minimum Due", f"Rs. {summary.get('minimum_due', 0):,.2f}" if summary.get("minimum_due") is not None else "Not found"],
            ["Credit Limit", f"Rs. {summary.get('credit_limit', 0):,.2f}" if summary.get("credit_limit") is not None else "Not found"],
            ["Available Credit", f"Rs. {summary.get('available_credit', 0):,.2f}" if summary.get("available_credit") is not None else "Not found"],
            ["Utilization %", str(derived.get("utilization_percent"))],
            ["Over Limit", str(derived.get("is_over_limit"))],
            ["Total Transactions", str(txn_summary.get("total_transactions", 0))],
            ["Total Spent", f"Rs. {txn_summary.get('total_spent', 0):,.2f}"],
        ]

        info_table = Table(info_box, colWidths=[55 * mm, 110 * mm])
        info_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(info_table)
        story.append(Spacer(1, 8))

        story.append(Paragraph("Category Spend Breakdown", styles["SectionHeading"]))
        category_data = [["Category", "Amount (Rs.)"]]
        for category, amount in analytics.get("category_spend", {}).items():
            category_data.append([category, f"{amount:,.2f}"])

        category_table = Table(category_data, colWidths=[70 * mm, 50 * mm])
        category_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7F4E4")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(category_table)
        story.append(Spacer(1, 8))

        suspicious = analytics.get("suspicious_transactions", [])
        story.append(Paragraph("Suspicious Transactions", styles["SectionHeading"]))
        if suspicious:
            suspicious_data = [["Date", "Merchant", "Amount", "Reason"]]
            for txn in suspicious:
                suspicious_data.append([
                    txn["date"],
                    txn["merchant"],
                    f"{txn['amount']:,.2f}",
                    txn["reason"],
                ])

            suspicious_table = Table(
                suspicious_data,
                colWidths=[25 * mm, 60 * mm, 28 * mm, 50 * mm],
                repeatRows=1,
            )
            suspicious_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FDE3E3")),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("PADDING", (0, 0), (-1, -1), 5),
                    ]
                )
            )
            story.append(suspicious_table)
        else:
            story.append(Paragraph("No suspicious transactions detected.", styles["BodyText"]))

        story.append(PageBreak())

    errors = data.get("errors", [])
    if errors:
        story.append(Paragraph("Errors", styles["SectionHeading"]))
        err_rows = [["Source File", "Error"]]
        for err in errors:
            err_rows.append([err.get("source_file", ""), err.get("error", "")])

        err_table = Table(err_rows, colWidths=[60 * mm, 100 * mm], repeatRows=1)
        err_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FDE3E3")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("PADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.append(err_table)

    doc.build(story)
    print(f"✅ Saved as {filename}")


# =========================================================
# MAIN PROGRAM
# =========================================================
def main() -> None:
    print("\n=== Credit Card Statement Parser ===")
    print("1. Enter raw text")
    print("2. Upload PDF file")
    print("3. Fetch statement PDF from email (IMAP)")

    choice = input("Choose input type (1 / 2 / 3): ").strip()
    multiple_choice = input("Do you want to process multiple inputs? (y/n): ").strip().lower()

    result: Dict[str, Any] = {"accounts": [], "errors": []}

    if choice == "1":
        if multiple_choice == "y":
            print("\nPaste each statement text below.")
            print("For each statement, press ENTER on an empty line to finish that statement.")
            print("Type DONE on a new line when finished with all statements.\n")

            while True:
                print("Enter a statement (or type DONE to finish all):")
                first_line = input().strip()
                if first_line.upper() == "DONE":
                    break

                lines = [first_line]
                while True:
                    line = input()
                    if line.strip() == "":
                        break
                    lines.append(line)

                text = "\n".join(lines)

                try:
                    parsed = parse_single_text(text)
                    result["accounts"].append(parsed)
                except Exception as e:
                    result["errors"].append({
                        "source_file": "raw_text_input",
                        "error": str(e)
                    })
        else:
            print("\nPaste your statement text below.")
            print("Press ENTER on an empty line to finish:\n")

            lines = []
            while True:
                line = input()
                if line.strip() == "":
                    break
                lines.append(line)

            text = "\n".join(lines)

            try:
                parsed = parse_single_text(text)
                result["accounts"].append(parsed)
            except Exception as e:
                result["errors"].append({
                    "source_file": "raw_text_input",
                    "error": str(e)
                })

    elif choice == "2":
        if multiple_choice == "y":
            pdf_input = input("Enter PDF file paths separated by commas:\n").strip()
            file_paths = [p.strip() for p in pdf_input.split(",") if p.strip()]

            if not file_paths:
                print("❌ No PDF file paths provided.")
                return

            result = parse_multiple_pdfs(file_paths)
        else:
            file_path = input("Enter PDF file path: ").strip()

            try:
                parsed = parse_single_pdf(file_path)
                result["accounts"].append(parsed)
            except Exception as e:
                result["errors"].append({
                    "source_file": file_path,
                    "error": str(e)
                })

    elif choice == "3":
        print("\nFor Gmail, use an App Password, not your normal password.")
        email_address = input("Enter email address: ").strip()
        app_password = input("Enter email app password: ").strip()
        sender_input = input("Optional sender filter (comma-separated, press Enter to skip): ").strip()
        subject_input = input("Optional subject keywords (comma-separated, press Enter for defaults): ").strip()

        sender_filters = [s.strip().lower() for s in sender_input.split(",") if s.strip()]
        subject_keywords = [s.strip().lower() for s in subject_input.split(",") if s.strip()]
        if not subject_keywords:
            subject_keywords = None

        downloaded_ids = set()

        if multiple_choice == "y":
            while True:
                try:
                    email_fetch_meta = fetch_pdf_from_email_imap(
                        email_address=email_address,
                        app_password=app_password,
                        sender_filters=sender_filters,
                        subject_keywords=subject_keywords,
                        already_downloaded=downloaded_ids,
                    )
                except Exception as e:
                    result["errors"].append({
                        "source_file": "email_input",
                        "error": str(e)
                    })
                    break

                if not email_fetch_meta:
                    print("No more suitable statement PDFs found.")
                    break

                print("\n✅ Statement PDF downloaded from email")
                print(f"Subject: {email_fetch_meta['email_subject']}")
                print(f"Sender: {email_fetch_meta['email_sender']}")
                print(f"Attachment: {email_fetch_meta['attachment_filename']}")
                print(f"Saved Path: {email_fetch_meta['saved_pdf_path']}")

                try:
                    parsed = parse_single_pdf(email_fetch_meta["saved_pdf_path"])
                    parsed["email_source"] = email_fetch_meta
                    result["accounts"].append(parsed)
                except Exception as e:
                    result["errors"].append({
                        "source_file": email_fetch_meta.get("saved_pdf_path", "email_input"),
                        "error": str(e)
                    })

                another = input("Fetch another statement from email? (y/n): ").strip().lower()
                if another != "y":
                    break
        else:
            try:
                email_fetch_meta = fetch_pdf_from_email_imap(
                    email_address=email_address,
                    app_password=app_password,
                    sender_filters=sender_filters,
                    subject_keywords=subject_keywords,
                    already_downloaded=downloaded_ids,
                )
            except Exception as e:
                print(f"❌ Error: {e}")
                return

            if not email_fetch_meta:
                print("❌ No suitable statement PDF attachment found in recent emails.")
                return

            print("\n✅ Statement PDF downloaded from email")
            print(f"Subject: {email_fetch_meta['email_subject']}")
            print(f"Sender: {email_fetch_meta['email_sender']}")
            print(f"Attachment: {email_fetch_meta['attachment_filename']}")
            print(f"Saved Path: {email_fetch_meta['saved_pdf_path']}")

            try:
                parsed = parse_single_pdf(email_fetch_meta["saved_pdf_path"])
                parsed["email_source"] = email_fetch_meta
                result["accounts"].append(parsed)
            except Exception as e:
                result["errors"].append({
                    "source_file": email_fetch_meta.get("saved_pdf_path", "email_input"),
                    "error": str(e)
                })

    else:
        print("❌ Invalid choice.")
        return

    print("\n=== Parsed Output ===")
    print(json.dumps(result, indent=2))

    print("\nSave Options:")
    print("1. Save as JSON")
    print("2. Save as PDF report")
    print("3. Save both JSON and PDF")
    print("4. Skip saving")

    save_choice = input("Choose option (1/2/3/4): ").strip()

    if save_choice == "1":
        save_as_json(result, "output.json")
    elif save_choice == "2":
        save_as_pdf(result, "output_report.pdf")
    elif save_choice == "3":
        save_as_json(result, "output.json")
        save_as_pdf(result, "output_report.pdf")
    elif save_choice == "4":
        print("Skipped saving.")
    else:
        print("Invalid option. Skipped saving.")


if __name__ == "__main__":
    main()