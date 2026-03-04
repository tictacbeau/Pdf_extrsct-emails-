"""
parser.py — Transaction amount extraction from PDFs, Excel files, and plain text.

Strategy:
1. Route by file extension to the appropriate extractor.
2. Within text, prefer amounts found on lines containing total-related keywords.
3. Return the largest amount found, or None.
"""

import re
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

DOLLAR_REGEX = re.compile(r'\$\s*[\d,]+(?:\.\d{2})?')

TOTAL_KEYWORDS = frozenset({
    "total", "amount", "grand total", "order total",
    "subtotal", "balance due", "amount due",
    "invoice total", "net total", "total amount",
    "total due", "payment due",
})


def extract_amount_from_file(file_path: str):
    """
    Dispatcher: routes to the correct extractor based on file extension.
    Returns float or None.
    """
    ext = Path(file_path).suffix.lower()
    try:
        if ext == ".pdf":
            return extract_amount_from_pdf(file_path)
        elif ext in (".xlsx", ".xls"):
            if ext == ".xls":
                logger.warning(
                    "File %s is .xls format (old Excel). "
                    "openpyxl may not support it; attempting anyway.",
                    file_path
                )
            return extract_amount_from_excel(file_path)
        else:
            # .txt, .html, .htm, .csv, and anything else: read as text
            try:
                with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                    text = f.read()
            except OSError as e:
                logger.warning("Could not read file %s: %s", file_path, e)
                return None
            return extract_amount_from_text(text)
    except Exception as e:
        logger.warning("Amount extraction failed for %s: %s", file_path, e)
        return None


def extract_amount_from_pdf(file_path: str):
    """
    Use pdfplumber to extract text page by page, then run regex on combined text.
    Handles encrypted PDFs gracefully.
    Returns float or None.
    """
    try:
        import pdfplumber
    except ImportError:
        logger.warning("pdfplumber not installed; cannot parse PDF %s", file_path)
        return None

    try:
        with pdfplumber.open(file_path) as pdf:
            pages_text = []
            for page in pdf.pages:
                try:
                    text = page.extract_text()
                    if text:
                        pages_text.append(text)
                except Exception as e:
                    logger.debug("Failed to extract text from page in %s: %s", file_path, e)
            combined = "\n".join(pages_text)
        return extract_amount_from_text(combined)
    except Exception as e:
        logger.warning("PDF parsing failed for %s: %s", file_path, e)
        return None


def extract_amount_from_excel(file_path: str):
    """
    Use openpyxl with read_only=True, data_only=True.

    Strategy (in priority order):
    1. KEYWORD ROW SCAN: If any cell in a row matches a TOTAL_KEYWORD,
       look for a numeric cell in that row (or adjacent). Collect amounts.
    2. REGEX FALLBACK: If no keyword match, scan all cells with the dollar regex.

    Returns the largest amount found across all sheets, or None.
    """
    try:
        import openpyxl
    except ImportError:
        logger.warning("openpyxl not installed; cannot parse Excel %s", file_path)
        return None

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    except Exception as e:
        logger.warning("Failed to open Excel file %s: %s", file_path, e)
        return None

    keyword_amounts = []
    fallback_amounts = []

    try:
        for sheet in wb.worksheets:
            try:
                for row in sheet.iter_rows():
                    row_has_keyword = False
                    row_values = []

                    for cell in row:
                        val = cell.value
                        row_values.append(val)
                        if isinstance(val, str):
                            lower = val.lower().strip()
                            if any(kw in lower for kw in TOTAL_KEYWORDS):
                                row_has_keyword = True

                    if row_has_keyword:
                        # Look for numeric values in the same row
                        for val in row_values:
                            if isinstance(val, (int, float)) and val >= 0:
                                keyword_amounts.append(float(val))
                            elif isinstance(val, str):
                                amt = extract_amount_from_text(val)
                                if amt is not None:
                                    keyword_amounts.append(amt)
                    else:
                        # Fallback: regex on string cells
                        for val in row_values:
                            if isinstance(val, str) and val.strip():
                                amt = extract_amount_from_text(val)
                                if amt is not None:
                                    fallback_amounts.append(amt)
            except Exception as e:
                logger.debug("Error reading sheet in %s: %s", file_path, e)
    finally:
        wb.close()

    if keyword_amounts:
        return find_largest_amount(keyword_amounts)
    if fallback_amounts:
        return find_largest_amount(fallback_amounts)
    return None


def extract_amount_from_text(text: str):
    """
    1. Apply DOLLAR_REGEX to find all dollar matches.
    2. Prefer amounts on lines containing a TOTAL_KEYWORD.
    3. Return max of keyword-line amounts if any; else max of all amounts.
    4. Return None if no matches.
    """
    if not text:
        return None

    all_amounts = []
    keyword_amounts = []

    for line in text.splitlines():
        line_lower = line.lower()
        line_has_keyword = any(kw in line_lower for kw in TOTAL_KEYWORDS)
        matches = DOLLAR_REGEX.findall(line)
        for m in matches:
            amount = parse_dollar_string(m)
            if amount is not None:
                all_amounts.append(amount)
                if line_has_keyword:
                    keyword_amounts.append(amount)

    if keyword_amounts:
        return find_largest_amount(keyword_amounts)
    return find_largest_amount(all_amounts)


def parse_dollar_string(dollar_str: str):
    """
    Convert a regex match string like "$1,234.56" or "$ 999" to float.
    Returns None if conversion fails.
    """
    try:
        cleaned = dollar_str.replace("$", "").replace(",", "").strip()
        return float(cleaned)
    except (ValueError, AttributeError):
        return None


def find_largest_amount(amounts: list):
    """Return max(amounts) or None if list is empty."""
    if not amounts:
        return None
    return max(amounts)
