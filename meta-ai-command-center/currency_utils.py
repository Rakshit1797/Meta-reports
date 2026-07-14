#!/usr/bin/env python3
"""
Shared currency formatting helpers for the Meta Ads reporting pipeline.

The Meta ad account's currency is collected in meta_collector.py (a
best-effort call to the ad account's own `currency` field), carried through
collection_metadata.json and then performance_findings.json, and used here
so Excel and email output never hardcode a currency symbol -- and never
assume USD for a non-USD account (e.g. an INR-denominated account).
"""

from typing import Optional

# Curated symbol map for common currencies. This is intentionally small --
# any currency not listed here still gets correct, non-USD-assuming
# treatment via the ISO-code fallback in the functions below.
CURRENCY_SYMBOLS = {
    "USD": "$",
    "INR": "₹",
    "EUR": "€",
    "GBP": "£",
    "JPY": "¥",
}


def get_currency_symbol(currency_code: Optional[str]) -> Optional[str]:
    """Return the known symbol for a currency code, or None if not mapped."""
    if not currency_code:
        return None
    return CURRENCY_SYMBOLS.get(currency_code.upper())


def format_currency_text(amount: Optional[float], currency_code: Optional[str]) -> str:
    """Format a number as currency text for plain-text/HTML output (e.g. email).

    Never assumes USD:
    - Known symbol (e.g. INR)  -> "₹10,282.30"
    - Known code, no symbol    -> "10,282.30 AED"
    - No currency code at all  -> "10,282.30" (no fabricated currency)
    """
    if amount is None:
        return "N/A"

    symbol = get_currency_symbol(currency_code)
    if symbol:
        return f"{symbol}{amount:,.2f}"
    if currency_code:
        return f"{amount:,.2f} {currency_code.upper()}"
    return f"{amount:,.2f}"


def get_currency_excel_number_format(currency_code: Optional[str]) -> str:
    """Build an Excel custom number_format string for the given currency.

    Mirrors format_currency_text()'s fallback behavior exactly:
    - Known symbol   -> literal symbol prefix, e.g. '"₹"#,##0.00'
    - Known code only -> trailing literal ISO code, e.g. '#,##0.00" AED"'
    - No code at all  -> plain number, no currency assumed
    """
    symbol = get_currency_symbol(currency_code)
    if symbol:
        return f'"{symbol}"#,##0.00'
    if currency_code:
        return f'#,##0.00" {currency_code.upper()}"'
    return '#,##0.00'
