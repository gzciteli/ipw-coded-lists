"""
Date validation utilities shared by input prompts and housekeeping tasks.
"""

import re
from typing import Tuple

import config


def _format_valid_years() -> str:
    return ", ".join(sorted(config.VALID_YEARS, key=lambda year: (len(year), year)))

_DATE_RE = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})\s*$")


def _parse_date_parts(date_str: str) -> Tuple[int, int, str]:
    """Parse and validate a date string. Returns (month, day, year_str)."""
    match = _DATE_RE.match(date_str)
    if not match:
        raise ValueError(
            "Use format M.D.YY or M.D.YYYY (e.g., 2.2.26 or 02.02.2026)."
        )

    month = int(match.group(1))
    day = int(match.group(2))
    year_str = match.group(3)

    if month < 1 or month > 12:
        raise ValueError("Month must be between 1 and 12.")
    if day < 1 or day > 31:
        raise ValueError("Day must be between 1 and 31.")
    if year_str not in config.VALID_YEARS:
        raise ValueError(f"Year must be one of: {_format_valid_years()}.")

    return month, day, year_str


def validate_and_normalize_date(date_str: str) -> str:
    """
    Validate a user-supplied date string and normalize it.

    Accepted formats:
      - 2.2.26
      - 02.02.2026
      - any 1-2 digits, '.', 1-2 digits, '.', 2 or 4 digits (with constraints)

    Normalization keeps the year length provided, but removes leading zeros
    from month/day.
    """
    month, day, year_str = _parse_date_parts(date_str)
    return f"{month}.{day}.{year_str}"


def looks_like_date_fragment(text: str) -> bool:
    """Return True if the text matches the allowed date pattern and ranges."""
    try:
        _parse_date_parts(text)
        return True
    except ValueError:
        return False
