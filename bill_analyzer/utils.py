"""
File operation utilities for backup and restore
"""

import datetime as dt
import json
import re
import shutil
from pathlib import Path
from typing import Any, cast

import dateutil.parser as dparser


def parse_json_from_markdown(text: str) -> dict[str, Any]:
    """Extract and parse JSON from a markdown code block or plain JSON string.

    Handles both formats:
    - Markdown: ```json\\n{...}\\n```
    - Plain JSON: {...}

    :param text: Text containing JSON, possibly wrapped in markdown code blocks
    :type text: str
    :return: Parsed JSON as a dictionary
    :rtype: dict[str, Any]
    :raises json.JSONDecodeError: If the JSON is malformed
    """
    # Try to extract JSON from markdown code block
    json_match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)

    if json_match:
        json_str = json_match.group(1)
    else:
        # If no markdown block found, assume the entire text is JSON
        json_str = text

    result = json.loads(json_str.strip())
    return cast(dict[str, Any], result)


def create_backup(file_path: str) -> str:
    """Create a timestamped backup of a file.

    :param file_path: Path to the file to backup
    :type file_path: str
    :return: Path to the backup file
    :rtype: str
    """
    path = Path(file_path)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_stem(f"{path.stem}_backup_{timestamp}")
    shutil.copy2(file_path, backup_path)
    print(f"Creating backup: {backup_path}")
    return str(backup_path)


def restore_from_backup(backup_path: str, target_path: str) -> None:
    """Restore a file from its backup.

    :param backup_path: Path to the backup file
    :type backup_path: str
    :param target_path: Path to restore to
    :type target_path: str
    """
    print("Restoring from backup...")
    shutil.copy2(backup_path, target_path)
    print("✓ Restored from backup")


def remove_backup(backup_path: str) -> None:
    """Remove a backup file if it exists.

    :param backup_path: Path to the backup file
    :type backup_path: str
    """
    path = Path(backup_path)
    if path.exists():
        path.unlink()
        print("✓ Removed backup file")


def parse_date(date_value: str) -> str | None:
    """Parse a date string into ISO format (YYYY-MM-DD).

    Attempts to parse dates in multiple formats:
    - ISO format (with or without timezone, e.g., "2025-12-19" or "2025-12-19Z")
    - European format with day first (e.g., "10.12.25")

    :param date_value: Date string to parse
    :type date_value: str
    :return: Date in ISO format (YYYY-MM-DD) or None if parsing fails
    :rtype: str | None
    """
    try:
        if not isinstance(date_value, str) or not date_value.strip():
            return None

        try:
            row_date: dt.date = dt.date.fromisoformat(date_value.replace("Z", "+00:00"))
            row_date_str: str = row_date.strftime("%Y-%m-%d")
        except (TypeError, ValueError):
            # Parse non-ISO format (e.g., "10.12.25") with dayfirst=True
            row_date = dparser.parse(date_value, dayfirst=True).date()
            row_date_str = row_date.strftime("%Y-%m-%d")

        return row_date_str
    except (ValueError, TypeError, OverflowError):
        return None


def is_number(string: str) -> bool:
    """Check if a string can be converted to a number.

    :param string: String to check for numeric value
    :type string: str
    :return: True if the string represents a valid number, False otherwise
    :rtype: bool
    """
    try:
        float(string)
        return True
    except ValueError:
        return False


def extract_price_number(string: str) -> str:
    """Extract numeric value from a price string with currency symbols.

    Handles German number formats (comma as decimal separator).
    Examples:
    - "13,5 €" -> "13.5"
    - "1.234,56€" -> "1234.56"
    - "-42,99 EUR" -> "-42.99"

    :param string: Price string with optional currency symbols and formatting
    :type string: str
    :return: Numeric string parseable by float()
    :rtype: str
    """
    # Remove whitespace and common currency symbols
    price = re.sub(r"[€$£¥\s]", "", string)

    # Keep only digits, comma, and minus sign
    price = re.sub(r"[^\d,\-]", "", price)

    # Replace commas with periods
    price = price.replace(",", ".")

    return price
