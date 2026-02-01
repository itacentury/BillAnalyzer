"""
AI Bill Analyzer - Extract and organize bill data from PDFs into ODS spreadsheets

This tool uses Claude AI to extract structured data from bill PDFs and automatically
inserts them into an ODS spreadsheet with proper formatting preservation.
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

from .claude_api import analyze_bill_pdf
from .config import EXPORT_JSON_PATH, PAPERLESS_TOKEN, PAPERLESS_TOTAL_ID, PAPERLESS_URL
from .paperless_api import upload_to_paperless
from .ui import select_pdf_files
from .utils import parse_json_from_markdown
from .validators import validate_bill_total


def convert_date_to_iso8601(date_str: str | None) -> str | None:
    """Convert various date formats to ISO 8601 datetime format.

    Supports: YYYY-MM-DD, DD.MM.YYYY, DD.MM.YY, YYYYMMDD
    """
    if not date_str:
        return None

    try:
        # Try parsing as YYYY-MM-DD first
        if "-" in date_str:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
        # Try DD.MM.YYYY or DD.MM.YY format
        elif "." in date_str:
            parts = date_str.split(".")
            # Check if year is 2 or 4 digits
            if len(parts[2]) == 4:
                dt = datetime.strptime(date_str, "%d.%m.%Y")
            else:
                dt = datetime.strptime(date_str, "%d.%m.%y")
        else:
            dt = datetime.strptime(date_str, "%Y%m%d")

        # Format as ISO 8601 with midnight time
        return dt.strftime("%Y-%m-%dT00:00:00Z")
    except ValueError:
        print(f"  ⚠ Warning: Could not parse date '{date_str}'")
        return None


def upload_bills_to_paperless(
    valid_pdfs: list[str], valid_bills: list[dict[str, Any]]
) -> None:
    """Upload PDF bills to Paperless-ngx with metadata."""
    if not (PAPERLESS_TOKEN and PAPERLESS_URL):
        print("\n⚠ Paperless is not configured! Upload will be skipped.")
        return

    print("\n📤 Uploading to Paperless-ngx...")

    for pdf, bill in zip(valid_pdfs, valid_bills):
        try:
            # Create a title from store and date
            title: str = f"{bill.get('store', 'Bill')}"

            # Get total price for custom field
            total_price: float = bill.get("total", 0.0)

            # Format date for Paperless (requires ISO 8601 format)
            created_datetime: str | None = convert_date_to_iso8601(bill.get("date"))

            # Upload the PDF
            task_uuid: str = upload_to_paperless(
                pdf_path=pdf,
                token=PAPERLESS_TOKEN,
                paperless_url=PAPERLESS_URL,
                title=title,
                created=created_datetime,
                custom_fields={PAPERLESS_TOTAL_ID: total_price},
            )

            print(f"✓ Uploaded successfully (Task UUID: {task_uuid})")

        except requests.HTTPError as e:
            print(f"⚠ Paperless upload failed: {e}")
            # Print detailed error response from Paperless
            if hasattr(e, "response") and e.response is not None:
                try:
                    error_details = e.response.json()
                    print(f"  Error details: {error_details}")
                except (ValueError, KeyError):
                    print(f"  Response text: {e.response.text}")
        except requests.RequestException as e:
            print(f"⚠ Paperless upload failed: {e}")
        except FileNotFoundError as e:
            print(f"⚠ PDF file not found: {e}")


def validate_and_print_bill(bill_data: dict[str, Any]) -> bool:
    """Validate bill total and print validation results.
    Returns true if bill data is valid otherwise false.
    """
    try:
        validation_result: dict[str, bool | float | str] = validate_bill_total(
            bill_data
        )
        print(f"\n{validation_result['message']}")

        if not validation_result["valid"]:
            print(f"  Calculated sum: {validation_result['calculated_sum']}€")
            print(f"  Declared total: {validation_result['declared_total']}€")
            print(f"  Difference: {validation_result['difference']}€")
            print("  ⚠ Warning: Price validation failed - data may be incorrect!")

        return bool(validation_result["valid"])
    except (KeyError, ValueError) as e:
        print(f"⚠ Validation error: {e}")
        return False


def save_bills_to_json(data: list[dict[str, Any]]) -> None:
    """Saves extracted items to a json file."""
    filename: str = f"bills-{datetime.now().date()}.json"
    filepath: Path = EXPORT_JSON_PATH / filename
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=False)


def print_statistics(valid_bills_length: int, all_bills_length: int) -> None:
    """Print summary statistics for processed bills."""

    print("\n" + "=" * 50)
    print("📊 STATISTICS")
    print("=" * 50)

    print(f"\n📄 Bills processed:     {valid_bills_length} of {all_bills_length}")


def main() -> None:
    """Main entry point."""
    print("=== AI BILL ANALYZER ===\n")

    pdfs: tuple[str, ...] = select_pdf_files()
    if not pdfs:
        print("No files selected.")
        return

    valid_bills: list[dict[str, Any]] = []
    valid_pdfs: list[str] = []
    for pdf in pdfs:
        print(f"\nAnalyzing: {pdf}")

        response: str | None = analyze_bill_pdf(pdf)

        if response is None:
            continue

        bill_data: dict[str, Any] = parse_json_from_markdown(response)
        print(json.dumps(bill_data, indent=2, ensure_ascii=False))

        is_valid = validate_and_print_bill(bill_data)

        if not is_valid:
            continue

        valid_bills.append(bill_data)
        valid_pdfs.append(pdf)

    save_bills_to_json(valid_bills)
    upload_bills_to_paperless(valid_pdfs, valid_bills)
    print_statistics(len(valid_bills), len(pdfs))


if __name__ == "__main__":
    main()
