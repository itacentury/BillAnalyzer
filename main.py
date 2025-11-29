"""
AI Bill Analyzer - Extract and organize bill data from PDFs into ODS spreadsheets

This tool uses Claude AI to extract structured data from bill PDFs and automatically
inserts them into an ODS spreadsheet with proper formatting preservation.
"""

import json
from typing import Any

import requests

from bill_analyzer.bill_inserter import process_multiple_bills
from bill_analyzer.claude_api import analyze_bill_pdf
from bill_analyzer.config import PAPERLESS_TOKEN, PAPERLESS_URL
from bill_analyzer.json_utils import parse_json_from_markdown
from bill_analyzer.paperless_api import upload_to_paperless
from bill_analyzer.ui import select_pdf_files
from bill_analyzer.validators import validate_bill_total


def main() -> None:
    """Main application entry point."""
    print("=== AI BILL ANALYZER ===\n")

    # Select PDF files
    pdfs: tuple[str, ...] = select_pdf_files()
    if not pdfs:
        print("No files selected.")
        return

    # Analyze all PDFs and collect bill data
    bills_data: list[dict[str, Any]] = []
    for pdf in pdfs:
        print(f"\nAnalyzing: {pdf}")

        # Analyze PDF with Claude
        response: str = analyze_bill_pdf(pdf)

        # Parse JSON from response
        bill_data: dict[str, Any] = parse_json_from_markdown(response)
        print(json.dumps(bill_data, indent=2, ensure_ascii=False))

        # Validate that sum of item prices equals total
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
        except (KeyError, ValueError) as e:
            print(f"⚠ Validation error: {e}")

        # Upload to Paperless-ngx if enabled
        if PAPERLESS_TOKEN:
            try:
                print("\n📤 Uploading to Paperless-ngx...")

                # Create a title from store and date
                title: str = f"{bill_data.get('store', 'Bill')}"

                # Upload the PDF
                result: dict[str, Any] = upload_to_paperless(
                    pdf_path=pdf,
                    token=PAPERLESS_TOKEN,
                    paperless_url=PAPERLESS_URL,
                    title=title,
                    created=bill_data.get("date"),
                )

                print(f"✓ Uploaded successfully (ID: {result.get('id', 'N/A')})")

            except requests.RequestException as e:
                print(f"⚠ Paperless upload failed: {e}")
            except FileNotFoundError as e:
                print(f"⚠ PDF file not found: {e}")

        bills_data.append(bill_data)

    # Insert all bills into ODS in a single batch operation
    if bills_data:
        print("\n" + "=" * 60)
        print(f"Inserting {len(bills_data)} bill(s) into ODS file...")
        print("=" * 60)
        process_multiple_bills(bills_data)


if __name__ == "__main__":
    main()
