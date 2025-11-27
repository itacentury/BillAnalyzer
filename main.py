"""
Claude API Examples - Sending Requests and Processing Responses
"""

import base64
import json
import os
import re
import shutil
import tkinter as tk
from datetime import datetime as dt
from tkinter import filedialog

import anthropic
import dateutil.parser as dparser
import pandas as pd

ODS_FILE = "/home/juli/Downloads/Alltags-Ausgaben.ods"

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))


def parse_json_from_markdown(text):
    # Try to extract JSON from markdown code block (```json ... ``` or ``` ... ```)
    json_match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)

    if json_match:
        json_str = json_match.group(1)
    else:
        # If no markdown block found, assume the entire text is JSON
        json_str = text

    # Parse the JSON string
    return json.loads(json_str.strip())


def pdf_analysis(pdf):
    with open(pdf, "rb") as pdf_file:
        pdf_data = base64.standard_b64encode(pdf_file.read()).decode("utf-8")

    message = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=2048,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_data,
                        },
                    },
                    {
                        "type": "text",
                        "text": """Bitte extrahiere folgende Daten aus der Rechnung:
                        1. Name des Supermarkts, ohne Gewerbeform, also nur 'REWE' oder 'Edeka'
                        2. Datum ohne Uhrzeit
                        3. Alle Artikel inklusive Preis
                        4. Gesamtpreis

                        Gebe mir die Daten im JSON-Format zurück, mit folgenden Namen und Datentypen:
                        'store' (str), 'date' (str), 'item' (list[dict[str, str | float]]), 'total' (float)""",
                    },
                ],
            }
        ],
    )

    response = message.content[0].text
    return response


def ask_pdfs():
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        parent=root,
        title="Select Bills to Analyze",
        filetypes=[("PDF files", "*.pdf")],
        initialdir=os.path.expanduser("~/Downloads"),
    )

    return file_paths


def read_ods(data):
    date_str = data["date"]
    date = dparser.parse(date_str, dayfirst=True)
    month = date.strftime("%b")
    year = date.strftime("%y")
    sheet_name = f"{month} {year}"

    # Read the sheet
    df = pd.read_excel(ODS_FILE, engine="odf", sheet_name=sheet_name)

    # Convert column 1 to date for comparison
    date_column = pd.to_datetime(df.iloc[:, 1], errors="coerce").dt.date

    # Find the row with matching date (column index 1)
    mask = date_column == date.date()
    found_indices = df.index[mask]

    if len(found_indices) == 0:
        print(f"⚠ No row found for date {date.date()}")
        return

    # Get the first matching row index
    row_idx = found_indices[0]

    # Fill the first row with date, store, first item, and first price
    first_item = data["item"][0]
    df.iloc[row_idx, 2] = data["store"]  # Store name in column 2
    df.iloc[row_idx, 3] = first_item["name"]  # First item in column 3
    df.iloc[row_idx, 4] = first_item["price"]  # First price in column 4

    # Create new rows for remaining items
    rows_to_insert = []
    for item in data["item"][1:]:
        # Create a new row with empty strings (not None) for better ODS compatibility
        new_row = pd.Series([""] * len(df.columns), index=df.columns)
        new_row.iloc[3] = item["name"]  # Item name in column 3
        new_row.iloc[4] = item["price"]  # Price in column 4
        rows_to_insert.append(new_row)

    # Add final row with total price
    total_row = pd.Series([""] * len(df.columns), index=df.columns)
    total_row.iloc[4] = data["total"]  # Total in column 4
    rows_to_insert.append(total_row)

    # Insert all new rows after the found row
    if rows_to_insert:
        df_before = df.iloc[: row_idx + 1]
        df_after = df.iloc[row_idx + 1 :]
        df_new_rows = pd.DataFrame(rows_to_insert)
        df = pd.concat([df_before, df_new_rows, df_after], ignore_index=True)

    # Replace NaN with empty strings for ODS compatibility
    df = df.fillna("")

    print(
        f"✓ Inserted {len(data['item'])} items + total for {data['store']} on {date.date()}"
    )

    # Create backup before writing
    backup_file = ODS_FILE.replace(
        ".ods", f"_backup_{dt.now().strftime('%Y%m%d_%H%M%S')}.ods"
    )
    print(f"Creating backup: {backup_file}")
    shutil.copy2(ODS_FILE, backup_file)

    # Write to a temporary file first
    temp_file = ODS_FILE.replace(".ods", "_temp.ods")

    try:
        print("Reading all sheets...")
        all_sheets = pd.read_excel(ODS_FILE, engine="odf", sheet_name=None)
        all_sheets[sheet_name] = df  # Update the modified sheet

        print(f"Writing to temporary file (this may take a while)...")
        with pd.ExcelWriter(temp_file, engine="odf") as writer:
            for name, sheet_df in all_sheets.items():
                print(f"  Writing sheet: {name}")
                sheet_df.to_excel(writer, sheet_name=name, index=False)

        # If successful, replace original with temp file
        print("Finalizing...")
        shutil.move(temp_file, ODS_FILE)
        print(f"✓ Successfully saved to {ODS_FILE}")

    except Exception as e:
        print(f"✗ Error saving file: {e}")
        print(f"Restoring from backup...")
        shutil.copy2(backup_file, ODS_FILE)
        print(f"✓ Restored from backup")
        raise


def main():
    print("=== AI BILL ANALYZER ===\n")

    pdfs = ask_pdfs()
    for pdf in pdfs:
        response = pdf_analysis(pdf)
        data = parse_json_from_markdown(response)
        print(json.dumps(data, indent=2, ensure_ascii=False))

        read_ods(data)


if __name__ == "__main__":
    main()
