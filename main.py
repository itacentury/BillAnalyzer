"""
Claude API Examples - Sending Requests and Processing Responses
"""

import base64
import json
import os
import re
import tkinter as tk
from tkinter import filedialog

import anthropic

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
                        1. Name des Supermarkts
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


def main():
    print("=== AI BILL ANALYZER ===\n")

    pdfs = ask_pdfs()
    for pdf in pdfs:
        response = pdf_analysis(pdf)
        data = parse_json_from_markdown(response)
        print(json.dumps(data, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
