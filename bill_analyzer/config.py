"""
Configuration constants for the AI Bill Analyzer
"""

import os
from pathlib import Path
from typing import Final

from dotenv import load_dotenv

load_dotenv(".env.bill_analyzer")

# ==============================================================================
# MISC CONFIGURATION
# ==============================================================================

EXPORT_JSON_PATH: Final[Path] = Path().home() / "Downloads"

# ==============================================================================
# CLAUDE API CONFIGURATION
# ==============================================================================

CLAUDE_MODEL: Final[str] = "claude-opus-4-5-20251101"
CLAUDE_MAX_TOKENS: Final[int] = 2048

# ==============================================================================
# PAPERLESS-NGX API CONFIGURATION
# ==============================================================================

PAPERLESS_UPLOAD_ENABLE: Final[bool] = True
PAPERLESS_URL: Final[str | None] = os.environ.get("PAPERLESS_URL")
PAPERLESS_TOKEN: Final[str | None] = os.environ.get("PAPERLESS_API_TOKEN")
PAPERLESS_TOTAL_ID: Final[int] = 1

# ==============================================================================
# EXTRACTION PROMPT
# ==============================================================================

EXTRACTION_PROMPT: Final[
    str
] = """Bitte extrahiere folgende Daten aus der Rechnung:
1. Name des Supermarkts, ohne Gewerbeform o.ä., also nur 'REWE' oder 'Edeka'.
2. Die Kategorie des Einkaufs. Zum Beispiel 'Lebensmittel', 'Restaurant' oder 'Elektronik'.
3. Datum ohne Uhrzeit im ISO-8601 Format.
4. Gesamtpreis.
5. Alle Artikel inklusive Preis, Artikel in korrekter deutschen Groß- und Kleinschreibung.

Wenn der gleiche Artikel mehrfach gekauft wurde, dann schreibe als Preis für den Artikel den zusammengerechneten Preis
und füge auch die Anzahl vor dem Artikelnamen hinzu (z.B. '4x Semmel'), außer bei Pfand.
Wenn ein Artikel Pfand hat, addiere den dazugehörigen Pfand zum Preis des Artikels.
Wenn ich bei einem Einkauf Pfand zurückgebe, behandle dies als Negativzahl und fasse jeden zurückgegebenen Pfand unter einem einzigen Eintrag 'Pfand' zusammen.
Schreibe das Gewicht bei zum Beispiel Gemüse oder Obst, hinten an den Namen des dazugehörigen Gemüse oder Obstes.

Gebe mir die Daten im JSON-Format zurück, mit folgenden Namen und Datentypen:
'store' (str), 'category' (str), 'date' (str), 'items' (list[dict[str, str | float]]), 'total' (float).
Die Keys für das 'items' dictionary sollen 'item_name' und 'item_price' heißen.
Die Values von 'item_name' sind strings und die values von 'item_price' sind floats.

Wenn es sich bei der PDF um keine Rechnung handelt, gebe 'error' zurück.
"""
