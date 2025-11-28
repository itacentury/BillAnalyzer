"""
Configuration constants for the AI Bill Analyzer
"""

# ==============================================================================
# FILE PATHS
# ==============================================================================

ODS_FILE = "/home/juli/Downloads/Alltags-Ausgaben.ods"


# ==============================================================================
# CLAUDE API CONFIGURATION
# ==============================================================================

# CLAUDE_MODEL = "claude-sonnet-4-5-20250929"
CLAUDE_MODEL = "claude-opus-4-5-20251101"
CLAUDE_MAX_TOKENS = 2048


# ==============================================================================
# ODS COLUMN INDICES
# ==============================================================================

COL_DATE = 1
COL_STORE = 2
COL_ITEM = 3
COL_PRICE = 4
COL_TOTAL = 5


# ==============================================================================
# EXTRACTION PROMPT
# ==============================================================================

EXTRACTION_PROMPT = """Bitte extrahiere folgende Daten aus der Rechnung:
1. Name des Supermarkts, ohne Gewerbeform o.ä., also nur 'REWE' oder 'Edeka'.
2. Datum ohne Uhrzeit.
3. Alle Artikel inklusive Preis, Artikel in korrekter deutschen Groß- und Kleinschreibung.
4. Gesamtpreis.

Wenn der gleiche Artikel mehrfach gekauft wurde, dann schreibe als Preis für den Artikel: Anzahl * Einzelpreis (z.B. '=4*0,59').
Wenn ein Artikel Pfand hat, dann schreibe als Preis für den Artikel: Artikelpreis + Pfand (z.B. '=0,89+0,08').
Schreibe das Gewicht bei zum Beispiel Gemüse oder Obst, hinten an den Namen des dazugehörigen Gemüse oder Obstes.

Gebe mir die Daten im JSON-Format zurück, mit folgenden Namen und Datentypen:
'store' (str), 'date' (str), 'items' (list[dict[str, str]]), 'total' (float)."""


# ==============================================================================
# ODS NAMESPACES
# ==============================================================================

CALCEXT_NS = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"


# ==============================================================================
# ODS ATTRIBUTE LISTS FOR CELL CLEARING
# ==============================================================================

OFFICE_ATTRS_TO_CLEAR = [
    "value",
    "date-value",
    "time-value",
    "boolean-value",
    "string-value",
    "value-type",
    "currency",
]

TABLE_ATTRS_TO_CLEAR = ["formula"]

CALCEXT_ATTRS_TO_CLEAR = ["value-type"]
