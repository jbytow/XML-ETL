"""
Configuration for XML ETL - Comarch Optima invoices.

Each invoice has a header (NAGLOWEK) and multiple line items (POZYCJA).
Excel output: 1 row per line item, with header fields repeated.
"""

from pathlib import Path

# --- Paths ---
BASE_DIR = Path(__file__).parent
XML_DIR = BASE_DIR / "input" / "xml"
OUTPUT_DIR = BASE_DIR / "output"
KARTOTEKA_PATH = BASE_DIR / "input" / "kartoteka z kluczem.xlsx"

# --- XML Namespace ---
XML_NAMESPACE = {"ns": "http://www.cdn.com.pl/optima/dokument"}

# --- Header fields ---
# Extracted once per file (from DOKUMENT/NAGLOWEK).
# XPath is relative to the DOKUMENT element.
HEADER_FIELDS = [
    {"column": "NumerPelny",      "xpath": "ns:NAGLOWEK/ns:NUMER_PELNY"},
    {"column": "DataDokumentu",   "xpath": "ns:NAGLOWEK/ns:DATA_DOKUMENTU"},
    {"column": "DataWystawienia", "xpath": "ns:NAGLOWEK/ns:DATA_WYSTAWIENIA"},
    {"column": "PlatnikKod",      "xpath": "ns:NAGLOWEK/ns:PLATNIK/ns:KOD"},
    {"column": "PlatnikNIP",      "xpath": "ns:NAGLOWEK/ns:PLATNIK/ns:NIP"},
    {"column": "PlatnikNazwa",    "xpath": "ns:NAGLOWEK/ns:PLATNIK/ns:NAZWA"},
    {"column": "FormaPlatnosci",  "xpath": "ns:NAGLOWEK/ns:PLATNOSC/ns:FORMA"},
    {"column": "RazemNetto",      "xpath": "ns:NAGLOWEK/ns:KWOTY/ns:RAZEM_NETTO"},
    {"column": "RazemBrutto",     "xpath": "ns:NAGLOWEK/ns:KWOTY/ns:RAZEM_BRUTTO"},
    {"column": "RazemVAT",        "xpath": "ns:NAGLOWEK/ns:KWOTY/ns:RAZEM_VAT"},
]

# --- Item fields ---
# Extracted per POZYCJA element. XPath is relative to each POZYCJA.
ITEM_FIELDS = [
    {"column": "LP",              "xpath": "ns:LP"},
    {"column": "TowarKod",        "xpath": "ns:TOWAR/ns:KOD"},
    {"column": "TowarNazwa",      "xpath": "ns:TOWAR/ns:NAZWA"},
    {"column": "TowarEAN",        "xpath": "ns:TOWAR/ns:EAN"},
    {"column": "NumerKatalogowy", "xpath": "ns:TOWAR/ns:NUMER_KATALOGOWY"},
    {"column": "Ilosc",           "xpath": "ns:ILOSC"},
    {"column": "JM",              "xpath": "ns:JM"},
    {"column": "StawkaVAT",      "xpath": "ns:STAWKA_VAT/ns:STAWKA"},
    {"column": "WartoscNetto",    "xpath": "ns:WARTOSC_NETTO"},
    {"column": "WartoscBrutto",   "xpath": "ns:WARTOSC_BRUTTO"},
    {"column": "CenaBrutto",      "xpath": "ns:CENY/ns:PO_RABACIE_WAL_DOKUMENTU"},
    {"column": "Rabat",           "xpath": "ns:RABAT"},
]

# --- Item XPath ---
# Path to repeating line-item elements, relative to DOKUMENT.
ITEM_XPATH = "ns:POZYCJE/ns:POZYCJA"

# --- Key fields ---
# ITEM_KEY: column that identifies a line item within a file (e.g. "LP")
ITEM_KEY = "LP"

# --- Update fields ---
# Which ITEM columns to write back to XML when running excel_to_xml.py.
# If empty, all ITEM_FIELDS present in the Excel will be updated.
# Example: ["TowarKod"] to only update product codes.
UPDATE_FIELDS = []
