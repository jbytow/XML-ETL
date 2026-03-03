# XML ETL - DRO-MASZ

ETL tool for processing Comarch Optima XML invoices. Exports invoice data from XML files to Excel and performs bulk product code updates across XML files based on a product mapping spreadsheet.

## Requirements

- Python 3.10+
- Dependencies: `pip install -r requirements.txt`

## Project Structure

```
XML ETL/
├── input/
│   ├── xml/                        # Comarch Optima XML invoices
│   └── kartoteka z kluczem.xlsx    # Product catalog with code mappings
├── output/
│   ├── xml/                        # Modified XML invoices (after code update)
│   ├── xml_export_*.xlsx           # Invoice data exported to Excel
│   └── propozycja_nazewnictwa*.xlsx # Naming convention proposals
├── config.py             # Configuration: paths, namespace, XPath field mappings
├── xml_to_excel.py       # Script 1: XML -> Excel
├── excel_to_xml.py       # Script 2: Product catalog -> updated XML files
└── requirements.txt      # openpyxl, lxml, tqdm
```

## Usage

### 1. Export XML -> Excel

```
python xml_to_excel.py
```

Parses all `.xml` files from `input/xml/`, extracts header data (invoice number, date, buyer, totals) and line items (product code, name, quantity, prices). Each line item = 1 row in Excel.

If the product catalog (`input/kartoteka z kluczem.xlsx`) is available, appends a **KodZaktualizowany** column with the Kod -> Kod Zaktualizowany mapping.

Output: `output/xml_export_YYYYMMDD_HHMMSS.xlsx`

### 2. Update Product Codes in XML

```
python excel_to_xml.py "input\kartoteka z kluczem.xlsx"
```

Reads the product catalog (columns **Kod** and **Kod Zaktualizowany**), scans all XML invoices, and replaces product codes (`TOWAR/KOD` and `TOWAR/NUMER_KATALOGOWY`) where a match is found.

- Original files in `input/xml/` are not modified
- Modified invoices are saved to `output/xml/` (each invoice as a separate file)
- Skips rows where Kod Zaktualizowany is empty or identical to Kod

## Configuration (`config.py`)

Pre-configured for Comarch Optima invoices (namespace `http://www.cdn.com.pl/optima/dokument`). Key settings:

- `HEADER_FIELDS` - invoice header fields (number, date, buyer, totals)
- `ITEM_FIELDS` - line item fields (product code, name, quantity, prices)
- `KARTOTEKA_PATH` - path to the product catalog

To add/remove columns from the export, edit the `HEADER_FIELDS` and `ITEM_FIELDS` lists.

## XML Format (Comarch Optima)

```
ROOT (xmlns="http://www.cdn.com.pl/optima/dokument")
└── DOKUMENT
    ├── NAGLOWEK          # Invoice header
    │   ├── NUMER_PELNY, DATA_DOKUMENTU, ...
    │   ├── PLATNIK / ODBIORCA / SPRZEDAWCA
    │   ├── PLATNOSC, WALUTA
    │   └── KWOTY         # Net/gross/VAT totals
    └── POZYCJE
        └── POZYCJA        # Repeats per line item
            ├── TOWAR      # KOD, NAZWA, EAN, NUMER_KATALOGOWY
            ├── CENY, STAWKA_VAT
            └── ILOSC, WARTOSC_NETTO, WARTOSC_BRUTTO
```

---

## Instrukcja (PL)

Narzędzie ETL do przetwarzania faktur XML z Comarch Optima. Umożliwia eksport danych z faktur XML do Excela oraz masową aktualizację kodów produktów w plikach XML na podstawie kartoteki.

### Wymagania

- Python 3.10+
- Zależności: `pip install -r requirements.txt`

### Struktura projektu

```
XML ETL/
├── input/
│   ├── xml/                        # Faktury XML z Comarch Optima
│   └── kartoteka z kluczem.xlsx    # Kartoteka produktów z mapowaniem kodów
├── output/
│   ├── xml/                        # Zmodyfikowane faktury XML (po aktualizacji kodów)
│   ├── xml_export_*.xlsx           # Eksport faktur do Excela
│   └── propozycja_nazewnictwa*.xlsx # Propozycje nowego nazewnictwa
├── config.py             # Konfiguracja: ścieżki, namespace, mapowania pól XPath
├── xml_to_excel.py       # Skrypt 1: XML -> Excel
├── excel_to_xml.py       # Skrypt 2: Kartoteka -> zaktualizowane XML
└── requirements.txt      # openpyxl, lxml, tqdm
```

### 1. Eksport XML -> Excel

```
python xml_to_excel.py
```

Parsuje wszystkie pliki `.xml` z `input/xml/`, wyciąga dane nagłówkowe (numer faktury, data, płatnik, kwoty) oraz pozycje (kod towaru, nazwa, ilość, ceny). Każda pozycja = 1 wiersz w Excelu.

Jeśli kartoteka (`input/kartoteka z kluczem.xlsx`) jest dostępna, dołącza kolumnę **KodZaktualizowany** z mapowaniem Kod -> Kod Zaktualizowany.

Wynik: `output/xml_export_YYYYMMDD_HHMMSS.xlsx`

### 2. Aktualizacja kodów produktów w XML

```
python excel_to_xml.py "input\kartoteka z kluczem.xlsx"
```

Wczytuje kartotekę produktów (kolumny **Kod** i **Kod Zaktualizowany**), skanuje wszystkie faktury XML i podmienia kody produktów (`TOWAR/KOD` i `TOWAR/NUMER_KATALOGOWY`) tam gdzie znajdzie dopasowanie.

- Oryginalne pliki w `input/xml/` nie są modyfikowane
- Zmodyfikowane faktury trafiają do `output/xml/` (każda faktura jako osobny plik)
- Pomija wiersze gdzie Kod Zaktualizowany jest pusty lub identyczny z Kod

### Konfiguracja (`config.py`)

Plik jest wstępnie skonfigurowany pod faktury z Comarch Optima (namespace `http://www.cdn.com.pl/optima/dokument`). Kluczowe ustawienia:

- `HEADER_FIELDS` - pola z nagłówka faktury (numer, data, płatnik, kwoty)
- `ITEM_FIELDS` - pola z pozycji faktury (kod towaru, nazwa, ilość, ceny)
- `KARTOTEKA_PATH` - ścieżka do kartoteki produktów

Aby dodać/usunąć kolumny z eksportu, edytuj listy `HEADER_FIELDS` i `ITEM_FIELDS`.

### Format XML (Comarch Optima)

```
ROOT (xmlns="http://www.cdn.com.pl/optima/dokument")
└── DOKUMENT
    ├── NAGLOWEK          # Nagłówek faktury
    │   ├── NUMER_PELNY, DATA_DOKUMENTU, ...
    │   ├── PLATNIK / ODBIORCA / SPRZEDAWCA
    │   ├── PLATNOSC, WALUTA
    │   └── KWOTY         # Razem netto/brutto/VAT
    └── POZYCJE
        └── POZYCJA        # Powtarza się per pozycja
            ├── TOWAR      # KOD, NAZWA, EAN, NUMER_KATALOGOWY
            ├── CENY, STAWKA_VAT
            └── ILOSC, WARTOSC_NETTO, WARTOSC_BRUTTO
```
