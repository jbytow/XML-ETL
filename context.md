# Project Context

## Goal

DRO-MASZ (manufacturer of plate compactors, vibrating rammers, and power trowels) needs a tool for bulk product code updates across ~6000 XML invoices exported from Comarch Optima. The current coding system is inconsistent - the same product type has different prefixes (ZAG, DRB, DR-, CZĘ, P0, CNP, bare numbers). The target format is a unified `CATEGORY/MODEL/SUBTYPE` convention.

## Data Flow

```
                    ┌─────────────────────┐
                    │  Comarch Optima      │
                    │  (XML export)        │
                    └────────┬────────────┘
                             │
                    ┌────────▼────────────┐
                    │  input/xml/          │
                    │  XML invoices        │
                    │  (e.g. FS_1_2025.xml)│
                    └────────┬────────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
     ┌────────▼────────┐    │    ┌─────────▼──────────┐
     │ xml_to_excel.py  │    │    │ excel_to_xml.py     │
     │                  │    │    │                      │
     │ Parses XML,      │    │    │ Reads product catalog│
     │ exports to a     │    │    │ replaces codes in    │
     │ single Excel     │    │    │ XML per mapping      │
     └────────┬────────┘    │    └─────────┬──────────┘
              │              │              │
              │    ┌─────────▼──────────┐   │
              │    │ input/kartoteka     │   │
              │    │ z kluczem.xlsx      │───┘
              │    │                     │
              │    │ Columns:            │
              │    │ - Kod (current)     │
              │    │ - Kod Zaktualizowany│
              │    └────────────────────┘
              │
     ┌────────▼──────────────────────────┐
     │  output/                           │
     │  ├── xml_export_*.xlsx  (report)   │
     │  └── xml/               (new XMLs) │
     └───────────────────────────────────┘
```

## Scripts

### xml_to_excel.py
- Scans `input/xml/*.xml`
- Parses each file with lxml (namespace: `http://www.cdn.com.pl/optima/dokument`)
- Extracts header fields (once per invoice) + item fields (once per POZYCJA)
- Each POZYCJA = 1 Excel row; header data repeated
- Joins with product catalog: appends **KodZaktualizowany** column (lookup by TowarKod -> Kod)
- Saves to `output/xml_export_YYYYMMDD_HHMMSS.xlsx`

### excel_to_xml.py
- Reads product catalog (`Kod` -> `Kod Zaktualizowany`)
- Scans all XML files in `input/xml/`
- For each invoice: checks every POZYCJA's `TOWAR/KOD` against the mapping
- If matched, replaces `TOWAR/KOD` and `TOWAR/NUMER_KATALOGOWY` with the new code
- Saves modified XML to `output/xml/` (originals untouched)

### config.py
- Paths: `XML_DIR`, `OUTPUT_DIR`, `KARTOTEKA_PATH`
- Comarch Optima XML namespace
- `HEADER_FIELDS` - XPath -> Excel column mapping for invoice headers
- `ITEM_FIELDS` - XPath -> Excel column mapping for line items
- `ITEM_XPATH` - path to POZYCJA elements

## Product Catalog

File `input/kartoteka z kluczem.xlsx` - export from Comarch Optima, 706 products (type TP = goods). Key columns:

| Column | Description |
|---|---|
| Kod | Current product code in the system |
| Nazwa | Full product name |
| Typ | TP (goods) or UP (service) |
| Nr katalogowy | Catalog number (often = Kod, but not always) |
| Kod Zaktualizowany | New code to assign |

## Product Categories

| Category | Prefix | Count | Code format | Example |
|---|---|---|---|---|
| Plate compactors (Zagęszczarki) | ZAG | ~37 | ZAG/MODEL/ENGINE | ZAG/DRB120/HGX160 |
| Vibrating rammers (Stopy wibracyjne) | STO | ~6 | STO/MODEL/ENGINE | STO/DRB72FW/HGX100 |
| Power trowels (Zacieraczki) | ZAC | ~8-31 | ZAC/MODEL/ENGINE | ZAC/DRB760/LG200F |
| Engines (Silniki) | SIL | ~11-14 | SIL/BRAND/MODEL | SIL/LONCIN/G200F |
| Parts (Części) | CZE | ~600+ | CZE/MACHINE/DESC | CZE/DRB90C/DEKIEL BOCZNY |

Naming convention proposals generated in `output/propozycja_nazewnictwa_v2.xlsx`.

## Engines Used in Machines

- **Loncin**: G200F, G270F, G390F, G420F, 168F-2H, 165F-3H, LD178F, LG200F, LG390F
- **Honda**: GX160, GX200, GX270, GX340, GX390
- **Hyundai (HGX)**: HGX100, HGX160, HGX270, HGX390
- **Lifan**: GX390 (188F)

## Dependencies

- `lxml` - XML parsing (fast, supports XPath with namespaces)
- `openpyxl` - Excel read/write (.xlsx)
- `tqdm` - progress bar

---

# Kontekst projektu (PL)

## Cel

DRO-MASZ (producent zagęszczarek, stóp wibracyjnych, zacieraczek) potrzebuje narzędzia do masowej aktualizacji kodów produktów w ~6000 fakturach XML wyeksportowanych z Comarch Optima. Aktualny system kodów jest niespójny - ten sam typ produktu ma różne prefiksy (ZAG, DRB, DR-, CZĘ, P0, CNP, numery bez prefiksu). Docelowo kody mają mieć ujednolicony format `KATEGORIA/MODEL/PODTYP`.

## Przepływ danych

```
                    ┌─────────────────────┐
                    │  Comarch Optima      │
                    │  (eksport XML)       │
                    └────────┬────────────┘
                             │
                    ┌────────▼────────────┐
                    │  input/xml/          │
                    │  Faktury XML         │
                    │  (np. FS_1_2025.xml) │
                    └────────┬────────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
     ┌────────▼────────┐    │    ┌─────────▼──────────┐
     │ xml_to_excel.py  │    │    │ excel_to_xml.py     │
     │                  │    │    │                      │
     │ Parsuje XML,     │    │    │ Wczytuje kartotekę,  │
     │ eksportuje do    │    │    │ podmienia kody w XML │
     │ jednego Excela   │    │    │ wg mapowania         │
     └────────┬────────┘    │    └─────────┬──────────┘
              │              │              │
              │    ┌─────────▼──────────┐   │
              │    │ input/kartoteka     │   │
              │    │ z kluczem.xlsx      │───┘
              │    │                     │
              │    │ Kolumny:            │
              │    │ - Kod (obecny)      │
              │    │ - Kod Zaktualizowany│
              │    └────────────────────┘
              │
     ┌────────▼──────────────────────────┐
     │  output/                           │
     │  ├── xml_export_*.xlsx  (raport)   │
     │  └── xml/               (nowe XML) │
     └───────────────────────────────────┘
```

## Skrypty

### xml_to_excel.py
- Skanuje `input/xml/*.xml`
- Parsuje każdy plik lxml (namespace: `http://www.cdn.com.pl/optima/dokument`)
- Wyciąga pola nagłówkowe (raz na fakturę) + pola pozycji (raz na POZYCJA)
- Każda pozycja = 1 wiersz; dane nagłówkowe powtórzone
- Łączy z kartoteką: dopisuje kolumnę **KodZaktualizowany** (lookup po TowarKod -> Kod w kartotece)
- Zapisuje do `output/xml_export_YYYYMMDD_HHMMSS.xlsx`

### excel_to_xml.py
- Wczytuje kartotekę (`Kod` -> `Kod Zaktualizowany`)
- Skanuje wszystkie XML w `input/xml/`
- W każdej fakturze: dla każdej POZYCJA sprawdza czy `TOWAR/KOD` jest w mapowaniu
- Jeśli tak, podmienia `TOWAR/KOD` i `TOWAR/NUMER_KATALOGOWY` na nowy kod
- Zapisuje zmodyfikowany XML do `output/xml/` (oryginały nienaruszone)

### config.py
- Ścieżki: `XML_DIR`, `OUTPUT_DIR`, `KARTOTEKA_PATH`
- Namespace XML Comarch Optima
- `HEADER_FIELDS` - mapowanie XPath -> kolumna Excel dla nagłówka faktury
- `ITEM_FIELDS` - mapowanie XPath -> kolumna Excel dla pozycji
- `ITEM_XPATH` - ścieżka do elementów POZYCJA

## Kartoteka produktów

Plik `input/kartoteka z kluczem.xlsx` - eksport z Comarch Optima, 706 produktów (typ TP = towar). Kluczowe kolumny:

| Kolumna | Opis |
|---|---|
| Kod | Obecny kod produktu w systemie |
| Nazwa | Pełna nazwa produktu |
| Typ | TP (towar) lub UP (usługa) |
| Nr katalogowy | Numer katalogowy (często = Kod, ale nie zawsze) |
| Kod Zaktualizowany | Nowy kod do nadania |

## Kategorie produktów

| Kategoria | Prefiks | Ilość | Format kodu | Przykład |
|---|---|---|---|---|
| Zagęszczarki gruntu | ZAG | ~37 | ZAG/MODEL/SILNIK | ZAG/DRB120/HGX160 |
| Stopy wibracyjne | STO | ~6 | STO/MODEL/SILNIK | STO/DRB72FW/HGX100 |
| Zacieraczki | ZAC | ~8-31 | ZAC/MODEL/SILNIK | ZAC/DRB760/LG200F |
| Silniki | SIL | ~11-14 | SIL/MARKA/MODEL | SIL/LONCIN/G200F |
| Części | CZE | ~600+ | CZE/MASZYNA/OPIS | CZE/DRB90C/DEKIEL BOCZNY |

Propozycje nazewnictwa wygenerowane w `output/propozycja_nazewnictwa_v2.xlsx`.

## Silniki stosowane w maszynach

- **Loncin**: G200F, G270F, G390F, G420F, 168F-2H, 165F-3H, LD178F, LG200F, LG390F
- **Honda**: GX160, GX200, GX270, GX340, GX390
- **Hyundai (HGX)**: HGX100, HGX160, HGX270, HGX390
- **Lifan**: GX390 (188F)

## Zależności

- `lxml` - parsowanie XML (szybkie, wspiera XPath z namespace)
- `openpyxl` - odczyt/zapis Excel (.xlsx)
- `tqdm` - pasek postępu
