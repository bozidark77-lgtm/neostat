"""
make_fixtures.py — generate small, valid sample workbooks for testing.

Writes two .xlsx files into tests/fixtures/ using the EXACT column headers the
app expects, named so convert.py routes them to the right reader:

  * "IZVESTAJ DOBAVLJACA TEST.xlsx"  -> supplier reader (_read_supplier_all_sheets)
  * "AMS Team TEST - BREZA.xlsx"     -> robust BREZA reader (header below a title row)

The data is synthetic but deliberately exercises every code path in analyze.py:
a clean match, LATE_IN, EARLY_OUT, MISSING_ON_BREZA, WRONG_CARD_ID and
OVERLAP_DIFFERENT_PLANTS. Replace these with your real CSV/XLSX files for a
production run — this is only here so CI can prove the pipeline end-to-end.
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook

FIX = Path(__file__).resolve().parent / "fixtures"
FIX.mkdir(parents=True, exist_ok=True)

# --- Supplier report (IZVEŠTAJ DOBAVLJAČA) -------------------------------------
# Dates are mm/dd/yyyy, times are hh:mm — exactly as the headers declare.
SUP_COLS = [
    "Broj kartice",
    "Ime i prezime",
    "Datum pocetka rada (format mm/dd/yyyy)",
    "Vreme pocetka rada (format hh:mm)",
    "Datum zavrsetka rada (format mm/dd/yyyy)",
    "Vreme zavrsetka rada (format hh:mm)",
    "Broj NZN-a",
]
supplier_rows = [
    # clean match: supplier == BREZA within tolerance
    ["12345", "Petar Petrović",  "03/01/2026", "07:00", "03/01/2026", "15:00", "NZN-001"],
    # LATE_IN: supplier entry 20 min after the BREZA entry
    ["22222", "Marko Marković",  "03/01/2026", "07:20", "03/01/2026", "15:00", "NZN-002"],
    # EARLY_OUT: BREZA exit 30 min after the supplier exit
    ["33333", "Jovan Jovanović", "03/01/2026", "07:00", "03/01/2026", "15:00", "NZN-003"],
    # MISSING_ON_BREZA: no BREZA event for this card/date at all
    ["99999", "Nikola Nikolić",  "03/01/2026", "08:00", "03/01/2026", "16:00", "NZN-004"],
    # WRONG_CARD_ID: same person/date exists in BREZA under a different card
    ["55555", "Ana Anić",        "03/01/2026", "09:00", "03/01/2026", "17:00", "NZN-005"],
]

# --- BREZA gate events (AMS Team … BREZA) -------------------------------------
# "Datum i Vreme" is day-first (dd.mm.yyyy). "Smer" = ULAZ (in) / IZLAZ (out).
BREZA_COLS = ["Broj kartice", "Ime i prezime", "Kapija", "Smer", "Datum i Vreme"]
breza_rows = [
    ["12345", "Petar Petrović",  "P1-ULAZ",  "ULAZ",  "01.03.2026 07:00:00"],
    ["12345", "Petar Petrović",  "P1-IZLAZ", "IZLAZ", "01.03.2026 15:00:00"],
    ["22222", "Marko Marković",  "P1-ULAZ",  "ULAZ",  "01.03.2026 07:00:00"],
    ["22222", "Marko Marković",  "P1-IZLAZ", "IZLAZ", "01.03.2026 15:00:00"],
    ["33333", "Jovan Jovanović", "P1-ULAZ",  "ULAZ",  "01.03.2026 07:00:00"],
    ["33333", "Jovan Jovanović", "P1-IZLAZ", "IZLAZ", "01.03.2026 15:30:00"],
    # Ana under the "wrong" card 66666 (supplier had her as 55555)
    ["66666", "Ana Anić",        "P1-ULAZ",  "ULAZ",  "01.03.2026 09:00:00"],
    ["66666", "Ana Anić",        "P1-IZLAZ", "IZLAZ", "01.03.2026 17:00:00"],
    # Stefan logged at two different plants at overlapping times
    ["77777", "Stefan Stefanović", "P1-ULAZ",  "ULAZ",  "01.03.2026 10:00:00"],
    ["77777", "Stefan Stefanović", "P1-IZLAZ", "IZLAZ", "01.03.2026 12:00:00"],
    ["88888", "Stefan Stefanović", "P2-ULAZ",  "ULAZ",  "01.03.2026 11:00:00"],
    ["88888", "Stefan Stefanović", "P2-IZLAZ", "IZLAZ", "01.03.2026 13:00:00"],
]


def write_supplier(path: Path) -> None:
    """Plain single-sheet workbook, header on row 1 (as the supplier reader expects)."""
    pd.DataFrame(supplier_rows, columns=SUP_COLS).to_excel(path, index=False, engine="openpyxl")


def write_breza(path: Path) -> None:
    """BREZA export shape: a title row, THEN the header row — the robust reader scans for it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "BREZA"
    ws.append(["AMS Team — BREZA izveštaj (mart 2026)"])   # title row above the header
    ws.append(BREZA_COLS)
    for r in breza_rows:
        ws.append(r)
    wb.save(path)


def main() -> None:
    sup = FIX / "IZVESTAJ DOBAVLJACA TEST.xlsx"
    brz = FIX / "AMS Team TEST - BREZA.xlsx"
    write_supplier(sup)
    write_breza(brz)
    print("Wrote fixtures:")
    print("  " + str(sup))
    print("  " + str(brz))


if __name__ == "__main__":
    main()
