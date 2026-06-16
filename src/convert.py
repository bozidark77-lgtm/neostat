"""
convert.py  —  RECONSTRUCTED SOURCE (from PyInstaller .pyc, Python 3.13)

Reconstructed from the extracted bytecode: function signatures, constants,
column aliases and string literals are exact (read from the code objects).
The control flow inside each function is reconstructed from the sequence of
referenced names + standard pandas idioms and may differ in minor detail from
the original. For byte-exact source, run `convert.pyc` through a 3.13-aware
decompiler (see README).

Purpose: Convert supplier / BREZA .xlsx files into normalized .csv, with two
special readers: one for the multi-sheet "IZVEŠTAJ DOBAVLJAČA" workbook and one
robust reader for "AMS Team … BREZA" exports whose header row is not row 0.
"""

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd


def _norm_text(value) -> str:
    """Lowercase, strip, remove diacritics (NFKD), collapse whitespace."""
    s = str(value).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_col(value) -> str:
    return _norm_text(value)


def _find_column(df: pd.DataFrame, aliases: dict):
    """Return the real column in df whose normalized name matches any alias."""
    norm_to_real = {_norm_col(c): c for c in df.columns}
    for canonical, names in aliases.items():
        for n in names:
            if any(_norm_col(n) == k for k in norm_to_real):
                return norm_to_real[_norm_col(n)]
    return None


def _is_empty_first_row_value(value) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip().lower() == ""


def _is_protected_column(col_name) -> bool:
    """NZN columns are never dropped even if the first data row is empty."""
    protected = {_norm_col(x) for x in ("Broj NZN-a", "Broj NZN", "NZN")}
    return _norm_col(col_name) in protected


def _drop_columns_empty_in_first_row(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    first = df.iloc[0]
    keep = [
        c for c in df.columns
        if _is_protected_column(c) or not _is_empty_first_row_value(first[c])
    ]
    return df[keep].copy()


def _is_ams_team_file(input_path: Path) -> bool:
    return "ams team" in _norm_text(input_path.stem)


def _is_supplier_file(input_path: Path) -> bool:
    stem = _norm_text(input_path.stem)
    return ("izvestaj dobavljaca" in stem) or ("izvestaj" in stem)


# Supplier column aliases (header variants seen across exports, incl. typos)
_SUPPLIER_ALIASES = {
    "name":     ["Ime I prezime", "Ime i prezime"],
    "date_in":  ["Datum pocetka rada (format mm/dd/yyyy)"],
    "time_in":  ["Vreme pocetka rada (fomat hh:mm)", "Vreme pocetka rada (format hh:mm)"],
    "date_out": ["Datum zavrsetka rada (format mm/dd/yyyy)"],
    "time_out": ["Vreme zavrsetka rada (fomat hh:mm)", "Vreme zavrsetka rada (format hh:mm)"],
}


def _plant_from_sheet(sheet_name) -> str:
    """Derive the plant (pogon) from a supplier sheet name.

    The supplier workbook keeps one sheet per plant, e.g. 'Izvestaj Aglomeracija',
    'Izvestaj VP REMONT(2)'. Strip the leading 'Izvestaj' and any trailing '(2)'
    copy-marker so we get 'Aglomeracija' / 'VP REMONT'. This is the only place the
    plant is recorded, so it must be preserved into the CSV for analyze.py.
    """
    s = str(sheet_name).strip()
    s = re.sub(r"^izve[sš]taj\s*", "", s, flags=re.IGNORECASE)  # drop 'Izvestaj' prefix
    s = re.sub(r"\s*\(\d+\)\s*$", "", s)                        # drop trailing '(2)'
    return s.strip()


def _read_supplier_all_sheets(input_path: Path) -> pd.DataFrame:
    """Read every sheet of the supplier workbook, dynamically find the header row,
    and concatenate non-empty rows. Each row is tagged with its plant (from the sheet name).
    """
    sheets = pd.read_excel(input_path, sheet_name=None, engine="openpyxl", header=None)
    frames = []
    
    # Ključni izrazi koji signaliziraju stvarni red sa zaglavljem
    wanted_supplier_keywords = {"ime i prezime", "broj kartice", "datum pocetka rada", "vreme pocetka rada"}
    
    for _name, sdf_raw in sheets.items():
        if sdf_raw.empty:
            continue
            
        # 1. Dinamička detekcija reda zaglavlja (skeniramo prvih 15 redova)
        header_row = 0
        for i in range(min(15, len(sdf_raw))):
            cells = {_norm_text(v) for v in sdf_raw.iloc[i].tolist() if pd.notna(v)}
            # Ako red sadrži bar jednu ključnu reč specifičnu za izveštaj dobavljača
            if any(k in cells for k in wanted_supplier_keywords) or any(any(k in c for k in wanted_supplier_keywords) for c in cells):
                header_row = i
                break
                
        # 2. Ponovo formiramo DataFrame za taj list koristeći detektovani red kao kolone
        columns_labels = sdf_raw.iloc[header_row].values
        sdf = sdf_raw.iloc[header_row + 1:].copy()
        sdf.columns = columns_labels
        sdf.reset_index(drop=True, inplace=True)
        
        # 3. Čišćenje i filtriranje praznih redova
        sdf = sdf.dropna(how="all").copy()
        if not sdf.empty:
            sdf["Pogon"] = _plant_from_sheet(_name)
            frames.append(sdf)
            
    if not frames:
        return pd.DataFrame()
        
    df = pd.concat(frames, ignore_index=True, sort=False)
    
    # Ostatak originalne logike za čišćenje i validaciju polja
    for c in df.columns:
        if pd.api.types.is_object_dtype(df[c]):
            df[c] = df[c].astype(str).str.strip().replace(("", "nan", "None", "NaT"), pd.NA)
            
    subset = [_find_column(df, {k: v}) for k, v in _SUPPLIER_ALIASES.items()
              if _find_column(df, {k: v}) is not None]
    if subset:
        df = df.dropna(subset=subset, how="any")
    df = df.dropna(axis=1, how="all")
    return df.reset_index(drop=True)


# BREZA / AMS-Team column aliases
_AMS_ALIASES = {
    "card": ["broj kartice"],
    "name": ["ime i prezime"],
    "gate": ["kapija"],
    "dir":  ["smer"],
    "dt":   ["datum i vreme"],
}


def _read_ams_sheet_robust(input_path: Path, sheet) -> pd.DataFrame:
    """
    AMS Team BREZA exports often have title rows before the real header.
    Scan the first ~30 rows for the row that contains the expected headers,
    then re-read using that row as the header.
    """
    raw = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl", header=None)
    if raw.empty:
        return raw
    wanted = {_norm_text(n) for names in _AMS_ALIASES.values() for n in names}
    header_row = None
    for i in range(min(30, len(raw))):
        cells = {_norm_text(v) for v in raw.iloc[i].tolist()}
        # need at least the card + datetime + direction-ish columns present
        if any(w in cells for w in wanted) and len(cells & wanted) >= 2:
            header_row = i
            break
    if header_row is None:
        header_row = 0
    df = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl", skiprows=header_row)
    df = df.dropna(how="all")
    df = df.fillna("")
    df = df.astype(str)
    # drop helper / unnamed columns
    df = df[[c for c in df.columns if not _norm_text(c).startswith("unnamed")]].copy()
    return df.reset_index(drop=True)


def convert_xlsx_to_csv(input_path, output_path, sheet=0, sep=","):
    """Dispatch to the correct reader by filename, then write a UTF-8-SIG CSV."""
    input_path = Path(input_path)
    output_path = Path(output_path)

    if _is_supplier_file(input_path):
        df = _read_supplier_all_sheets(input_path)
    elif _is_ams_team_file(input_path):
        df = _read_ams_sheet_robust(input_path, sheet)
    else:
        df = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl")
        df = _drop_columns_empty_in_first_row(df)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8-sig", sep=sep)
    return output_path


def main():
    ap = argparse.ArgumentParser(description="Convert XLSX to CSV.")
    ap.add_argument("input", help="Path to .xlsx file or folder containing .xlsx files")
    ap.add_argument("-o", "--output", help="Output file (for single input) or output folder")
    ap.add_argument("-s", "--sheet", default="0", help="Sheet name or index (default: 0)")
    ap.add_argument("--sep", default=",", help="CSV separator (default: ','). Example: ';'")
    args = ap.parse_args()

    inp = Path(args.input)
    if not inp.exists():
        raise FileNotFoundError("Input not found: " + str(inp))

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    if inp.is_file():
        if inp.suffix.lower() != ".xlsx":
            raise ValueError("Input file must be .xlsx")
        out = Path(args.output) if args.output else inp.with_suffix(".csv")
        convert_xlsx_to_csv(inp, out, sheet=sheet, sep=args.sep)
        print("Converted: " + str(inp) + " -> " + str(out))
    else:
        files = sorted(inp.glob("*.xlsx"))
        if not files:
            print("No .xlsx files found.")
            return
        outdir = Path(args.output) if args.output else inp
        outdir.mkdir(parents=True, exist_ok=True)
        for f in files:
            out = outdir / (f.stem + ".csv")
            convert_xlsx_to_csv(f, out, sheet=sheet, sep=args.sep)
            print("Converted: " + f.name + " -> " + out.name)


if __name__ == "__main__":
    main()
