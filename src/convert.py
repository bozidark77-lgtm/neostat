"""
convert.py  —  Universal Multi-Format Data Normalization Engine (Neostat™).

Handles conversion and dynamic header alignment for Excel (.xlsx, .xls) and 
textual (.csv, .txt) files. Automatically resolves duplicate and empty column 
names to prevent Pandas reindexing errors.
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
    stem = _norm_text(input_path.stem)
    return ("ams team" in stem) or ("breza" in stem)


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
    s = str(sheet_name).strip()
    s = re.sub(r"^izve[sš]taj\s*", "", s, flags=re.IGNORECASE)  # drop 'Izvestaj' prefix
    s = re.sub(r"\s*\(\d+\)\s*$", "", s)                        # drop trailing '(2)'
    return s.strip()


def _make_headers_unique(raw_labels) -> list:
    """Helper to eliminate empty or duplicated column names dynamically."""
    clean_labels = []
    seen = set()
    for idx, col in enumerate(raw_labels):
        col_str = str(col).strip() if pd.notna(col) else ""
        
        # Ako je naziv kolone prazan, dajemo unikatno ime
        if col_str == "" or col_str.lower().startswith("unnamed"):
            col_str = f"PraznaKolona_{idx}"
            
        # Razrešavanje duplikata (npr. ako postoje dva 'Datum zavrsetka rada')
        original_col = col_str
        counter = 1
        while col_str in seen:
            col_str = f"{original_col}.{counter}"
            counter += 1
            
        seen.add(col_str)
        clean_labels.append(col_str)
    return clean_labels


def _read_supplier_all_sheets(input_path: Path) -> pd.DataFrame:
    """Read every sheet of the supplier workbook, dynamically find the header row,
    ensure unique columns, and concatenate non-empty rows.
    """
    sheets = pd.read_excel(input_path, sheet_name=None, engine="openpyxl", header=None)
    frames = []
    
    wanted_supplier_keywords = {"ime", "kartice", "kartica", "pocetka", "početka", "zavrsetka", "završetka", "nzn"}
    
    for _name, sdf_raw in sheets.items():
        if sdf_raw.empty:
            continue
            
        header_row = 0
        for i in range(min(15, len(sdf_raw))):
            cells = {str(v).strip().lower() for v in sdf_raw.iloc[i].tolist() if pd.notna(v)}
            if any(any(kw in cell for kw in wanted_supplier_keywords) for cell in cells):
                header_row = i
                break
                
        raw_labels = sdf_raw.iloc[header_row].values
        clean_labels = _make_headers_unique(raw_labels)
        
        sdf = sdf_raw.iloc[header_row + 1:].copy()
        sdf.columns = clean_labels
        sdf.reset_index(drop=True, inplace=True)
        
        sdf = sdf.dropna(how="all").copy()
        if not sdf.empty:
            sdf["Pogon"] = _plant_from_sheet(_name)
            frames.append(sdf)
            
    if not frames:
        return pd.DataFrame()
        
    df = pd.concat(frames, ignore_index=True, sort=False)
    
    for c in df.columns:
        if pd.api.types.is_object_dtype(df[c]):
            df[c] = df[c].astype(str).str.strip().replace(("", "nan", "None", "NaT"), pd.NA)
            
    subset = [_find_column(df, {k: v}) for k, v in _SUPPLIER_ALIASES.items()
              if _find_column(df, {k: v}) is not None]
    if subset:
        df = df.dropna(subset=subset, how="any")
    df = df.dropna(axis=1, how="all")
    return df.reset_index(drop=True)


_AMS_ALIASES = {
    "card": ["broj kartice"],
    "name": ["ime i prezime"],
    "gate": ["kapija"],
    "dir":  ["smer"],
    "dt":   ["datum i vreme"],
}


def _read_ams_sheet_robust(input_path: Path, sheet) -> pd.DataFrame:
    raw = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl", header=None)
    if raw.empty:
        return raw
    wanted = {_norm_text(n) for names in _AMS_ALIASES.values() for n in names}
    header_row = None
    for i in range(min(30, len(raw))):
        cells = {_norm_text(v) for v in raw.iloc[i].tolist() if pd.notna(v)}
        if any(w in cells for w in wanted) and len(cells & wanted) >= 2:
            header_row = i
            break
    if header_row is None:
        header_row = 0
        
    raw_labels = raw.iloc[header_row].values
    clean_labels = _make_headers_unique(raw_labels)
    
    df = raw.iloc[header_row + 1:].copy()
    df.columns = clean_labels
    df.reset_index(drop=True, inplace=True)
    
    df = df.dropna(how="all")
    df = df.fillna("")
    df = df.astype(str)
    df = df[[c for c in df.columns if not _norm_text(c).startswith("praznakolona")]].copy()
    return df.reset_index(drop=True)


def convert_xlsx_to_csv(input_path, output_path, sheet=0, sep=","):
    """Universal multi-format data parser and normalizer."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    ext = input_path.suffix.lower()

    if ext in [".csv", ".txt"]:
        df_raw = pd.read_csv(str(input_path), sep=None, engine="python", encoding="utf-8-sig", dtype=str, header=None)
        if df_raw.empty:
            df = df_raw
        else:
            header_row = 0
            wanted_supplier_keywords = {"ime", "kartice", "pocetka", "početka", "zavrsetka", "završetka", "nzn"}
            
            for i in range(min(15, len(df_raw))):
                cells = {str(v).strip().lower() for v in df_raw.iloc[i].tolist() if pd.notna(v)}
                if any(any(kw in cell for kw in wanted_supplier_keywords) for cell in cells):
                    header_row = i
                    break
            
            raw_labels = df_raw.iloc[header_row].values
            clean_labels = _make_headers_unique(raw_labels)
            
            df_cleaned = df_raw.iloc[header_row + 1:].copy()
            df_cleaned.columns = clean_labels
            df_cleaned = df_cleaned.dropna(how="all").copy()
            
            if _is_supplier_file(input_path):
                df_cleaned["Pogon"] = _plant_from_sheet(input_path.stem)
            df = df_cleaned.reset_index(drop=True)

    elif ext in [".xlsx", ".xls"]:
        if _is_supplier_file(input_path):
            df = _read_supplier_all_sheets(input_path)
        elif _is_ams_team_file(input_path):
            df = _read_ams_sheet_robust(input_path, sheet)
        else:
            df = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl")
            df = _drop_columns_empty_in_first_row(df)
    else:
        raise ValueError(f"Format fajla '{ext}' nije podržan unutar Neostat sistema.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8-sig", sep=sep)
    return output_path


def main():
    ap = argparse.ArgumentParser(description="Convert XLSX to CSV.")
    ap.add_argument("input", help="Path to .xlsx file or folder containing .xlsx files")
    ap.add_argument("-o", "--output", help="Output file (for single input) or output folder")
    ap.add_argument("-s", "--sheet", default="0", help="Sheet name or index (default: 0)")
    ap.add_argument("--sep", default=",", help="CSV separator (default: ',')")
    args = ap.parse_args()

    inp = Path(args.input)
    if not inp.exists():
        raise FileNotFoundError("Input not found: " + str(inp))

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    if inp.is_file():
        allowed = {".xlsx", ".xls", ".csv", ".txt"}
        if inp.suffix.lower() not in allowed:
            raise ValueError("Input file must be .xlsx, .xls, .csv, or .txt")
        out = Path(args.output) if args.output else inp.with_suffix(".csv")
        convert_xlsx_to_csv(inp, out, sheet=sheet, sep=args.sep)
        print("Converted: " + str(inp) + " -> " + str(out))
    else:
        files = sorted([f for f in inp.glob("*") if f.suffix.lower() in [".xlsx", ".xls", ".csv", ".txt"]])
        if not files:
            print("No supported files found.")
            return
        outdir = Path(args.output) if args.output else inp
        outdir.mkdir(parents=True, exist_ok=True)
        for f in files:
            out = outdir / (f.stem + ".csv")
            convert_xlsx_to_csv(f, out, sheet=sheet, sep=args.sep)
            print("Converted: " + f.name + " -> " + out.name)


if __name__ == "__main__":
    main()
