"""
analyze.py  —  RECONSTRUCTED SOURCE (from PyInstaller .pyc, Python 3.13)

Reconstructed from extracted bytecode. Signatures, column aliases, issue-type
codes, Serbian messages, sheet names and the export column maps are EXACT.
Per-function control flow is reconstructed from the referenced-name sequence and
pandas idioms; verify against a 3.13 decompile of `analyze.pyc` for exact logic.

Purpose: Compare a supplier report (IZVEŠTAJ DOBAVLJAČA) against BREZA gate
access events. Build IN/OUT intervals from BREZA, match them to supplier rows
within a time tolerance, flag irregularities (late in / early out / missing /
wrong card / cross-plant overlap), and write a styled multi-sheet .xlsx.
"""

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill

DEFAULT_SUPPLIER_CSV = "csv_output/IZVESTAJ DOBAVLJACA 03.2026.csv"
DEFAULT_BREZA_CSV = "csv_output/AMS Team MART 2026 - BREZA.csv"


def _norm_col(s) -> str:
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def _read_csv_auto(path) -> pd.DataFrame:
    # python engine + sep=None sniffs the delimiter; utf-8-sig handles BOM
    return pd.read_csv(str(path), sep=None, engine="python", encoding="utf-8-sig", dtype=str)


def _resolve_columns(df: pd.DataFrame, aliases: dict) -> dict:
    norm_to_real = {_norm_col(c): c for c in df.columns}
    resolved = {}
    for canonical, names in aliases.items():
        hit = None
        for n in names:
            if _norm_col(n) in norm_to_real:
                hit = norm_to_real[_norm_col(n)]
                break
        if hit is None:
            raise ValueError("Nedostaje kolona za '" + canonical + "'. Pronađene kolone: " + str(list(df.columns)))
        resolved[canonical] = hit
    return resolved


def _parse_datetime_robust(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    out = pd.to_datetime(s, errors="coerce", format="mixed")
    if out.isna().any():
        alt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        out = out.fillna(alt)
    return out


def norm_text(s) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def norm_card(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):   # Excel turns card numbers into floats
        s = s[:-2]
    return s


def parse_plant_from_gate(gate) -> str:
    """First alpha(+digits) token of the gate code, e.g. 'P1-ULAZ' -> 'P1'."""
    s = str(gate).strip().upper()
    m = re.match(r"^([A-ZŠĐŽČĆ]+[0-9]*)", s)
    return m.group(1) if m else s[:20]


def parse_supplier(path) -> pd.DataFrame:
    aliases = {
        "card":     ["Broj kartice"],
        "name":     ["Ime I prezime", "Ime i prezime"],
        "date_in":  ["Datum pocetka rada (format mm/dd/yyyy)"],
        "time_in":  ["Vreme pocetka rada (fomat hh:mm)", "Vreme pocetka rada (format hh:mm)"],
        "date_out": ["Datum zavrsetka rada (format mm/dd/yyyy)"],
        "time_out": ["Vreme zavrsetka rada (fomat hh:mm)", "Vreme zavrsetka rada (format hh:mm)"],
        "nzn":      ["Broj NZN-a"],
    }
    df = _read_csv_auto(path)
    cols = _resolve_columns(df, aliases)
    df = df.dropna(subset=[cols["name"], cols["date_in"], cols["time_in"],
                           cols["date_out"], cols["time_out"]], how="any").copy()

    out = pd.DataFrame()
    out["employee_name"] = df[cols["name"]].astype(str).str.strip()
    out["name_key"] = out["employee_name"].map(norm_text)
    out["card_id"] = df[cols["card"]].map(norm_card)
    out["nzn"] = df[cols["nzn"]].astype(str).str.strip() if cols.get("nzn") else ""
    out["in_time"] = _parse_datetime_robust(df[cols["date_in"]].astype(str).str.strip()
                                            + " " + df[cols["time_in"]].astype(str).str.strip())
    out["out_time"] = _parse_datetime_robust(df[cols["date_out"]].astype(str).str.strip()
                                             + " " + df[cols["time_out"]].astype(str).str.strip())
    out = out.dropna(subset=["in_time", "out_time"])
    # If out < in, assume the shift crosses midnight -> add a day
    out.loc[out["out_time"] < out["in_time"], "out_time"] += pd.Timedelta(days=1)
    out = out.reset_index(drop=True)
    out["row_id"] = range(len(out))
    return out[["row_id", "employee_name", "name_key", "card_id", "nzn", "in_time", "out_time"]]


def parse_breza_events(path) -> pd.DataFrame:
    aliases = {
        "card": ["Broj kartice"],
        "name": ["Ime i Prezime", "Ime i prezime"],
        "gate": ["Kapija"],
        "dir":  ["Smer"],
        "dt":   ["Datum i Vreme"],
    }
    df = _read_csv_auto(path)
    cols = _resolve_columns(df, aliases)
    df = df.dropna(subset=[cols["name"], cols["dir"], cols["dt"]], how="any").copy()

    out = pd.DataFrame()
    out["employee_name"] = df[cols["name"]].astype(str).str.strip()
    out["name_key"] = out["employee_name"].map(norm_text)
    out["card_id"] = df[cols["card"]].map(norm_card)
    out["gate"] = df[cols["gate"]].astype(str).str.strip()
    out["plant"] = out["gate"].map(parse_plant_from_gate)

    def _dir(s):
        u = str(s).strip().upper()
        if "ULAZ" in u:
            return "IN"
        if "IZLAZ" in u:
            return "OUT"
        return "OTHER"
    out["direction"] = df[cols["dir"]].map(_dir)
    out["event_time"] = pd.to_datetime(df[cols["dt"]], errors="coerce", dayfirst=True)
    out = out.dropna(subset=["event_time"])
    out = out[out["direction"].isin(["IN", "OUT"])]
    out = out.drop_duplicates(subset=["name_key", "card_id", "direction", "event_time", "gate"])
    out = out.sort_values("event_time").reset_index(drop=True)
    out["event_date"] = out["event_time"].dt.date
    out["event_row_id"] = range(len(out))
    return out[["event_row_id", "employee_name", "name_key", "card_id",
                "gate", "plant", "direction", "event_time", "event_date"]]


def parse_breza(path) -> pd.DataFrame:
    """Pair consecutive IN/OUT events per card into work intervals."""
    ev = parse_breza_events(path)
    rows = []
    for card_id, g in ev.groupby("card_id", sort=True):
        g = g.sort_values("event_time")
        pending_in = None
        for _, e in g.iterrows():
            if e["direction"] == "IN":
                pending_in = e
            elif e["direction"] == "OUT" and pending_in is not None:
                rows.append({
                    "employee_name": pending_in["employee_name"],
                    "name_key": pending_in["name_key"],
                    "card_id": card_id,
                    "in_time": pending_in["event_time"],
                    "out_time": e["event_time"],
                    "in_gate": pending_in["gate"],
                    "out_gate": e["gate"],
                    "plant_in": pending_in["plant"],
                    "plant_out": e["plant"],
                })
                pending_in = None
    cols = ["employee_name", "name_key", "card_id", "in_time", "out_time",
            "in_gate", "out_gate", "plant_in", "plant_out"]
    out = pd.DataFrame(rows, columns=cols)
    if out.empty:
        return out
    out = out.reset_index(drop=True)
    out["row_id"] = range(len(out))
    return out[["row_id"] + cols]


def match_supplier_to_breza(sup: pd.DataFrame, brz: pd.DataFrame, tol_min=5) -> pd.DataFrame:
    """For each supplier row find the best BREZA interval (same card + date),
    score by |in_delta| + |out_delta|; emit per-row irregularities."""
    tol = pd.Timedelta(minutes=tol_min)
    matched_event_ids = set()
    out_rows = []

    for _, s in sup.iterrows():
        day = s["in_time"].date()
        same_card = brz[(brz["card_id"] == s["card_id"])
                        & (brz["in_time"].dt.date == day)]

        if same_card.empty:
            # Same person + date present under a DIFFERENT card?
            same_name = brz[(brz["name_key"] == s["name_key"])
                            & (brz["in_time"].dt.date == day)]
            if not same_name.empty:
                out_rows.append({
                    "issue_type": "WRONG_CARD_ID",
                    "employee_name": s["employee_name"],
                    "supplier_card": s["card_id"],
                    "breza_card": same_name.iloc[0]["card_id"],
                    "supplier_in": s["in_time"],
                    "supplier_out": s["out_time"],
                    "details": "Isti radnik i datum postoje u BREZA, ali sa drugim ID kartice.",
                })
            else:
                out_rows.append({
                    "issue_type": "MISSING_ON_BREZA",
                    "employee_name": s["employee_name"],
                    "supplier_card": s["card_id"],
                    "supplier_in": s["in_time"],
                    "supplier_out": s["out_time"],
                    "details": "Nema događaja u BREZA za isti ID kartice i isti datum.",
                })
            continue

        # pick best interval by combined timing delta
        best = None
        best_score = None
        for _, b in same_card.iterrows():
            in_delta = abs((s["in_time"] - b["in_time"]).total_seconds())
            out_delta = abs((s["out_time"] - b["out_time"]).total_seconds())
            score = in_delta + out_delta
            if best_score is None or score < best_score:
                best_score, best = score, b
        b = best
        in_delta_min = int(round((s["in_time"] - b["in_time"]).total_seconds() / 60))
        out_delta_min = int(round((s["out_time"] - b["out_time"]).total_seconds() / 60))

        common = {
            "employee_name": s["employee_name"],
            "supplier_card": s["card_id"],
            "breza_card": b["card_id"],
            "supplier_in": s["in_time"],
            "breza_in": b["in_time"],
            "supplier_out": s["out_time"],
            "breza_out": b["out_time"],
            "in_delta_min": in_delta_min,
            "out_delta_min": out_delta_min,
            "in_gate": b["in_gate"],
            "out_gate": b["out_gate"],
        }
        out_rows.append(common)

        if (s["in_time"] - b["in_time"]) > tol:
            out_rows.append({
                "issue_type": "LATE_IN",
                "employee_name": s["employee_name"],
                "supplier_card": s["card_id"], "breza_card": b["card_id"],
                "supplier_in": s["in_time"], "breza_in": b["in_time"],
                "details": "Kašnjenje ulaza: " + str(in_delta_min) + " min.",
            })
        if (b["out_time"] - s["out_time"]) > tol:
            out_rows.append({
                "issue_type": "EARLY_OUT",
                "employee_name": s["employee_name"],
                "supplier_card": s["card_id"], "breza_card": b["card_id"],
                "supplier_out": s["out_time"], "breza_out": b["out_time"],
                "details": "Raniji izlaz: " + str(abs(out_delta_min)) + " min.",
            })

    return pd.DataFrame(out_rows)


def detect_overlaps(brz_intervals: pd.DataFrame) -> pd.DataFrame:
    """Same person working overlapping intervals at different plants."""
    if brz_intervals.empty:
        return pd.DataFrame()
    rows = []
    df = brz_intervals.sort_values("in_time")
    for name_key, g in df.groupby("name_key"):
        g = g.sort_values("in_time")
        prev = None
        for _, cur in g.iterrows():
            if prev is not None and cur["in_time"] < prev["out_time"] \
                    and str(cur["plant_in"]) != str(prev["plant_in"]):
                rows.append({
                    "issue_type": "OVERLAP_DIFFERENT_PLANTS",
                    "employee_name": cur["employee_name"],
                    "plant_1": prev["plant_in"], "in_1": prev["in_time"], "out_1": prev["out_time"],
                    "plant_2": cur["plant_in"], "in_2": cur["in_time"], "out_2": cur["out_time"],
                    "details": "Preklapanje rada na različitim pogonima.",
                })
            prev = cur
    return pd.DataFrame(rows)


def _autosize_worksheet(writer, sheet_name, df, min_width=2, max_width=60):
    ws = writer.sheets.get(sheet_name)
    if ws is None:
        return
    for idx, col in enumerate(df.columns, start=1):
        body = df[col]
        longest = 0 if body.empty else body.map(lambda x: len(str(x)) if pd.notna(x) else 0).max()
        width = max(len(str(col)), int(longest))
        width = min(max(width + 2, min_width), max_width)
        ws.column_dimensions[get_column_letter(idx)].width = width


def _format_worksheet(writer, sheet_name, df):
    ws = writer.sheets.get(sheet_name)
    if ws is None:
        return
    ws.freeze_panes = "A2"
    last_col = get_column_letter(max(len(df.columns), 1))
    ws.auto_filter.ref = "A1:" + last_col + str(len(df) + 1)
    fill = PatternFill(fill_type="solid", fgColor="D9E1F2")
    font = Font(bold=True)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


# Internal issue code -> Serbian export label
_ISSUE_LABELS = {
    "LATE_IN": "KASNJENJE_ULAZ",
    "EARLY_OUT": "RANIJI_IZLAZ",
    "MISSING_ON_BREZA": "NEMA_U_BREZA_EVIDENCIJI",
    "WRONG_CARD_ID": "POGRESAN_ID_KARTICE",
    "OVERLAP_DIFFERENT_PLANTS": "PREKLAPANJE_RAZLICITI_POGONI",
}


def _translate_issue_types(df):
    if df.empty or "issue_type" not in df.columns:
        return df
    df = df.copy()
    df["tip_nepravilnosti"] = df["issue_type"].map(_ISSUE_LABELS).fillna(df["issue_type"])
    return df


def _rename_columns_for_export(df, mapping):
    if df.empty:
        return df
    return df.rename(columns=mapping)


def _build_timing_detail_sheets(matched, tol_min):
    """Return (late_in_df, early_out_df) sorted worst-first, Serbian headers."""
    late_cols = {"employee_name": "ime_prezime", "supplier_in": "ulaz_dobavljac",
                 "breza_in": "ulaz_breza", "in_delta_min": "kasnjenje_min"}
    early_cols = {"employee_name": "ime_prezime", "supplier_out": "izlaz_dobavljac",
                  "breza_out": "izlaz_breza", "out_delta_min": "raniji_izlaz_min"}
    if matched.empty:
        return pd.DataFrame(), pd.DataFrame()

    late = matched[matched.get("in_delta_min", pd.Series(dtype=float)).notna()] \
        [["employee_name", "supplier_in", "breza_in", "in_delta_min"]].copy()
    late = late.rename(columns=late_cols).sort_values("kasnjenje_min", ascending=False).reset_index(drop=True)

    early = matched[matched.get("out_delta_min", pd.Series(dtype=float)).notna()] \
        [["employee_name", "supplier_out", "breza_out", "out_delta_min"]].copy()
    early = early.rename(columns=early_cols).sort_values("raniji_izlaz_min", ascending=True).reset_index(drop=True)
    return late, early


def generate_report(supplier_csv, breza_csv, out_xlsx, tol_min=5):
    sup = parse_supplier(str(supplier_csv))
    ev = parse_breza_events(str(breza_csv))      # noqa: F841 (kept for parity)
    brz = parse_breza(str(breza_csv))

    matched = match_supplier_to_breza(sup, brz, tol_min=tol_min)
    overlaps = detect_overlaps(brz)

    issues = pd.concat([_translate_issue_types(matched[matched.get("issue_type").notna()]
                                               if "issue_type" in matched else pd.DataFrame()),
                        _translate_issue_types(overlaps)], ignore_index=True) \
        if not matched.empty else _translate_issue_types(overlaps)

    late, early = _build_timing_detail_sheets(matched, tol_min)

    n_issues = 0 if issues is None or issues.empty else len(issues)
    summary = pd.DataFrame([
        ("Broj redova - izveštaj dobavljača", len(sup)),
        ("Broj intervala - BREZA", len(brz)),
        ("Broj uparenih redova", int((matched.get("in_delta_min").notna()).sum()) if not matched.empty else 0),
        ("Ukupno nepravilnosti", n_issues),
        ("Kašnjenje ulaza", int((issues["issue_type"] == "LATE_IN").sum()) if n_issues else 0),
        ("Raniji izlaz", int((issues["issue_type"] == "EARLY_OUT").sum()) if n_issues else 0),
        ("Nema u BREZA evidenciji", int((issues["issue_type"] == "MISSING_ON_BREZA").sum()) if n_issues else 0),
        ("Pogrešan ID kartice", int((issues["issue_type"] == "WRONG_CARD_ID").sum()) if n_issues else 0),
        ("Preklapanje na različitim pogonima",
         int((issues["issue_type"] == "OVERLAP_DIFFERENT_PLANTS").sum()) if n_issues else 0),
    ], columns=["metrika", "vrednost"])

    issues_export = _rename_columns_for_export(issues, {
        "tip_nepravilnosti": "oznaka_nepravilnosti", "employee_name": "ime_prezime",
        "supplier_card": "kartica_dobavljac", "breza_card": "kartica_breza",
        "supplier_in": "ulaz_dobavljac", "supplier_out": "izlaz_dobavljac",
        "breza_in": "ulaz_breza", "breza_out": "izlaz_breza", "details": "opis",
        "plant_1": "pogon_1", "in_1": "ulaz_1", "out_1": "izlaz_1",
        "plant_2": "pogon_2", "in_2": "ulaz_2", "out_2": "izlaz_2",
    }) if issues is not None and not issues.empty else pd.DataFrame()

    matched_export = _rename_columns_for_export(
        matched[matched.get("in_delta_min").notna()] if not matched.empty else pd.DataFrame(), {
            "employee_name": "ime_prezime", "supplier_card": "kartica_dobavljac",
            "breza_card": "kartica_breza", "supplier_in": "ulaz_dobavljac",
            "breza_in": "ulaz_breza", "supplier_out": "izlaz_dobavljac",
            "breza_out": "izlaz_breza", "in_delta_min": "razlika_ulaz_min",
            "out_delta_min": "razlika_izlaz_min", "in_gate": "kapija_ulaz", "out_gate": "kapija_izlaz",
        })

    sup_export = _rename_columns_for_export(sup, {
        "row_id": "redni_broj", "employee_name": "ime_prezime", "name_key": "kljuc_imena",
        "card_id": "id_kartice", "nzn": "broj_nzn", "in_time": "ulaz", "out_time": "izlaz",
    })
    brz_export = _rename_columns_for_export(brz, {
        "row_id": "redni_broj", "employee_name": "ime_prezime", "name_key": "kljuc_imena",
        "card_id": "id_kartice", "in_time": "ulaz", "out_time": "izlaz",
        "in_gate": "kapija_ulaz", "out_gate": "kapija_izlaz",
        "plant_in": "pogon_ulaz", "plant_out": "pogon_izlaz",
    })

    out_xlsx = Path(out_xlsx)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        for name, frame in [
            ("REZIME", summary),
            ("NEPRAVILNOSTI", issues_export),
            ("UPARENO", matched_export),
            ("KASNJENJE_ULAZI", late),
            ("RANIJI_IZLAZI", early),
            ("DOBAVLJAC_NORMALIZOVANO", sup_export),
            ("BREZA_INTERVALI", brz_export),
        ]:
            frame.to_excel(writer, sheet_name=name, index=False)
            _autosize_worksheet(writer, name, frame)
            _format_worksheet(writer, name, frame)
    return out_xlsx


def main():
    ap = argparse.ArgumentParser(description="Poređenje BREZA i Izveštaja dobavljača")
    ap.add_argument("--supplier", default=DEFAULT_SUPPLIER_CSV,
                    help="Putanja do CSV fajla izveštaja dobavljača (default: " + DEFAULT_SUPPLIER_CSV + ")")
    ap.add_argument("--breza", default=DEFAULT_BREZA_CSV,
                    help="Putanja do CSV fajla BREZA evidencije (default: " + DEFAULT_BREZA_CSV + ")")
    ap.add_argument("--out", default="Nalaz.xlsx", help="Naziv izlaznog Excel fajla")
    ap.add_argument("--tol", type=int, default=5, help="Tolerancija u minutima")
    args = ap.parse_args()
    out = generate_report(args.supplier, args.breza, args.out, tol_min=args.tol)
    print("Završeno: " + str(Path(out).resolve()))


if __name__ == "__main__":
    main()
