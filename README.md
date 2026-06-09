# Neostat_Analiza — AMS · Analiza izveštaja

Desktop tool (Neostat™) that reconciles a **supplier work-hours report**
(`IZVEŠTAJ DOBAVLJAČA`, `.xlsx`) against **BREZA gate access-control events**
(`AMS Team … BREZA`, `.xlsx`) and writes a styled, multi-sheet Excel report
(`Nalaz.xlsx`).

It pairs BREZA `ULAZ`/`IZLAZ` (in/out) events into work intervals, matches each
supplier row to the best BREZA interval within a time tolerance, and flags
irregularities.

| Code                       | Meaning                                            |
|----------------------------|----------------------------------------------------|
| `LATE_IN`                  | Supplier entry later than BREZA beyond tolerance   |
| `EARLY_OUT`                | Supplier exit earlier than BREZA beyond tolerance  |
| `MISSING_ON_BREZA`         | No BREZA event for that card + date                |
| `WRONG_CARD_ID`            | Same worker/date in BREZA under a different card   |
| `OVERLAP_DIFFERENT_PLANTS` | Overlapping work logged at different plants        |

The report (`Nalaz.xlsx`) has seven sheets: `REZIME`, `NEPRAVILNOSTI`,
`UPARENO`, `KASNJENJE_ULAZI`, `RANIJI_IZLAZI`, `DOBAVLJAC_NORMALIZOVANO`,
`BREZA_INTERVALI`.

## Run from source (any OS)

```bash
python -m pip install -r requirements.txt   # pandas, openpyxl
# Linux only — tkinter is bundled with Python on Windows/macOS:
#   sudo apt install python3-tk
python src/app.py
```

Pick the supplier `.xlsx` and the BREZA `.xlsx`, choose an output path, click
**Analiziraj**. The app converts both workbooks to CSV, runs the analysis, and
writes the report.

> Tip: keep your real input files in `data/` — that folder is git-ignored so the
> personal data (names, card IDs) never gets committed.

## Download a build

Prebuilt apps are produced by CI for every push and attached to tagged releases:

- **Windows:** `Neostat_Analiza.exe` (single file, double-click to run)
- **macOS (Apple Silicon):** `Neostat_Analiza-macos.zip` → unzip →
  `Neostat_Analiza.app`. The app is unsigned, so the first launch needs
  **right-click → Open → Open** (see `BUILD_SPEC.md` → *macOS code signing*).

Find them under **Actions → latest run → Artifacts**, or under **Releases** for
tagged versions.

## Build it yourself

```bash
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
# -> dist/Neostat_Analiza.exe   (Windows)
# -> dist/Neostat_Analiza.app   (macOS)
# -> dist/Neostat_Analiza       (Linux)
```

PyInstaller cannot cross-compile, so build each target on its own OS. The CI
workflow (`.github/workflows/build.yml`) does this on `windows-latest` and
`macos-14`. Full details and release instructions are in **`BUILD_SPEC.md`**.

## Tests

```bash
python tests/make_fixtures.py        # write sample supplier + BREZA workbooks
python tests/run_pipeline_smoke.py   # convert → analyze → assert the 7 sheets
```

The same two steps run in CI on Windows and macOS, so every build is proven to
produce a valid report end-to-end.

## Provenance

The `src/` modules were reconstructed from the original `Neostat_Analiza.exe`
(a PyInstaller bundle). Signatures, column aliases, issue codes, Serbian labels
and sheet names are exact; intra-function control flow is reconstructed. See
**`SUMMARY.md`** for the full recovery story and how to obtain a byte-exact
decompile if you need one.
