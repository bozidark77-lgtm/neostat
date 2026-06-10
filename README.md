# Neostat_Analiza — AMS · Analiza izveštaja

Desktop tool (Neostat™) that reconciles a **supplier work-hours report**
(`IZVEŠTAJ DOBAVLJAČA`, `.xlsx`) against **BREZA gate access-control events**
(`AMS Team … BREZA`, `.xlsx`) and writes a styled, multi-sheet Excel report
(`Nalaz.xlsx`).

It pairs BREZA `ULAZ`/`IZLAZ` (in/out) events into work intervals, compares each
supplier row's claimed hours to the worker's actual BREZA presence that day, and
flags irregularities — **each labelled with the plant (pogon)**.

| Code                       | Meaning                                                                     |
|----------------------------|-----------------------------------------------------------------------------|
| `DUPLIRANI_SATI`           | Same worker billed for overlapping hours the same day at the same plant (e.g. 2×8h) |
| `MANJAK_SATI`              | Claimed hours (`Sati rada`) exceed actual BREZA presence — late in, early out, or a mid-shift exit |
| `OVERLAP_DIFFERENT_PLANTS` | Same worker logged at two different plants with overlapping hours           |
| `MISSING_ON_BREZA`         | No BREZA event for that card + date                                         |
| `WRONG_CARD_ID`            | Same worker/date in BREZA under a different card                            |

The report (`Nalaz.xlsx`) has seven sheets: `REZIME`, `NEPRAVILNOSTI`,
`DUPLIRANI_SATI`, `MANJAK_SATI`, `UPARENO`, `DOBAVLJAC_NORMALIZOVANO`,
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

CI builds the Windows app on every push to `main` and attaches it to a GitHub
**Release** — no need to build it yourself:

- **Latest dev build:** the rolling **`latest`** pre-release under **Releases**
  always carries the newest `.exe` built from `main`.
- **Stable build:** tagged **Releases** (`v*`) carry a fixed, versioned `.exe`.

Either way the download is `Neostat_Analiza.exe` — a single file, double-click to
run. The `.exe` is also kept as a build **artifact** on each Actions run (30-day
retention).

> **macOS / Linux:** CI builds the Windows target only. Build the `.app` or a
> Linux binary locally with the steps below — PyInstaller can't cross-compile.

## Build it yourself

```bash
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
# -> dist/Neostat_Analiza.exe   (Windows)
# -> dist/Neostat_Analiza.app   (macOS)
# -> dist/Neostat_Analiza       (Linux)
```

PyInstaller cannot cross-compile, so build each target on its own OS. The CI
workflow (`.github/workflows/build.yml`) builds the Windows `.exe` on
`windows-latest`; build the macOS `.app` and the Linux binary locally. Full
details and release instructions are in **`BUILD_SPEC.md`**.

## Tests

```bash
python tests/make_fixtures.py        # write sample supplier + BREZA workbooks
python tests/run_pipeline_smoke.py   # convert → analyze → assert the 7 sheets
```

The same two steps run in CI on Windows, so every build is proven to produce a
valid report end-to-end.

## Contributing

Setup, the test gate, build/CI details, and project conventions — including
**instructions for AI agents** — are in
**[`CONTRIBUTING.md`](CONTRIBUTING.md)**.

## Provenance

The `src/` modules were reconstructed from the original `Neostat_Analiza.exe`
(a PyInstaller bundle); the parsing, column aliases, Serbian labels and styling
are kept verbatim. The **detection logic was then reworked** to fix the errors
HBIS reported for May 2026 (duplicated hours, mid-shift short hours, and showing
the plant) and to cut the false-positive noise — so the engine no longer matches
the original binary by design. See **`BUILD_SPEC.md` → "Analysis changes"** for
what changed and why, and **`SUMMARY.md`** for the recovery story.
