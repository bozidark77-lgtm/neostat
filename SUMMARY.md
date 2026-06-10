# Summary — recovering source from Neostat_Analiza.exe

## The artifact

`Neostat_Analiza.exe` — a 42 MB Windows executable (PE32+, x86-64, GUI). It is a
**PyInstaller bundle of a Python 3.13 program**: a tkinter desktop app titled
**"AMS – Analiza izveštaja"** (powered by Neostat™).

## What the program does

It reconciles two data sources for plant/site access:

- **IZVEŠTAJ DOBAVLJAČA** — a supplier-reported work-hours report (.xlsx)
- **BREZA** — gate access-control events (.xlsx)

Pipeline: convert both workbooks to CSV → parse → pair BREZA IN/OUT events into
work intervals → match each supplier row to the best BREZA interval within a
minute tolerance → flag irregularities → write a styled 7-sheet Excel report
(`Nalaz.xlsx`: REZIME, NEPRAVILNOSTI, UPARENO, KASNJENJE_ULAZI, RANIJI_IZLAZI,
DOBAVLJAC_NORMALIZOVANO, BREZA_INTERVALI).

**Irregularity types detected**

| Code                       | Meaning                                          |
|----------------------------|--------------------------------------------------|
| `LATE_IN`                  | Supplier entry later than BREZA beyond tolerance |
| `EARLY_OUT`                | Supplier exit earlier than BREZA beyond tolerance|
| `MISSING_ON_BREZA`         | No BREZA event for that card + date              |
| `WRONG_CARD_ID`            | Same worker/date in BREZA under a different card |
| `OVERLAP_DIFFERENT_PLANTS` | Overlapping work logged at different plants      |

The whole application is three modules: `app.py` (GUI), `convert.py` (XLSX→CSV),
`analyze.py` (engine). Everything else in the bundle is the standard library and
third-party packages (pandas, openpyxl, numpy, PIL, tkinter).

## How it was recovered

1. **Identified the packer** — strings revealed PyInstaller + `python313.dll`.
2. **Parsed the CArchive** — located the `MEI\014\013\012\013\016` cookie, read
   the table of contents (1,709 entries), and confirmed `pyver = 313`.
3. **Extracted everything** — all 6 bootstrap scripts and unpacked the 8.8 MB
   `PYZ.pyz` (972 modules), reconstructing valid `.pyc` files with the correct
   Python 3.13 magic header.
4. **Isolated the app code** — filtered out stdlib/third-party to find the three
   application modules.
5. **Read the bytecode metadata** — Python 3.12's `marshal` successfully loaded
   the 3.13 code objects, exposing every function signature, class, docstring,
   import, and string/numeric constant.
6. **Reconstructed source** — wrote faithful, compilable `.py` for all three
   modules from that metadata.

## What's exact vs. reconstructed

- **Exact** (read directly from bytecode): function signatures, all constants,
  column aliases (incl. the Serbian header variants and typos), regexes, issue
  codes, Serbian UI labels and messages, colors/geometry, sheet names, export
  column maps.
- **Reconstructed** (inferred from the call sequence + pandas/tkinter idioms):
  statement-level control flow inside each function. It compiles and runs, but a
  branch or loop bound may differ from the original.

For a byte-exact decompile, run the original `.pyc` files through a 3.13-aware
decompiler — **PyLingual** (pylingual.io) is currently the most reliable; or
build **Decompyle++/pycdc**. (uncompyle6/decompyle3 do not support 3.13.)

## Running and rebuilding

- **Cross-platform source.** No Windows-only APIs; runs on Linux/macOS/Windows
  with `pandas`, `openpyxl`, and stdlib `tkinter` (Linux needs `python3-tk`).
  Run with `python src/app.py`.
- **No cross-compiling.** PyInstaller embeds the host OS's interpreter — a
  Windows `.exe` must be built on Windows. The included GitHub Actions workflow
  builds it on a `windows-latest` runner (and a Linux PyInstaller run would yield
  a Linux binary instead).

## Deliverables produced in this thread

- `Neostat_recovered/source/` — reconstructed `app.py`, `analyze.py`, `convert.py`
- `Neostat_recovered/app_pyc/` — exact 3.13 bytecode for the three modules
- `Neostat_recovered/Neostat_full_extraction.zip` — all 972 bundled modules + scripts
- `neostat-analiza/` — ready-to-commit repo: source, PyInstaller spec, CI workflow,
  build specification

## Suggested next steps

1. Create the repo from `neostat-analiza/` and push.
2. Add `neostat.ico` / `neostat_logo.png` to `assets/` (from the extraction zip).
3. Validate a CI-built `.exe` against a known-good `Nalaz.xlsx`.
4. If you need certainty on the original logic, decompile `app_pyc/*.pyc` with
   PyLingual and diff against `src/`.
5. Pin dependency versions before any release you intend to support.

## Rights note

Recovering source from your own binary, or one you're licensed to inspect, is
fine. If this binary isn't yours, check its license/EULA before redistributing
the recovered code.
