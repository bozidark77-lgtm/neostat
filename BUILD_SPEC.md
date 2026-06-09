# Build Specification — Neostat_Analiza

Reproducible build of the **AMS – Analiza izveštaja** desktop tool (Neostat™)
from recovered source into single-file desktop apps for **Windows and macOS**,
via GitHub Actions.

## Repository layout

```
neostat/
├── .github/workflows/build.yml   CI: build Windows .exe + macOS .app, publish on tag
├── src/
│   ├── app.py        tkinter GUI launcher (entry point, ConvertLauncherApp)
│   ├── analyze.py    engine: parse → match → detect overlaps → write .xlsx
│   └── convert.py    XLSX → CSV conversion (supplier + BREZA readers)
├── tests/
│   ├── make_fixtures.py        writes sample supplier + BREZA .xlsx
│   ├── run_pipeline_smoke.py   convert → analyze → assert 7 sheets + issue counts
│   └── fixtures/               (generated; git-ignored)
├── assets/
│   ├── neostat.ico         Windows window icon   (optional — add yours)
│   ├── neostat.icns        macOS app icon        (optional — add yours)
│   └── neostat_logo.png    window logo           (optional — add yours)
├── data/                   drop your real CSV/XLSX here (git-ignored)
├── Neostat_Analiza.spec    PyInstaller build recipe (cross-platform)
├── requirements.txt        runtime deps (pandas, openpyxl)
├── SUMMARY.md              how the source was recovered from the .exe
└── BUILD_SPEC.md           this file
```

`tkinter` ships with CPython on the Windows and macOS runners, so it is **not**
in `requirements.txt`. On Linux you install it via the system package
`python3-tk`.

## Build targets

| Property        | Windows                          | macOS                                  |
|-----------------|----------------------------------|----------------------------------------|
| Output          | `dist/Neostat_Analiza.exe`       | `dist/Neostat_Analiza.app` (zipped)    |
| Runner          | `windows-latest`                 | `macos-14` (Apple Silicon, arm64)      |
| Mode            | one-file, windowed (no console)  | one-file `.app` bundle, windowed       |
| Python          | 3.13                             | 3.13                                   |
| Packager        | PyInstaller ≥ 6.6                | PyInstaller ≥ 6.6                       |
| Runtime deps    | pandas ≥ 2.2.3, openpyxl ≥ 3.1.2, stdlib tkinter                          ||

PyInstaller **does not cross-compile** — it embeds the interpreter of the OS it
runs on. So the `.exe` must be built on Windows and the `.app` on macOS; the CI
matrix runs each on its matching runner. (Run the same spec on Linux and you get
an ELF binary instead.)

`openpyxl` is declared as a PyInstaller **hidden import** because pandas loads
the Excel engine lazily by name (`engine="openpyxl"`), which static analysis can
miss.

### macOS architecture

The macOS job builds on `macos-14` (**Apple Silicon / arm64**) only — the build
runs natively on M1/M2/M3+ Macs. To additionally ship an Intel build, add an
`os: macos-13` entry to the matrix in `build.yml`; it will produce a separate
x86_64 `.app`.

## CI pipeline (`.github/workflows/build.yml`)

**Triggers**
- push to `main` — build both platforms + upload artifacts
- pull request to `main` — build (validation only)
- push of a `v*` tag (e.g. `v1.0.0`) — build + publish a GitHub Release with both
  the `.exe` and the macOS `.zip` attached
- `workflow_dispatch` — manual run from the Actions tab

**`build` job** (matrix: `windows-latest`, `macos-14`)
1. **Checkout** the repo.
2. **Set up Python 3.13** with pip caching.
3. **Install** `requirements.txt` + PyInstaller.
4. **Build** via `pyinstaller --noconfirm --clean Neostat_Analiza.spec`.
5. **Import smoke test** — import `analyze` and `convert` to catch packaging-
   independent breakage early (the GUI is not launched; runners are headless).
6. **Functional pipeline test** — generate sample workbooks, run the real
   convert → analyze path, and assert the report has the seven expected sheets
   with the expected irregularity counts.
7. **Package** — on macOS, zip the `.app` with `ditto` (preserves the bundle's
   symlinks/metadata). On Windows the `.exe` is already a single file.
8. **Upload artifact** — `Neostat_Analiza-windows` (the `.exe`) and
   `Neostat_Analiza-macos` (the `.app` zip).

**`release` job** (only on `v*` tags)
- Downloads both build artifacts and publishes a single GitHub Release with the
  `.exe` and the macOS `.zip` attached, using auto-generated notes.

**Permissions:** the workflow requests `contents: write` so the release step can
create/attach a Release. Nothing else is granted.

## Releasing

```bash
git tag v1.0.0
git push origin v1.0.0
```

The tag push builds both platforms and publishes a Release with the `.exe` and
the macOS `.zip` attached. Without a tag, the build still runs and the artifacts
are downloadable from the run page (30-day retention).

## macOS code signing (Gatekeeper)

The `.app` is **unsigned and not notarized**. On Apple Silicon PyInstaller
ad-hoc-signs the binary so it runs, but Gatekeeper will still warn on first
launch (the app was downloaded from the internet). To open it:

> **right-click the app → Open → Open**, or run
> `xattr -dr com.apple.quarantine /path/to/Neostat_Analiza.app`

For a frictionless install you'd sign + notarize with an Apple Developer ID
(\$99/yr) — that needs certificate/secret setup and is intentionally out of
scope for this build. The hooks are easy to add to the macOS job later.

## Local builds

**Windows (produces the .exe):**
```powershell
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
.\dist\Neostat_Analiza.exe
```

**macOS (produces the .app):**
```bash
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
open dist/Neostat_Analiza.app
```

**Linux (produces a Linux binary, not a .exe):**
```bash
sudo apt install python3-tk
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
./dist/Neostat_Analiza
```

**Run from source, no packaging (any OS):**
```bash
cd src
python app.py
```

## Customization notes

- **App name / icon:** change `name=` and the icon files in
  `Neostat_Analiza.spec` / `assets/`.
- **Console for debugging:** set `console=True` in the spec to see tracebacks.
- **Reproducibility:** pin exact versions in `requirements.txt` and the
  PyInstaller version in the workflow before a release you intend to support.
- **Assets:** the icon/logo are optional. `app.py` wraps icon loading in
  `try/except`, and the spec only bundles files that exist — so the build is
  green even with an empty `assets/`.

## Validation reminder

The `src/` modules are a faithful **reconstruction** from the original binary's
bytecode. Signatures, constants, column maps, labels, and sheet names are exact;
intra-function control flow is reconstructed. The CI functional test proves the
pipeline runs and produces all seven sheets, but before relying on a rebuilt app
for production data, run it against a known-good `Nalaz.xlsx` and confirm the
seven output sheets match — or diff against a byte-exact decompile of the
original `.pyc` files (see `SUMMARY.md`).
