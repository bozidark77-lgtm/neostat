# Build Specification — Neostat_Analiza

Reproducible build of the **AMS – Analiza izveštaja** desktop tool (Neostat™)
from recovered source into single-file desktop apps. **CI (GitHub Actions)
builds, tests, and publishes the Windows `.exe`**; the macOS `.app` and the
Linux binary are produced by running the same spec locally on those systems.

## Repository layout

```
neostat/
├── .github/workflows/build.yml   CI: build + test Windows .exe, publish Release on main/tag
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

`tkinter` ships with CPython on Windows and macOS, so it is **not**
in `requirements.txt`. On Linux you install it via the system package
`python3-tk`.

## Build targets

| Property        | Windows                          | macOS                                  |
|-----------------|----------------------------------|----------------------------------------|
| Output          | `dist/Neostat_Analiza.exe`       | `dist/Neostat_Analiza.app` (zipped)    |
| Built by        | CI on `windows-latest`           | locally on a Mac (no CI runner)        |
| Mode            | one-file, windowed (no console)  | one-file `.app` bundle, windowed       |
| Python          | 3.13                             | 3.13                                   |
| Packager        | PyInstaller ≥ 6.6                | PyInstaller ≥ 6.6                       |
| Runtime deps    | pandas ≥ 2.2.3, openpyxl ≥ 3.1.2, stdlib tkinter                          ||

PyInstaller **does not cross-compile** — it embeds the interpreter of the OS it
runs on. So the `.exe` must be built on Windows and the `.app` on macOS. CI
builds only the Windows target; build the `.app` on a Mac and the Linux ELF
binary on Linux from the same spec (see *Local builds* below).

`openpyxl` is declared as a PyInstaller **hidden import** because pandas loads
the Excel engine lazily by name (`engine="openpyxl"`), which static analysis can
miss.

### macOS architecture

A local PyInstaller build produces a `.app` for the architecture of the Mac it
runs on — Apple Silicon (arm64) on M1/M2/M3+, x86_64 on Intel. To ship both,
build on each. There is no macOS CI runner, so the `.app` is not built or
published automatically.

## CI pipeline (`.github/workflows/build.yml`)

CI builds the **Windows `.exe`** only (see *Build targets*).

**Triggers**
- push to `main` (including a merged PR) — build + test, then publish/update the
  rolling `latest` pre-release with the new `.exe`
- pull request to `main` — build + test (validation only; no Release)
- push of a `v*` tag (e.g. `v1.0.0`) — build + test, then publish a permanent,
  versioned GitHub Release with the `.exe` attached
- `workflow_dispatch` — manual run from the Actions tab

Runs are serialized per ref (`concurrency`) so two quick pushes to `main` can't
race on the rolling release.

**`build` job** (`windows-latest`)
1. **Checkout** the repo.
2. **Set up Python 3.13** with pip caching.
3. **Install** `requirements.txt` + PyInstaller.
4. **Build** via `pyinstaller --noconfirm --clean Neostat_Analiza.spec`.
5. **Import smoke test** — import `analyze` and `convert` to catch packaging-
   independent breakage early (the GUI is not launched; the runner is headless).
6. **Functional pipeline test** — generate sample workbooks, run the real
   convert → analyze path, and assert the report has the seven expected sheets
   with the expected irregularity counts.
7. **Upload artifact** — `Neostat_Analiza-windows` (the `.exe`), 30-day retention.

**`release` job** (push events only)
- **`v*` tag** → publishes a versioned GitHub Release with the `.exe` attached,
  using auto-generated notes.
- **push to `main`** → drops the previous `latest` release + tag (the Releases
  API won't move an existing tag) and recreates a `latest` **pre-release** at the
  new commit with the fresh `.exe`. Pull requests never reach this job.

**Permissions:** the workflow requests `contents: write` so the release step can
create/attach Releases. Nothing else is granted.

## Releasing

Every push to `main` already publishes the newest `.exe` to the rolling `latest`
pre-release. For a permanent, versioned Release, push a tag:

```bash
git tag v1.0.0
git push origin v1.0.0
```

The tag push builds the Windows `.exe` and publishes a versioned Release with it
attached. The `.exe` is also downloadable from each run's artifacts (30-day
retention).

## macOS code signing (Gatekeeper)

The `.app` is **unsigned and not notarized**. On Apple Silicon PyInstaller
ad-hoc-signs the binary so it runs, but Gatekeeper will still warn on first
launch (the app was downloaded from the internet). To open it:

> **right-click the app → Open → Open**, or run
> `xattr -dr com.apple.quarantine /path/to/Neostat_Analiza.app`

For a frictionless install you'd sign + notarize with an Apple Developer ID
(\$99/yr) — that needs certificate/secret setup and is intentionally out of
scope for this build. You'd run those steps as part of a macOS build — locally,
or in a macOS CI job if one is added later.

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

## Analysis changes (May 2026 HBIS fixes)

The original engine missed the errors HBIS flagged for May 2026 and buried them
under false positives (≈1,280 of 1,316 findings were `LATE_IN`/`EARLY_OUT` from
comparing rounded supplier times to exact gate times). The detection logic in
`analyze.py` was reworked accordingly — it **no longer matches the original
binary**, by design:

| Finding (label)               | What it catches                                                        |
|-------------------------------|------------------------------------------------------------------------|
| `DUPLIRANI_SATI`              | A worker's same-day rows whose hours overlap on the **same plant** (e.g. duplicated 2×8h). The original only checked BREZA and skipped same-plant overlaps. |
| `MANJAK_SATI`                 | Claimed `Sati rada` exceeds the worker's **actual BREZA presence** for the day (sum of intervals, so a mid-shift exit reduces it). Replaces the noisy per-timestamp late/early checks. |
| `OVERLAP_DIFFERENT_PLANTS`    | Same worker logged at two different plants with overlapping hours (now from the supplier report, not gate-code prefixes). |
| `MISSING_ON_BREZA` / `WRONG_CARD_ID` | Unchanged in meaning; now carry the plant.                     |

Supporting changes:

- **Plant (pogon) on every finding.** The supplier report is one sheet per plant;
  `convert.py` now writes a `Pogon` column from the sheet name (stripping the
  `Izvestaj` prefix and `(2)` copy-marker), and `analyze.py` threads it through.
- **Claimed hours.** `analyze.py` reads `Sati rada` (falling back to out−in).
- **Tolerance.** `--tol` (default **5 min**) is the *hours-shortfall* threshold
  for `MANJAK_SATI` — the original 5-minute value, but now applied to total daily
  hours instead of per-timestamp. On the May data anything from 0–29 min yields
  the same result; the margin only matters for trivial edge cases in future months.
- **Sheets.** `KASNJENJE_ULAZI` / `RANIJI_IZLAZI` are replaced by
  `DUPLIRANI_SATI` / `MANJAK_SATI`; `UPARENO` now shows claimed vs actual hours.

Validated end-to-end against the real May workbooks: the reworked engine
reproduces every worker HBIS named (4 duplicated-hours + 3 short-hours cases),
labels each with its plant, and cuts total findings from 1,316 to 45.

## Validation reminder

The original `src/` parsing/labels were a faithful **reconstruction** from the
binary's bytecode; the detection logic has since been deliberately changed (see
above). The CI functional test proves the pipeline runs and produces all seven
sheets with the expected findings and plant on every row. Before relying on a
rebuilt app for a new month's data, spot-check a few `DUPLIRANI_SATI` /
`MANJAK_SATI` rows against the source workbooks.
