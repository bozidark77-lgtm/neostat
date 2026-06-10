# Contributing to Neostat_Analiza

Thanks for helping improve this tool. It reconciles a supplier work-hours report
against BREZA gate access-control events and writes a styled, multi-sheet Excel
report (see [`README.md`](README.md) for what it does and
[`BUILD_SPEC.md`](BUILD_SPEC.md) for the full build/architecture spec).

This guide applies to **human and AI contributors alike**. If you are an AI
coding agent (Claude Code, Copilot, Cursor, …), read the
[For AI agents](#for-ai-agents) section first — it is the contract for changes
made on your behalf.

## Project at a glance

```
src/
├── app.py        tkinter GUI launcher (entry point)
├── analyze.py    engine: parse → match → detect irregularities → write .xlsx
└── convert.py    XLSX → CSV conversion (supplier + BREZA readers)
tests/
├── make_fixtures.py        writes sample supplier + BREZA .xlsx
└── run_pipeline_smoke.py   convert → analyze → assert the 7 sheets + counts
.github/workflows/build.yml  CI: build Windows .exe, test, publish Release
Neostat_Analiza.spec         PyInstaller build recipe (cross-platform)
data/                        your real input files — git-ignored, never commit
```

## Development setup

Requires **Python 3.13** (CI pins 3.13; 3.11+ runs the tests fine).

```bash
python -m pip install -r requirements.txt   # pandas, openpyxl
# Linux only — tkinter is bundled with Python on Windows/macOS:
#   sudo apt install python3-tk
```

`tkinter` is only needed to launch the GUI (`src/app.py`). The analysis pipeline
and the tests use `pandas` + `openpyxl` only, so they run headless without it.

## Run from source

```bash
python src/app.py
```

Pick the supplier `.xlsx` and the BREZA `.xlsx`, choose an output path, click
**Analiziraj**.

## Tests — the gate for every change

Both steps must pass before you commit. They are exactly what CI runs:

```bash
python tests/make_fixtures.py        # write sample supplier + BREZA workbooks
python tests/run_pipeline_smoke.py   # convert → analyze → assert 7 sheets + counts
```

The smoke test asserts the report has the seven expected sheets, that each of
the five irregularity types is detected exactly once on the fixtures, and that
every finding names its plant (`pogon`). If you change detection logic, sheet
names, labels, or counts, update the fixtures and the expected values in
`tests/run_pipeline_smoke.py` in the same change — never loosen an assertion to
make a test pass.

## Building

PyInstaller does **not** cross-compile — build each target on its own OS:

```bash
python -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean Neostat_Analiza.spec
# -> dist/Neostat_Analiza.exe   (Windows)
# -> dist/Neostat_Analiza.app   (macOS)
# -> dist/Neostat_Analiza       (Linux)
```

## CI/CD

`.github/workflows/build.yml` runs on every push to `main`, every pull request
to `main`, every `v*` tag, and on manual dispatch. It builds the **Windows
`.exe`** on a `windows-latest` runner, runs the smoke + functional tests, then:

- **push to `main` (incl. a merged PR)** → publishes/updates a rolling
  `latest` **pre-release** carrying the newest `.exe`;
- **`v*` tag** (e.g. `v1.0.0`) → publishes a permanent, versioned Release.

Pull requests build and test but do **not** publish a Release. CI builds the
Windows target only; the macOS `.app` and Linux binary are local builds (above).

## Conventions

- **Match the surrounding code.** Mirror the existing naming, comment density,
  and idioms in the file you are editing. Keep diffs minimal and focused on the
  task — no drive-by reformatting or unrelated refactors.
- **Serbian text is data, not decoration.** Column aliases, sheet names, finding
  labels, and the Serbian report strings (`č`, `š`, `ž`, …) are matched against
  real workbooks. Preserve them verbatim unless the task is specifically to
  change them, and keep source files UTF-8.
- **Detection logic has known provenance.** `src/` parsing/labels were
  reconstructed from the original `.exe`; the detection logic was then
  deliberately reworked (see `BUILD_SPEC.md` → *Analysis changes*). If you touch
  it, update `BUILD_SPEC.md` and the test expectations to match.
- **Never commit real data.** `data/` and generated reports/fixtures are
  git-ignored. Do not add real names, card IDs, or gate logs to the repo, to
  tests, or to commit messages.
- **Keep docs in sync.** When you change behavior, the build, or the workflow,
  update `README.md` and `BUILD_SPEC.md` in the same change.

## Commits & pull requests

- Write clear, imperative commit messages (`Add …`, `Fix …`, `Remove …`) that
  say what changed and why.
- Keep one logical change per PR; ensure the tests above pass first.
- Do not open a pull request unless it was explicitly requested.

## For AI agents

Follow everything above, plus these specifics:

1. **Verify, don't assume.** Run both test commands and confirm they pass before
   committing. Report failures with their output — never claim success you did
   not observe, and never weaken a test to get green.
2. **Stay in scope.** Implement what was asked; if a change needs a decision the
   requester hasn't made (a design trade-off, a destructive step, anything
   outward-facing), ask before proceeding rather than guessing.
3. **Read before you edit.** Open the file and match its existing style; prefer
   small, surgical edits over rewrites.
4. **Preserve the Serbian strings and column aliases verbatim** — they are part
   of how the parser matches real-world workbooks.
5. **Keep secrets and personal data out** of code, tests, fixtures, and commit
   messages. Do not commit anything from `data/`.
6. **No tooling/model fingerprints in the repo.** Don't add agent or model
   identifiers, "generated by" banners, or similar to commit messages, PR
   text, code comments, or any committed file.
7. **Push only to the branch you were asked to work on**, and only open a PR
   when explicitly requested.
