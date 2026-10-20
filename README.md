# AHAM Top-Freezer DOE Energy Test Tool

A Python rewrite of a legacy Excel/VBA macro used to analyze AHAM energy
test data for top-freezer refrigerator/freezer units. It ingests raw
data-cart CSV exports, detects compressor on/off cycles and defrost
events, verifies test stability against AHAM thresholds, and (once
complete) computes the standardized energy figures (EP1, EP2, ET) used in
DOE/AHAM compliance reporting.

The original `.bas` VBA modules are kept in [`vba_modules/`](./vba_modules)
for reference. The Python port replaces them with a testable pipeline and
replaces the Excel "Data Sheet" dropdowns with a browser-based form.

## Status

This is an active port. The cyclic-test pipeline (import, cycles, stability,
EP1/T1, EP2/T2, ET) reproduces the original macro on one real file. Fixed-window
(100% duty cycle) tests, the rendered report and the web-app integration are not
finished.

| Module         | Replaces                                             | Status                                                                                                    |
| -------------- | ---------------------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `io_import.py` | `Module1.bas` (`ImportData`, `CorrectNames`)         | ✅ Implemented, tested, validated on a real file                                                           |
| `setup.py`     | `Setup.bas` (`SetupVars`, `AvailChannel`, `RowCol`)  | ✅ Implemented, tested                                                                                     |
| `cycles.py`    | `Setup.bas` (`GetCycles`) + `Module1.bas` (`Run100`) | ✅ Event-based path validated on a real file; fixed-window (`Run100`) path not yet validated               |
| `checks.py`    | `Checks.bas` (stability, gradient, ambient)          | 🟡 Cyclic stability path matches the original macro on one file; other paths not yet validated (see below) |
| `models.py`    | shared `Public` variables                            | ✅ Implemented                                                                                             |
| `webapp.py`    | `AddDataSheet`, `CheckUserInput`                     | ✅ Collects and validates technician input                                                                 |
| `part_one.py`  | `PartOne.bas` (`EP1Calc`, `ECalc`)                   | ✅ `calc_et` validated on a real file; `calc_ep1_fixed_window` (non-cyclic) not validated                  |
| `defrost.py`   | `DefrostPortion.bas` (`EP2calc`)                     | 🟡 Cyclic path validated on a real file; non-cyclic watt-band path not ported                              |
| `report.py`    | `TestSummary.bas`                                    | 🟡 Report numbers built; rendering and the T2 segment table not done                                       |
| `security.py`  | `SecurityModule.bas`                                 | 🚧 Not started                                                                                             |
| `main.py`      | `AHAM()` orchestration                               | 🟡 `run_analysis` covers cyclic tests; raises `NotImplementedError` for fixed-window                       |
### Validation against the original macro

The sample file `7272_RAW-DATA___008_27JAN20261051_T25013475_001.csv` was
run through both the original macro and this port. The stable cycle, T1
window, `t1`, `EP1` and test duration agree (for example `EP1` = 0.458 kWh,
`t1` ≈ 431.8 min, test duration ≈ 673.7 min). This is a single file, so
treat it as a sanity check, not full validation.

## How it works

1. A technician uploads a raw data-cart CSV through a web form.
2. `io_import.py` parses the file. It reads the ragged metadata block
   separately from the uniform data block, drops open-thermocouple
   readings, extracts Unit On / Defrost events by detecting state changes,
   removes duplicate events, and normalizes column names against known
   aliases.
3. The technician assigns each sensor role (Watts, FF BW 1, RSOC HI, and
   so on) to a column from the uploaded file and enters test parameters
   (cabinet height, CT value, freezer-only K-factor).
4. Submitting the form validates the input using the same functions the
   pipeline uses (`build_channel_config`, `TestConfig`) and saves it to
   `config/channel_map.yaml`. **It does not run the analysis.**
5. The pipeline detects compressor cycles (`cycles.py`) and checks
   stability (`checks.py`). EP1/EP2/ET calculation and reporting are not
   implemented yet.

Saving the configuration is intentionally separate from running the
analysis, so a saved configuration can be re-used or inspected without
resubmitting the form.

## Setup

### 1. Install uv

If you don't already have `uv` installed:

**macOS / Linux (including WSL):**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Or via a package manager:
```bash
# macOS (Homebrew)
brew install uv

# pipx
pipx install uv
```

Verify it installed:
```bash
uv --version
```

See the [uv installation docs](https://docs.astral.sh/uv/getting-started/installation/)
for other options.

### 2. Clone and sync the project

```bash
git clone <repo-url>
cd aham-top-fz-doe-4-1
uv sync
```

This creates `.venv` and installs everything in `pyproject.toml` /
`uv.lock`. `uv` also manages the Python version from
[`.python-version`](./.python-version), so you don't need to install it
separately.

> **Note:** if you see a warning that `VIRTUAL_ENV` does not match the
> project environment, run `deactivate` and let `uv` manage the
> environment, or pass `--active` to use the activated one.

This project is a standalone application, not an installable package
(`[tool.uv] package = false`). Run everything with `uv run`.

## Usage

### 1. Start the web form

```bash
uv run uvicorn webapp:app --reload --app-dir src
```

Open `http://127.0.0.1:8000`, upload the CSV, assign channels, enter the
test parameters, and submit. Errors are shown on the form so you can
correct them. A valid submission is saved to `config/channel_map.yaml`.

### 2. Run the pipeline

```bash
uv run python src/main.py
```

> `main.py` does not yet run the full pipeline. Energy calculation and
> report generation are not implemented (see [Status](#status)).

## Running the tests

```bash
uv run pytest            # everything
uv run pytest -v         # list each test
uv run pytest tests/test_checks.py
uv run pytest -x --tb=long   # stop at first failure, full traceback
```

- `tests/conftest.py` holds shared pytest fixtures (`alias_map`,
  `channel_map`, `channel_config`, `test_config`).
- `tests/fixtures/` holds the sample CSV used by the real-file tests.
- `tests/test_checks_real_file.py` compares the port against the original
  macro's output for the sample file.
- Bugs found by running against real data have regression tests. Add one
  whenever you fix a bug, and confirm it fails before the fix and passes
  after.

## Project structure

```
.
├── config/        # saved technician configuration
├── src/
│   ├── models.py      # shared dataclasses
│   ├── io_import.py   # CSV parsing, cleaning, event extraction
│   ├── setup.py       # resolves technician channel assignments
│   ├── cycles.py      # cycle detection and per-cycle statistics
│   ├── checks.py      # stability / gradient / ambient checks
│   ├── webapp.py      # FastAPI technician input form
│   └── main.py        # pipeline entry point
├── static/        # web form CSS
├── templates/     # web form HTML (Jinja2)
├── tests/         # pytest suite and fixtures
├── vba_modules/   # original .bas files, for reference
└── pyproject.toml
```

Modules communicate only through the dataclasses in `models.py`, so each
can be tested on its own.

## Design notes for contributors

- **No global state.** The VBA used about 150 module-level `Public`
  variables. The port passes explicit dataclasses (`ImportResult`,
  `ChannelConfig`, `TestConfig`, `CycleData`) as arguments and return
  values.
- **Validation lives in one place.** Rules such as "CT must be positive"
  or "K-factor must be 1, 0.85 or 0.7" are enforced in `models.py` and
  `setup.py`. The web form calls those functions instead of duplicating
  the rules.
- **Events are detected by state change, not by blank rows.** The macro
  patched missing events by inserting synthetic rows. The port detects
  changes in the state columns directly. Event counters and `Watt-Hours`
  are never used to decide whether two events are the same state.
- **A Unit On boundary means Unit On changed.** Defrost-only or Comp-only
  rows can carry a Unit On value without a Unit On transition, and must
  not be treated as cycle boundaries.
- **Match the macro first.** This tool feeds compliance results. Where the
  port differs from the VBA, document why. Some quirks are deliberate: for
  example, test duration is measured from the second scan because that
  matches the macro's output.


## Known differences from the original macro

Where the port deliberately differs from the VBA. Each of these is intentional.

- **Ambient / Gradient flags.** The port checks the main cycle table up to the
  stable cycle only. The macro also tests the T2 segment rows using
  `If I <= stabcycle`, which compares a T2 row number with the stable cycle
  number. That looks accidental, so it isn't reproduced. A file with an
  out-of-band T2 segment could show `TRUE` here where the macro shows `FALSE`.
- **Cabinet height is technician input.** Gradients use the height entered on the
  Data Sheet plus 10" for thermocouple offset. The CSV header's `CABINET_HEIGHT` can
  differ (7272: header 68.5, entered 71) and must not be used as a substitute.
- **FF channels in the cyclic stability check.** The port checks FF and FZ
  (ETP 8.4.1: "all measured compartments"). In the macro's cyclic
  `StabilityCheck` the FF window averages are never assigned, so the FF
  comparison always passes. The port can reject a test the macro accepts.
- **K-factor.** Applied to `ET` unconditionally, as in the macro's cyclic
  path. The macro's non-cyclic `ECalc` applies it only when freezer-only is
  set. Only K = 1 has been checked against the macro.
- **Duty cycle.** Computed as on-time divided by cycle duration, as in the
  macro report's "% Runtime". The Cycles-sheet formula in `Setup.bas`
  subtracts one minute, which gives values above 100% for cycles with almost
  no off time.
- **Freezer-only runs.** The macro calls `EP2calc` before it computes the T1
  averages, so it compares against unset or stale values. The port uses the
  computed T1 averages. Results can differ from the macro on those files.
- **Window midpoints in the cyclic stability check** use the last row of each
  window, where the VBA uses cycle on/off times. This changes the threshold
  slightly and is not yet matched.
- **Defrost / Recovery boundary.** When a defrost-off event and the compressor-off
  event fall seconds apart (7272: 17:55:14 and 17:55:16), the macro's import step keeps
  the earlier timestamp and drops the later row. The port uses the later one, the actual
  compressor-off. Defrost and Recovery durations differ from the macro by about 2 s.
  Energy (Wh), EP2, T2 and ET are unaffected.
- **Rounding.** The macro uses Excel's ROUND (ties round up). The port uses Python's
  round, which can round an exact tie down. Displayed averages can differ from the
  macro by 0.01 in those cases. No energy result is affected.

## Known limitations

- Fixed-window (non-cyclic) tests aren't supported end to end: `run_analysis`
  raises `NotImplementedError`.
- The defrost-restart path in the cyclic stability check is only covered
  by a skipped synthetic test. The sample file stabilizes before its
  defrost, so real data hasn't exercised it.
- The non-cyclic stability path and the 100%-duty-cycle fallback
  (`Run100`) haven't been validated against real files.
- The VBA's second unconditional duplicate-event pass isn't ported.
- Only one real data file has been checked against the original macro.
- The tool is intended for trusted lab use. File uploads are saved under a
  sanitized filename and it hasn't been hardened for public exposure.