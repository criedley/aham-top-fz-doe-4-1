# AHAM Top-Freezer DOE Energy Test Tool

A Python rewrite of a legacy Excel/VBA macro used to analyze AHAM energy test
data for top-freezer refrigerator/freezer units. It ingests a raw data-cart CSV,
detects compressor cycles and defrost events, verifies stability against AHAM
thresholds, computes EP1, EP2 and ET, and shows a report in the browser.

The original `.bas` VBA modules are kept in [`vba_modules/`](./vba_modules) for
reference. The port replaces them with a testable pipeline, and replaces the Excel
"Data Sheet" dropdowns with a browser form.

## Status

The cyclic-test pipeline works end to end and matches the original macro on one real
file. Fixed-window (100% duty cycle) tests are not supported end to end, and only one
real file has been checked (see [Known limitations](#known-limitations)).

| Module                | Replaces                                             | Status                                                                                   |
| --------------------- | ---------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| `io_import.py`        | `Module1.bas` (`ImportData`, `CorrectNames`)         | ✅ Validated on a real file                                                               |
| `setup.py`            | `Setup.bas` (`SetupVars`, `AvailChannel`, `RowCol`)  | ✅ Implemented, tested                                                                    |
| `cycles.py`           | `Setup.bas` (`GetCycles`) + `Module1.bas` (`Run100`) | ✅ Event-based path validated; fixed-window (`Run100`) path not validated                 |
| `checks.py`           | `Checks.bas` (stability, gradient, ambient)          | 🟡 Cyclic path validated on one file; fixed-window path not validated                     |
| `part_one.py`         | `PartOne.bas` (`EP1Calc`, `ECalc`)                   | 🟡 `calc_et` validated; `calc_ep1_fixed_window` (non-cyclic) not validated                |
| `defrost.py`          | `DefrostPortion.bas` (`EP2calc`, `CycleData`)        | 🟡 Cyclic path and T2 table validated; non-cyclic watt-band path not ported               |
| `report.py`           | `TestSummary.bas`                                    | 🟡 Summary and T2 table done; main cycle table lacks timing columns; no cell highlighting |
| `main.py`             | `AHAM()` orchestration                               | 🟡 `run_analysis` covers cyclic tests; raises `NotImplementedError` for fixed-window      |
| `webapp.py`           | `AddDataSheet`, `CheckUserInput`                     | ✅ Upload, configure, run, report                                                         |
| `channel_defaults.py` | `ParameterNames` sheet, Data Sheet defaults          | ✅ Alias map and default channel assignments                                              |
| `models.py`           | shared `Public` variables                            | ✅ Implemented                                                                            |
| `security.py`         | `SecurityModule.bas`                                 | 🚧 Not started                                                                            |

### Validation against the original macro

The sample file `7272_RAW-DATA___008_27JAN20261051_T25013475_001.csv` was run through
both the original macro and the port, with Unit Height 71, CT 10 and K-factor 1:

| Quantity                                           | Macro                         | Port          |
| -------------------------------------------------- | ----------------------------- | ------------- |
| Test duration                                      | 673.7 min                     | 673.7 min     |
| Energy / Energy Corr.                              | 1.677 kWh/day                 | 1.677 kWh/day |
| EP1 / T1                                           | 0.458 kWh / 431.8 min         | same          |
| EP2 / T2                                           | 0.447 kWh / 303.5 min         | same          |
| FF Avg / FZ Avg                                    | 37.64 / -0.6 °F               | same          |
| LSOC / RSOC Avg                                    | 90.4 / 90.26 °F               | same          |
| Gradients (L Lo-Mid, L Mid-Hi, R Lo-Mid, R Mid-Hi) | 0.043 / 0.031 / 0.197 / 0.027 | same          |

The six-row T2 table (A cycle, Precool, Defrost, Recovery, Post 1, B cycle) also
matches in temperatures, gradients, energy and times, apart from the differences listed
under [Known differences](#known-differences-from-the-original-macro). This is a single
file, so treat it as a strong sanity check, not full validation.

## How it works

1. A technician uploads a raw data-cart CSV through the web form.
2. `io_import.py` parses it. It reads the ragged metadata block separately from the
   uniform data block, drops open-thermocouple readings, extracts Unit On / Defrost
   events by detecting state changes, removes duplicate events, and normalizes column
   names against known aliases.
3. The form lists the file's columns. Each channel role (Watts, FF BW 1, RSOC HI and so
   on) is pre-selected from `channel_defaults.CHANNEL_MAP` where that column exists in the
   file. The technician checks the assignments and enters the unit height, CT, and the
   freezer-only options.
4. Submitting validates the input with the same functions the pipeline uses
   (`build_channel_config`, `TestConfig`), runs the analysis, shows the report, and saves
   the configuration to `config/channel_map.yaml`.

The analysis (`main.run_analysis`) runs these stages:

1. `cycles.detect_cycles`: compressor on/off cycles and per-cycle averages.
2. `checks.check_stability`: finds the stable cycle and the 7-hour T1 window (EP1, T1).
3. `defrost.calc_ep2`: finds the defrost, the A and B reference cycles, and EP2 / T2.
4. `part_one.calc_et`: ET and the K-factor-corrected value.
5. `report.build_report`: the numbers shown on the macro's "AHAM Report" sheet.

Columns for optional channels (FZ BW 4, FZ BW 5, ICE BW, Volts) appear in the report
tables only when those channels were assigned.

## Setup

### 1. Install uv

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
brew install uv      # macOS (Homebrew)
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

This creates `.venv` and installs everything in `pyproject.toml` / `uv.lock`. `uv` also
manages the Python version from [`.python-version`](./.python-version).

> **Note:** if you see a warning that `VIRTUAL_ENV` does not match the project
> environment, run `deactivate` and let `uv` manage the environment, or pass `--active`.

This project is a standalone application, not an installable package
(`[tool.uv] package = false`). Run everything with `uv run`.

## Usage

### Web form

```bash
uv run uvicorn webapp:app --reload --app-dir src
```

Open `http://127.0.0.1:8000`, upload the CSV, check the channel assignments, enter the
unit height and CT, and submit. Errors are shown on the form. A successful submission
shows the report.

**Unit height** is the value the technician enters on the macro's Data Sheet. The tool
adds 10 inches for the thermocouple offset, as the macro does. Do not use the CSV
header's `CABINET_HEIGHT`, which can differ (7272: header 68.5, Data Sheet 71).

### From Python

`main.run_analysis(csv_path, channel_map, test_config)` runs the pipeline and returns an
`AnalysisResult`. `report.build_report(analysis, test_config)` turns that into report
values. See `tests/test_main.py` and `tests/test_report.py` for examples.

## Running the tests

```bash
uv run pytest            # everything
uv run pytest -v         # list each test
uv run pytest tests/test_defrost.py
uv run pytest -x --tb=long   # stop at first failure, full traceback
```

- `tests/conftest.py` holds shared fixtures (`alias_map`, `channel_map`,
  `channel_config`, `test_config`, `real_test_config`, `real_pipeline`).
- `tests/factories.py` builds synthetic raw data, events and cycle data.
- `tests/fixtures/sample_raw.csv` is the real file used by the macro-comparison tests.
- `test_checks_real_file.py`, `test_t1_averages.py`, `test_defrost.py`, `test_main.py`,
  `test_report.py` and `test_t2_table.py` compare the port against the macro's output
  for that file. `test_webapp.py` runs the web app end to end.
- Bugs found against real data have regression tests. Add one whenever you fix a bug,
  and confirm it fails before the fix and passes after.

## Project structure

```
.
├── config/        # saved technician configuration
├── src/
│   ├── models.py            # shared dataclasses
│   ├── channel_defaults.py  # alias map and default channel assignments
│   ├── io_import.py         # CSV parsing, cleaning, event extraction
│   ├── setup.py             # resolves technician channel assignments
│   ├── cycles.py            # cycle detection and per-cycle statistics
│   ├── checks.py            # stability / gradient / ambient checks
│   ├── part_one.py          # EP1 (fixed window) and ET
│   ├── defrost.py           # EP2 / T2 and the T2 table
│   ├── report.py            # report values
│   ├── main.py              # run_analysis
│   └── webapp.py            # FastAPI form and report page
├── static/        # web CSS
├── templates/     # upload, channel form and report pages (Jinja2)
├── tests/         # pytest suite and fixtures
├── vba_modules/   # original .bas files, for reference
└── pyproject.toml
```

Modules communicate through the dataclasses in `models.py` (plus `DefrostResult` and
`ReportData`), so each can be tested on its own.

## Design notes for contributors

- **No global state.** The VBA used about 150 module-level `Public` variables. The port
  passes explicit dataclasses as arguments and return values.
- **Validation lives in one place.** Rules such as "CT must be positive", "height must be
  positive" or "K-factor must be 1, 0.85 or 0.7" are enforced in `models.py` and
  `setup.py`. The web form calls those functions instead of repeating the rules.
- **Events are detected by state change, not by blank rows.** The macro patched missing
  events by inserting synthetic rows. The port detects changes in the state columns
  directly. Event counters and `Watt-Hours` are never used to decide whether two events
  are the same state.
- **A Unit On boundary means Unit On changed.** Defrost-only or Comp-only rows can carry
  a Unit On value without a Unit On transition, and must not be treated as cycle
  boundaries.
- **T2 segments are bounded by compressor-off events.** `defrost.py` works from the
  events table, and uses raw rows only for the ±0.5 °F reference-cycle checks.
- **Match the macro first.** This tool feeds compliance results. Where the port differs
  from the VBA, document why, in the code and below. Some quirks are deliberate: test
  duration is measured from the second scan because that matches the macro's output.

## Known differences from the original macro

Each of these is intentional.

- **Ambient / Gradient flags.** The port checks the main cycle table up to the stable
  cycle only. The macro also tests the T2 rows with `If I <= stabcycle`, which compares a
  T2 row number with the stable cycle number. That looks accidental, so it isn't
  reproduced. A file with an out-of-band T2 segment could show TRUE here where the macro
  shows FALSE.
- **FF channels in the cyclic stability check.** The port checks FF and FZ (ETP 8.4.1:
  "all measured compartments"). In the macro's cyclic `StabilityCheck` the FF window
  averages are never assigned, so its FF comparison always passes. The port can reject a
  test the macro accepts.
- **K-factor.** Applied to ET unconditionally, as in the macro's cyclic path. The macro's
  non-cyclic `ECalc` applies it only when freezer-only is set. Only K = 1 has been
  checked against the macro.
- **Duty cycle.** On-time divided by cycle duration, as in the macro report's
  "% Runtime". The Cycles-sheet formula in `Setup.bas` subtracts one minute, which gives
  values above 100% for a cycle with almost no off time.
- **Freezer-only runs.** The macro calls `EP2calc` before it computes the T1 averages, so
  it compares against unset or stale values. The port uses the computed T1 averages, so
  results can differ on those files.
- **Cabinet height** is technician input plus 10 inches. The CSV header's
  `CABINET_HEIGHT` is not used.
- **Defrost / Recovery boundary.** When a defrost-off event and the compressor-off event
  fall seconds apart (7272: 17:55:14 and 17:55:16), the macro's import step keeps the
  earlier timestamp and drops the later row. The port uses the later one, the actual
  compressor-off. Defrost and Recovery durations differ from the macro by about 2 s.
  Energy (Wh), EP2, T2 and ET are unaffected.
- **Rounding.** The macro uses Excel's ROUND (ties round up). The port uses Python's
  round, which can round an exact tie down. Displayed averages can differ from the macro
  by 0.01 in those cases (7272: Defrost LSOC 90.415 shows as 90.41, the macro shows
  90.42). No energy result is affected.
- **Window midpoints in the cyclic stability check** use the last row of each window,
  where the VBA uses cycle on/off times. This changes the threshold slightly and is not
  yet matched.

## Known limitations

- **Only one real file** has been checked against the original macro. A file with a
  different defrost position, an FZ BW 4/5 channel, or K ≠ 1 would test more.
- **Fixed-window (non-cyclic) tests are not supported end to end.** `run_analysis` raises
  `NotImplementedError`. The `Run100` cycle search, the fixed-window stability search and
  `calc_ep1_fixed_window` exist but are unvalidated, and the watt-band EP2 path isn't
  ported.
- **The defrost-restart path** in the cyclic stability search (a defrost before stability
  is confirmed) is only covered by a skipped synthetic test. The 7272 file stabilizes
  before its defrost.
- **The report is incomplete.** The main cycle table lacks on-time, off-time, duration,
  Wh and end-of-cycle columns. There is no cell highlighting (voltage outside 114–116 V,
  LSOC/RSOC outside 88.95–91.05 °F, gradient over the limit) and no export to Excel.
- **CT is entered directly.** The macro takes CTL and CTM and computes CT; the port does
  not.
- **The VBA's second unconditional duplicate-event pass isn't ported.**
- **`security.py` is not started.**
- The tool is for trusted lab use. Uploads are saved under their base filename in
  `data/`, and the form only accepts paths inside that folder, but it hasn't been
  hardened for public exposure.