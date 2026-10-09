# AHAM Top-Freezer DOE Energy Test Tool

A Python rewrite of a legacy Excel/VBA macro used to analyze AHAM energy
test data for top-freezer refrigerator/freezer units. It ingests raw
data-cart CSV exports, detects compressor on/off cycles and defrost
events, verifies test stability against AHAM thresholds, and computes
the standardized energy consumption figures (EP1, EP2, ET) used in
DOE/AHAM compliance reporting.

This project replaces the original `.bas` VBA modules (kept under
[`vba_modules/`](./vba_modules) for reference) with a modern,
testable Python pipeline, and a browser-based form for technician input
in place of the old Excel "Data Sheet" dropdowns.

## Status

This is an active port, not a finished tool. Some parts of the pipeline
are implemented and usable; others are still being ported from VBA.

| Module | Replaces | Status |
|---|---|---|
| `io_import.py` | `Module1.bas` (`ImportData`, `backupimport`, `CorrectNames`) | ✅ Implemented |
| `setup.py` | `Setup.bas` (`SetupVars`, `AvailChannel`, `RowCol`) | ✅ Implemented |
| `cycles.py` | `Setup.bas` (`GetCycles`) + `Module1.bas` (`Run100`) | ✅ Implemented |
| `models.py` | shared `Public` variable declarations | ✅ Implemented |
| `webapp.py` | `AddDataSheet`, `CheckUserInput` (technician input UI) | ✅ Implemented |
| `checks.py` | `Checks.bas` (stability/gradient/ambient checks) | 🚧 Not yet started |
| `defrost.py` | `DefrostPortion.bas` (`EP2calc`) | 🚧 Not yet started |
| `part_one.py` | `PartOne.bas` (`EP1Calc`, `ECalc`) | 🚧 Not yet started |
| `report.py` | `TestSummary.bas` (`PrintReport`, `AHAMSummary`) | 🚧 Not yet started |
| `security.py` | `SecurityModule.bas` (user/machine identification) | 🚧 Not yet started |

**Until the modules above are complete, this tool can import and clean
raw data-cart files, detect test cycles, and collect technician channel
configuration — but it cannot yet produce a final pass/fail energy
report.**

## How it works

1. A technician uploads a raw data-cart CSV export through a web form.
2. The file is parsed and cleaned (`io_import.py`): header metadata is
   extracted, open-thermocouple readings are dropped, Unit On / Defrost
   events are identified, and column names are normalized against known
   aliases.
3. The technician assigns each logical sensor role (Watts, FF BW 1,
   RSOC HI, cabinet height, CT value, etc.) to an actual column in the
   uploaded file, via dropdowns populated from the real file contents.
4. Submitting the form validates the assignment (every required role is
   present, every assigned column actually exists, K-factor and CT are
   within valid ranges) and saves it as a configuration file —
   **it does not run the analysis itself.**
5. Running the pipeline (`main.py`) reads that saved configuration,
   detects compressor cycles (`cycles.py`), and — once complete — will
   run stability checks, compute EP1/EP2/ET, and generate a report.

This import/configure step is intentionally decoupled from running the
actual analysis, so a saved configuration can be re-used, re-run after a
fix, or inspected without needing to resubmit the form.

## Requirements

- Python (version pinned in [`.python-version`](./.python-version))
- [`uv`](https://docs.astral.sh/uv/) for dependency management

This project is set up as a standalone application, not an installable
package (`[tool.uv] package = false` in `pyproject.toml`) — there's
nothing to `pip install`, just a set of scripts run with `uv run`.

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

Or via a package manager, if you prefer:
```bash
# macOS (Homebrew)
brew install uv

# pipx
pipx install uv
```

Verify it installed correctly:
```bash
uv --version
```

See the [official uv installation docs](https://docs.astral.sh/uv/getting-started/installation/) for other options.

### 2. Clone and sync the project

```bash
git clone <repo-url>
cd aham-top-fz-doe-4-1
uv sync
```

This creates `.venv` and installs everything listed in
`pyproject.toml` / `uv.lock`. `uv` also manages the correct Python
version automatically based on [`.python-version`](./.python-version)
— you don't need to separately install or manage Python versions
yourself.

> **Note:** if you see a warning about `VIRTUAL_ENV` not matching the
> project environment, either run `deactivate` first and let `uv`
> manage the environment itself, or pass `--active` to the `uv` command
> to use your currently-activated environment instead.