# src/webapp.py
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import yaml
from fastapi import FastAPI, Request, UploadFile
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from starlette.datastructures import FormData

from channel_defaults import ALIAS_MAP
from io_import import load_datacart_csv
from main import run_analysis
from models import TestConfig
from report import build_report
from setup import OPTIONAL_ROLES, REQUIRED_ROLES

app = FastAPI()
templates = Jinja2Templates(directory=str(Path(__file__).resolve().parent.parent / "templates"))

DATA_DIR = Path("data")
CONFIG_PATH = Path("config/channel_map.yaml")

# AHAMCalc.bas: hght = <technician's Unit Height> + 10  ("10 in. adjustment for t/c offset")
TC_OFFSET_IN = 10.0

SEGMENT_LABELS = {"A": "A Cycle", "Precool": "Precool", "Defrost": "Defrost",
                  "Recovery": "Recovery", "Post": "Post", "B": "B Cycle"}


def _as_str(value: object) -> str | None:
    """Plain-text form fields should never be UploadFile objects."""
    return value if isinstance(value, str) else None


def _form_float(form: FormData, name: str, label: str, default: float | None = None) -> float:
    raw = _as_str(form.get(name))
    if raw is None or raw.strip() == "":
        if default is None:
            raise ValueError(f"{label} is required")
        return default
    try:
        return float(raw)
    except ValueError:
        raise ValueError(f"{label} must be a number, got '{raw}'") from None


def _build_test_config(form: FormData) -> TestConfig:
    height_in = _form_float(form, "height_in", "Unit height")
    adjusted_start = form.get("adjusted_start") == "on"
    time_to_start = None
    if adjusted_start:
        raw = _as_str(form.get("time_to_start"))
        if not raw:
            raise ValueError("A start time is required when 'Adjust Start Time' is checked")
        try:
            time_to_start = datetime.fromisoformat(raw)
        except ValueError:
            raise ValueError(f"Start time '{raw}' is not a valid date and time") from None

    return TestConfig(
        height_in=height_in,
        height_adjusted_in=height_in + TC_OFFSET_IN,
        ct=_form_float(form, "ct", "CT"),
        fz_only=form.get("fz_only") == "on",
        fz_only_kfactor=_form_float(form, "fz_only_kfactor", "K-factor", default=1.0),
        adjusted_start=adjusted_start,
        time_to_start=time_to_start,
    )


def _resolve_upload(raw: str | None) -> Path | None:
    """Only accept files that live inside DATA_DIR."""
    if not raw:
        return None
    path = Path(raw).resolve()
    if path.is_file() and path.is_relative_to(DATA_DIR.resolve()):
        return path
    return None


def _technician_form(request: Request, csv_path: Path, error: str | None = None):
    import_result = load_datacart_csv(csv_path, ALIAS_MAP)
    context: dict[str, object] = {
        "columns": list(import_result.raw_df.columns),
        "required_roles": REQUIRED_ROLES,
        "optional_roles": OPTIONAL_ROLES,
        "csv_path": str(csv_path),
    }
    if error:
        context["error"] = error
    return templates.TemplateResponse(request, "technician_form.html", context)


@app.get("/", response_class=HTMLResponse)
def show_upload_form(request: Request):
    return templates.TemplateResponse(request, "upload.html", {})


@app.post("/configure", response_class=HTMLResponse)
async def show_channel_form(request: Request, file: UploadFile):
    if not file.filename:
        return templates.TemplateResponse(request, "upload.html", {
            "error": "No file was selected. Please choose a CSV file to upload."
        })

    DATA_DIR.mkdir(exist_ok=True)
    csv_path = DATA_DIR / Path(file.filename).name      # drop any directory components
    csv_path.write_bytes(await file.read())

    try:
        return _technician_form(request, csv_path)
    except Exception as e:
        return templates.TemplateResponse(
            request, "upload.html", {"error": f"Could not read file: {e}"}
        )


@app.post("/submit", response_class=HTMLResponse)
async def submit_form(request: Request):
    form = await request.form()

    csv_path = _resolve_upload(_as_str(form.get("csv_path")))
    if csv_path is None:
        return templates.TemplateResponse(request, "upload.html", {
            "error": "Missing or invalid file reference. Please upload the file again."
        })

    channel_map = {role: value for role in REQUIRED_ROLES + OPTIONAL_ROLES
                   if (value := _as_str(form.get(role))) is not None}

    try:
        test_config = _build_test_config(form)
        analysis = run_analysis(csv_path, channel_map, test_config)
        report = build_report(analysis, test_config)
    except (ValueError, ImportError, NotImplementedError) as e:
        return _technician_form(request, csv_path, error=str(e))

    CONFIG_PATH.parent.mkdir(exist_ok=True)
    CONFIG_PATH.write_text(yaml.dump({
        "csv_path": str(csv_path),
        "channel_map": channel_map,
        "test_config": asdict(test_config),
    }))

    segment_rows = [
        {"label": SEGMENT_LABELS[name], "start": seg.start, "end": seg.end, "wh": seg.energy_wh}
        for name, seg in report.t2_segments.items()
    ]
    return templates.TemplateResponse(request, "report.html", {
        "report": report,
        "cycle_rows": report.cycle_table.to_dict("records"),
        "segment_rows": segment_rows,
        "filename": csv_path.name,
    })