# src/webapp.py

from pathlib import Path

from fastapi import FastAPI, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
import yaml

from io_import import load_datacart_csv
from setup import build_channel_config, REQUIRED_ROLES, OPTIONAL_ROLES
from models import TestConfig

app = FastAPI()
templates = Jinja2Templates(directory="templates")

DATA_DIR = Path("data")
CONFIG_PATH = Path("config/channel_map.yaml")
ALIAS_MAP: dict[str, list[str]] = {}  # load your real channel_aliases.yaml here

def _as_str(value: object) -> str | None:
    """
    Form fields we expect to be plain text (dropdowns, hidden inputs)
    should never actually be UploadFile objects -- this narrows the type
    for the type checker and guards against a malformed submission.
    """
    return value if isinstance(value, str) else None

@app.get("/", response_class=HTMLResponse)
def show_upload_form(request: Request):
    """Step 1: just ask for the CSV file, no columns to show yet."""
    return templates.TemplateResponse(request, "upload.html", {})

@app.post("/configure", response_class=HTMLResponse)
async def show_channel_form(request: Request, file: UploadFile):
    """Step 2: file is uploaded, now show the channel-assignment form
    populated with real columns from the actual uploaded file."""
    if not file.filename:
        return templates.TemplateResponse(request, "upload.html", {
            "error": "No file was selected. Please choose a CSV file to upload."
        })

    DATA_DIR.mkdir(exist_ok=True)
    csv_path = DATA_DIR / file.filename
    csv_path.write_bytes(await file.read())

    try:
        import_result = load_datacart_csv(csv_path, ALIAS_MAP)
    except Exception as e:
        return templates.TemplateResponse(
            request, "upload.html", {"error": f"Could not read file: {e}"}
        )

    return templates.TemplateResponse(request, "technician_form.html", {
        "columns": list(import_result.raw_df.columns),
        "required_roles": REQUIRED_ROLES,
        "optional_roles": OPTIONAL_ROLES,
        "csv_path": str(csv_path),
    })

@app.post("/submit", response_class=HTMLResponse)
async def submit_form(request: Request):
    form = await request.form()

    raw_csv_path = form.get("csv_path")
    if not isinstance(raw_csv_path, str) or not raw_csv_path:
        return templates.TemplateResponse(request, "upload.html", {
            "error": "Missing or invalid file reference. Please start over."
        })
    csv_path = Path(raw_csv_path)
    if not csv_path.exists():
        return templates.TemplateResponse(request, "upload.html", {
            "error": f"Uploaded file no longer found at {csv_path}. Please upload again."
        })

    channel_map: dict[str, str] = {}
    for role in REQUIRED_ROLES + OPTIONAL_ROLES:
        value = _as_str(form.get(role))
        if value is not None:
            channel_map[role] = value

    try:
        import_result = load_datacart_csv(csv_path, ALIAS_MAP)
        build_channel_config(import_result, channel_map)   # now receives a real dict[str, str]

        test_config_fields = {
            "height_in": float(_as_str(form.get("height_in")) or 0),
            "ct": float(_as_str(form.get("ct")) or 0),
            "fz_only": form.get("fz_only") == "on",
            "fz_only_kfactor": float(_as_str(form.get("fz_only_kfactor")) or 1.0),
            "adjusted_start": form.get("adjusted_start") == "on",
            "time_to_start": _as_str(form.get("time_to_start")),
        }
        TestConfig(**test_config_fields)

    except (ValueError, FileNotFoundError) as e:
        import_result = load_datacart_csv(csv_path, ALIAS_MAP)
        return templates.TemplateResponse(request, "technician_form.html", {
            "columns": list(import_result.raw_df.columns),
            "required_roles": REQUIRED_ROLES,
            "optional_roles": OPTIONAL_ROLES,
            "csv_path": str(csv_path),
            "error": str(e),
        })

    CONFIG_PATH.parent.mkdir(exist_ok=True)
    CONFIG_PATH.write_text(yaml.dump({
        "csv_path": str(csv_path),
        "channel_map": channel_map,
        "test_config": test_config_fields,
    }))

    return RedirectResponse(url="/done", status_code=303)


@app.get("/done", response_class=HTMLResponse)
def done():
    return """
    <h2>Configuration saved.</h2>
    <p>Run <code>python main.py</code> to process this test.</p>
    """