# tests/test_webapp.py
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import webapp
from channel_defaults import CHANNEL_MAP

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(webapp, "DATA_DIR", tmp_path / "data")
    monkeypatch.setattr(webapp, "CONFIG_PATH", tmp_path / "config" / "channel_map.yaml")
    return TestClient(webapp.app)


def upload(client) -> str:
    r = client.post("/configure",
                    files={"file": (FIXTURE.name, FIXTURE.read_bytes(), "text/csv")})
    assert r.status_code == 200
    return str(webapp.DATA_DIR / FIXTURE.name)


def form_data(csv_path: str, **overrides) -> dict:
    data = dict(CHANNEL_MAP)                       # role -> column for the 16 required roles
    data.update(csv_path=csv_path, height_in="71", ct="10", fz_only_kfactor="1")
    data.update(overrides)
    return {k: v for k, v in data.items() if v is not None}


def test_configure_lists_columns_from_the_uploaded_file(client):
    r = client.post("/configure",
                    files={"file": (FIXTURE.name, FIXTURE.read_bytes(), "text/csv")})
    assert "FF BW 1" in r.text                     # alias-normalized column name


def test_submit_runs_pipeline_and_shows_macro_numbers(client):
    csv_path = upload(client)
    r = client.post("/submit", data=form_data(csv_path))
    assert r.status_code == 200
    assert "AHAM Report" in r.text
    for expected in ("1.677", "303.5", "673.7", "0.458", "0.447", "0.043", "0.197", "Recovery", "121.32", "98.14%"):
        assert expected in r.text, expected


def test_submit_saves_the_configuration(client):
    csv_path = upload(client)
    client.post("/submit", data=form_data(csv_path))
    assert webapp.CONFIG_PATH.exists()


def test_missing_height_shows_error_on_form(client):
    csv_path = upload(client)
    r = client.post("/submit", data=form_data(csv_path, height_in=None))
    assert "AHAM Report" not in r.text
    assert "Unit height is required" in r.text     # needs the {% if error %} block in the template


def test_zero_height_is_rejected(client):
    csv_path = upload(client)
    r = client.post("/submit", data=form_data(csv_path, height_in="0"))
    assert "AHAM Report" not in r.text
    assert "greater than zero" in r.text


def test_csv_path_outside_data_dir_is_rejected(client):
    upload(client)
    r = client.post("/submit", data=form_data("/etc/passwd"))
    assert r.status_code == 200
    assert "AHAM Report" not in r.text


def test_upload_page_renders(client):
    r = client.get("/")
    assert r.status_code == 200
    assert 'name="file"' in r.text


def test_unreadable_file_shows_error_on_upload_page(client):
    r = client.post("/configure",
                    files={"file": ("bad.csv", b"not a data-cart file", "text/csv")})
    assert r.status_code == 200
    assert "Could not read file" in r.text


def test_configure_form_has_hidden_path_and_default_selections(client):
    r = client.post("/configure",
                    files={"file": (FIXTURE.name, FIXTURE.read_bytes(), "text/csv")})
    assert 'name="csv_path"' in r.text
    assert 'value="Watts" selected' in r.text