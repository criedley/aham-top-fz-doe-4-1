# tests/conftest.py
from pathlib import Path

import pytest
import pandas as pd

from models import ImportResult, ChannelConfig, TestConfig, CycleData
from io_import import load_datacart_csv
from setup import build_channel_config
from cycles import detect_cycles
from channel_defaults import ALIAS_MAP, CHANNEL_MAP

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample_raw.csv"

@pytest.fixture
def alias_map():
    return ALIAS_MAP

@pytest.fixture
def channel_map():
    return CHANNEL_MAP

@pytest.fixture
def real_test_config():
    return TestConfig(height_in=68.5, height_adjusted_in=78.5, ct=10.0,
                      fz_only=False, fz_only_kfactor=1.0, adjusted_start=False)

@pytest.fixture
def real_pipeline(alias_map, channel_map, real_test_config):
    imp = load_datacart_csv(FIXTURE_PATH, alias_map)
    config = build_channel_config(imp, channel_map)
    cycle_data = detect_cycles(imp, config, real_test_config)
    return imp, config, cycle_data

@pytest.fixture(scope="session")
def real_import_result():
    return load_datacart_csv(FIXTURE_PATH, ALIAS_MAP)

@pytest.fixture
def sample_import_result():
    columns = ["Watts", "Watt-Hours", "Unit On", "Defrost On", "FF BW 1"]
    df = pd.DataFrame({c: [1, 2, 3] for c in columns})
    return ImportResult(
        raw_df=df, events_df=df.iloc[0:0], test_info={},
        watts_col="Watts", column_map={c: c for c in columns},
    )

@pytest.fixture
def channel_config():
    return ChannelConfig(
        ff_bw1_col="FF BW 1", ff_bw2_col="FF BW 2", ff_bw3_col="FF BW 3",
        fz_bw1_col="FZ BW 1", fz_bw2_col="FZ BW 2", fz_bw3_col="FZ BW 3",
        fz_bw4_col=None, fz_bw5_col=None, ice_bw_col=None, volts_col=None,
        rsoc_col="RSOC Mid", lsoc_col="LSOC Mid",
        rsoc_hi_col="RSOC HI", rsoc_lo_col="RSOC LO",
        lsoc_hi_col="LSOC HI", lsoc_lo_col="LSOC LO",
        unit_on_col="Event 1 Unit On", defrost_on_col="Event 2 Defrost",
        watts_col="Watts", watt_hours_col="Watt-Hours",
    )

@pytest.fixture
def test_config():
    return TestConfig(
        height_in=55.0, height_adjusted_in=65.0, ct=10.0,
        fz_only=False, fz_only_kfactor=1.0, adjusted_start=False,
    )

@pytest.fixture
def test_config_fz_only():
    return TestConfig(
        height_in=55.0, height_adjusted_in=65.0, ct=10.0,
        fz_only=True, fz_only_kfactor=0.85, adjusted_start=False,
    )