# tests/test_io_import.py
from pathlib import Path
import pandas as pd
from io_import import load_datacart_csv

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"
ALIAS_MAP = {
    "Watts": ["Watts"],
    "Watt-Hours": ["Watt-Hours"],
    "Event 1 Unit On": ["Event 1 Unit On"],
    "Event 2 Defrost": ["Event 2 Defrost"],
    "Event 3 Comp": ["Event 3 Comp"],
    "FF BW 1": ["FF BW1"],
    "FF BW 2": ["FF BW2"],
    "FF BW 3": ["FF BW3"],
    "FZ BW 1": ["FZ BW1"],
    "FZ BW 2": ["FZ BW2"],
    "FZ BW 3": ["FZ BW3"],
    "RSOC HI": ["RSOC UP"],
    "RSOC Mid": ["RSOC MD"],
    "RSOC LO": ["RSOC LO"],
    "LSOC HI": ["LSOC UP"],
    "LSOC Mid": ["LSOC MD"],
    "LSOC LO": ["LSOC LO"],
}  # fill in your real channel_aliases.yaml equivalent

def test_loads_without_error():
    result = load_datacart_csv(FIXTURE, ALIAS_MAP)
    assert len(result.raw_df) > 0

def test_timestamp_column_parsed_as_datetime():
    result = load_datacart_csv(FIXTURE, ALIAS_MAP)
    assert pd.api.types.is_datetime64_any_dtype(result.raw_df["Timestamp"])

def test_required_columns_present():
    result = load_datacart_csv(FIXTURE, ALIAS_MAP)
    for col in ["Watts", "Watt-Hours", "Event 1 Unit On", "Event 2 Defrost"]:
        assert col in result.raw_df.columns

def test_minimum_scan_count_enforced():
    import pytest
    from io_import import _check_minimum_scans
    tiny_df = pd.DataFrame({"Timestamp": pd.date_range("2024-01-01", periods=10, freq="min")})
    with pytest.raises(ImportError):
        _check_minimum_scans(tiny_df)

def test_events_extracted_match_known_transitions():
    """
    Spot-check against the specific Unit-On transitions we manually
    identified in the raw file earlier (e.g. the 08:57 / 23:23 events).
    """
    result = load_datacart_csv(FIXTURE, ALIAS_MAP)
    on_transitions = result.events_df[result.events_df["Event 1 Unit On"] == 0]
    # confirm at least one event near the known 08:57 transition exists
    assert any(
        abs((ts - pd.Timestamp("2026-01-29 08:57:28")).total_seconds()) < 60
        for ts in on_transitions["Timestamp"]
    )