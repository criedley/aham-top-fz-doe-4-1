# tests/test_io_import.py
import pandas as pd
import pytest

from pathlib import Path
from io_import import load_datacart_csv, _dedupe_events

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"

def test_loads_without_error(alias_map):
    result = load_datacart_csv(FIXTURE, alias_map)
    assert len(result.raw_df) > 0

def test_timestamp_column_parsed_as_datetime(alias_map):
    result = load_datacart_csv(FIXTURE, alias_map)
    assert pd.api.types.is_datetime64_any_dtype(result.raw_df["Timestamp"])

def test_required_columns_present(alias_map):
    result = load_datacart_csv(FIXTURE, alias_map)
    for col in ["Watts", "Watt-Hours", "Event 1 Unit On", "Event 2 Defrost"]:
        assert col in result.raw_df.columns

def test_minimum_scan_count_enforced():
    import pytest
    from io_import import _check_minimum_scans
    tiny_df = pd.DataFrame({"Timestamp": pd.date_range("2024-01-01", periods=10, freq="min")})
    with pytest.raises(ImportError):
        _check_minimum_scans(tiny_df)

def test_events_extracted_match_known_transitions(alias_map):
    """
    Spot-check against the specific Unit-On transitions we manually
    identified in the raw file earlier (e.g. the 08:57 / 23:23 events).
    """
    result = load_datacart_csv(FIXTURE, alias_map)
    on_transitions = result.events_df[result.events_df["Event 1 Unit On"] == 0]
    # confirm at least one event near the known 08:57 transition exists
    assert any(
        abs((ts - pd.Timestamp("2026-01-29 08:57:28")).total_seconds()) < 60
        for ts in on_transitions["Timestamp"]
    )

def test_dedupe_collapses_duplicate_state_despite_differing_watt_hours_and_counters(channel_config=None):
    """
    Regression test for a real bug found while auditing _dedupe_events
    after validating cycles.py against the 7272 sample file: comparing
    ALL non-Timestamp columns (rather than just the Unit On / Defrost /
    Comp state columns) meant two rows were only ever considered
    "duplicates" if byte-for-byte identical -- which essentially never
    happens in real data, since Watt-Hours changes continuously and the
    "Event Count N ..." columns are monotonically-incrementing counters
    that increment on every real event.

    Confirmed on the actual 7272 sample file: 67 events in, 67 events
    out -- i.e. dedup was a no-op on real data, not just a theoretical risk.

    This test builds two consecutive rows with IDENTICAL Unit On / Defrost
    / Comp state, 10 seconds apart (within the 120s window), but with
    DIFFERENT Watt-Hours and Event Count values -- exactly the pattern
    found in real data cart output. Correct behavior: collapse to 1 row.
    """
    base = pd.Timestamp("2026-01-28 12:00:00")
    events_df = pd.DataFrame({
        "Timestamp": [base, base + pd.Timedelta(seconds=10)],
        "Event 1 Unit On": [1, 1],
        "Event 2 Defrost": [0, 0],
        "Event 3 Comp": [0, 0],
        "Watt-Hours": [500, 501],                    # differs row to row
        "Event Count 1 Unit On": [10, 10],
        "Event Count 2 Defrost": [2, 2],
        "Event Count 3 Comp": [0, 1],                 # differs row to row
    })

    result = _dedupe_events(events_df)

    assert len(result) == 1, (
        f"Expected duplicate state rows to collapse to 1, got {len(result)} -- "
        f"_dedupe_events is likely comparing Watt-Hours/Event Count columns "
        f"instead of only the state columns (Unit On / Defrost / Comp)"
    )
    assert result.iloc[0]["Timestamp"] == base


def test_dedupe_keeps_genuinely_different_states_even_if_close_in_time():
    """
    Contrast case: two rows 10 seconds apart with genuinely DIFFERENT
    Unit On state must both be kept, regardless of how close in time they
    are -- the 120s window only collapses true duplicates, not every
    rapid transition.
    """
    base = pd.Timestamp("2026-01-28 12:00:00")
    events_df = pd.DataFrame({
        "Timestamp": [base, base + pd.Timedelta(seconds=10)],
        "Event 1 Unit On": [1, 0],   # genuinely different
        "Event 2 Defrost": [0, 0],
        "Event 3 Comp": [0, 0],
        "Watt-Hours": [500, 501],
        "Event Count 1 Unit On": [10, 11],
        "Event Count 2 Defrost": [2, 2],
        "Event Count 3 Comp": [0, 0],
    })

    result = _dedupe_events(events_df)

    assert len(result) == 2


def test_dedupe_keeps_duplicate_state_outside_the_time_window():
    """
    Contrast case: identical state, but more than 120 seconds apart --
    should NOT be collapsed, since the window itself is still meant to
    matter, not just the state comparison.
    """
    base = pd.Timestamp("2026-01-28 12:00:00")
    events_df = pd.DataFrame({
        "Timestamp": [base, base + pd.Timedelta(seconds=130)],
        "Event 1 Unit On": [1, 1],
        "Event 2 Defrost": [0, 0],
        "Event 3 Comp": [0, 0],
        "Watt-Hours": [500, 505],
        "Event Count 1 Unit On": [10, 10],
        "Event Count 2 Defrost": [2, 2],
        "Event Count 3 Comp": [0, 0],
    })

    result = _dedupe_events(events_df)

    assert len(result) == 2