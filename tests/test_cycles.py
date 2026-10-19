# tests/test_cycles.py
import math
import pytest
import pandas as pd

from cycles import CycleBoundaries, _compute_cycle_stats, _find_boundaries_from_events
from models import ChannelConfig, TestConfig

def test_defrost_only_change_does_not_create_phantom_cycle_boundary(channel_config, test_config):
    """
    Regression test for a real bug found validating against the 7272
    sample file at ~17:39-17:55 on 28 Jan 2026: a defrost-only transition
    row can have "Event 1 Unit On" sitting at value 1 even though Unit On
    itself did NOT change on that row. Filtering event rows purely by
    their Unit On *value* (rather than whether Unit On changed) incorrectly
    treated that row as a fresh "on" event, pairing it with the next real
    off-event 2 seconds later and fabricating a bogus near-instantaneous
    cycle (duty_cycle ~0.0037 in the real output).

    Real sequence, condensed to the event-only rows (matching the actual
    CSV):
        t=0        Unit On 1->0              (real compressor off)
        t=18s      Defrost 0->1, Unit On=0   (defrost heater starting; Unit On unchanged)
        t=18s      Unit On 0->1, Defrost=1   (same timestamp: heater draws through Unit On signal)
        t=16m      Defrost 1->0, Unit On=1   (defrost heater off -- Unit On UNCHANGED, still 1)
        t=16m2s    Unit On 1->0, Defrost=0   (real compressor/unit off)

    Correct result: exactly ONE on/off pair (on at t=18s, off at t=16m2s).
    The row at t=16m must NOT be treated as an "on" event just because
    its Unit On value happens to be 1.
    """
    base = pd.Timestamp("2026-01-28 17:39:00")
    events_df = pd.DataFrame({
        "Timestamp": [
            base,
            base + pd.Timedelta(seconds=18),
            base + pd.Timedelta(seconds=18),
            base + pd.Timedelta(minutes=16),
            base + pd.Timedelta(minutes=16, seconds=2),
        ],
        "Event 1 Unit On": [0, 0, 1, 1, 0],
        "Event 2 Defrost": [0, 1, 1, 0, 0],
        "Watt-Hours": [1974, 1974, 1974, 2056, 2056],
    })

    raw_df = pd.DataFrame({
        "Timestamp": pd.date_range(base - pd.Timedelta(minutes=1), periods=2000, freq="s"),
    })

    boundaries = _find_boundaries_from_events(events_df, raw_df, channel_config, test_config)

    assert len(boundaries.time_on) == 1, (
        f"Expected exactly 1 cycle, got {len(boundaries.time_on)} -- "
        f"a Defrost-only transition likely got mistaken for a Unit On event"
    )
    assert len(boundaries.time_off) == 1

    expected_on = base + pd.Timedelta(seconds=18)
    expected_off = base + pd.Timedelta(minutes=16, seconds=2)
    assert boundaries.time_on[0] == expected_on
    assert boundaries.time_off[0] == expected_off
    # assert boundaries.cyclic is True

def test_rapid_genuine_unit_on_transitions_are_still_captured(channel_config, test_config):
    """
    Contrast case for the fix above: confirms that filtering by 'did Unit
    On actually change' doesn't accidentally suppress legitimate rapid
    on/off/on toggling -- only rows where Unit On's value is unchanged
    should be excluded.
    """
    base = pd.Timestamp("2026-01-28 12:00:00")
    events_df = pd.DataFrame({
        "Timestamp": [
            base,
            base + pd.Timedelta(seconds=5),
            base + pd.Timedelta(seconds=10),
        ],
        "Event 1 Unit On": [1, 0, 1],
        "Event 2 Defrost": [0, 0, 0],
        "Watt-Hours": [100, 100, 101],
    })
    raw_df = pd.DataFrame({
        "Timestamp": pd.date_range(base - pd.Timedelta(seconds=5), periods=100, freq="s"),
    })

    boundaries = _find_boundaries_from_events(events_df, raw_df, channel_config, test_config)

    assert len(boundaries.time_on) == 1
    assert boundaries.time_on[0] == base
    assert boundaries.time_off[0] == base + pd.Timedelta(seconds=5)

def make_stats_raw_df(n_rows: int = 300) -> pd.DataFrame:
    """
    FF BW 1 varies across three roughly-equal blocks (for n_rows=300, this
    is 100/100/100, matching make_three_cycle_boundaries' cycle layout).
    For other n_rows values, the three-block pattern still fills the full
    length -- the exact block sizes only matter for tests that correlate
    this with a specific CycleBoundaries fixture.
    """
    timestamps = pd.date_range("2026-01-27 00:00:00", periods=n_rows, freq="min")

    block = max(1, n_rows // 3)
    ff_bw1 = [38.0] * block + [40.0] * block + [42.0] * block
    if len(ff_bw1) < n_rows:
        ff_bw1 += [42.0] * (n_rows - len(ff_bw1))   # pad any remainder
    else:
        ff_bw1 = ff_bw1[:n_rows]                      # trim if block*3 overshot

    return pd.DataFrame({
        "Timestamp": timestamps,
        "FF BW 1": ff_bw1,
        "FF BW 2": [39.0] * n_rows,
        "FF BW 3": [37.0] * n_rows,
        "FZ BW 1": [0.0] * n_rows,
        "FZ BW 2": [1.0] * n_rows,
        "FZ BW 3": [-1.0] * n_rows,
        "RSOC HI": [91.0] * n_rows, "RSOC LO": [90.0] * n_rows, "RSOC Mid": [90.5] * n_rows,
        "LSOC HI": [91.0] * n_rows, "LSOC LO": [90.0] * n_rows, "LSOC Mid": [90.5] * n_rows,
        "Watt-Hours": list(range(n_rows)),
    })

def make_three_cycle_boundaries(raw_df: pd.DataFrame) -> CycleBoundaries:
    """Three cycles: [0,100), [100,200), [200,300). On=start, Off=start+50."""
    index_on = [0, 100, 200]
    index_off = [50, 150, 250]
    ts = raw_df["Timestamp"]
    return CycleBoundaries(
        time_on=[ts.iloc[i] for i in index_on],
        time_off=[ts.iloc[i] for i in index_off],
        index_on=index_on,
        index_off=index_off,
        watt_hours_at_on=[raw_df["Watt-Hours"].iloc[i] for i in index_on],
        cyclic=True,
        last_scan_index=len(raw_df) - 1,
        last_scan_time=ts.iloc[-1],
    )


# ---------------------------------------------------------------------------
# Contract: averages computed correctly, per-cycle (not globally)
# ---------------------------------------------------------------------------

def test_ff_bw_averages_differ_per_cycle(channel_config, test_config):
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    assert result.loc[0, "ff_bw1"] == pytest.approx(38.0)
    assert result.loc[1, "ff_bw1"] == pytest.approx(40.0)
    assert result.loc[2, "ff_bw1"] == pytest.approx(42.0)

def test_constant_channels_average_to_their_constant(channel_config, test_config):
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    for cycle in range(3):
        assert result.loc[cycle, "ff_bw2"] == pytest.approx(39.0)
        assert result.loc[cycle, "fz_bw1"] == pytest.approx(0.0)
        assert result.loc[cycle, "fz_bw2"] == pytest.approx(1.0)

def test_cycle_num_and_indices_pass_through_from_boundaries(channel_config, test_config):
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    assert result["cycle_num"].tolist() == [1, 2, 3]
    assert result["index_on"].tolist() == [0, 100, 200]
    assert result["index_off"].tolist() == [50, 150, 250]
    assert result["watt_hours_at_on"].tolist() == [0, 100, 200]


# ---------------------------------------------------------------------------
# Gradient calculation
# ---------------------------------------------------------------------------

def test_gradient_matches_hand_computed_value(channel_config, test_config):
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    # RSOC HI=91.0, RSOC LO=90.0 -> diff 1.0; height_adjusted_in=65 -> height_ft = 65/12
    # r_grad = round(1.0 / (65/12), 3) = round(0.184615..., 3) = 0.185
    expected = round(1.0 / (65.0 / 12), 3)
    assert result.loc[0, "r_grad"] == pytest.approx(expected)
    assert result.loc[0, "l_grad"] == pytest.approx(expected)

def test_gradient_is_zero_when_hi_equals_lo(channel_config, test_config):
    raw_df = make_stats_raw_df()
    raw_df["RSOC LO"] = raw_df["RSOC HI"]
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)
    assert result.loc[0, "r_grad"] == pytest.approx(0.0)


# ---------------------------------------------------------------------------
# Optional channels: None in config -> 0.0, configured -> real values
# ---------------------------------------------------------------------------

def test_optional_channels_are_zero_when_not_configured(channel_config, test_config):
    # channel_config fixture already has fz_bw4_col=None, ice_bw_col=None, volts_col=None
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    assert (result["fz_bw4"] == 0.0).all()
    assert (result["fz_bw5"] == 0.0).all()
    assert (result["ice_bw"] == 0.0).all()
    assert (result["volts_avg"] == 0.0).all()
    assert (result["volts_min"] == 0.0).all()
    assert (result["volts_max"] == 0.0).all()

def test_optional_channels_computed_when_configured(channel_config, test_config):
    raw_df = make_stats_raw_df()
    raw_df["FZ BW 4"] = [5.0] * len(raw_df)
    raw_df["Volts"] = [118.0] * 100 + [120.0] * 100 + [122.0] * 100
    boundaries = make_three_cycle_boundaries(raw_df)

    config_with_optional = ChannelConfig(
        **{**channel_config.__dict__, "fz_bw4_col": "FZ BW 4", "volts_col": "Volts"}
    )

    result = _compute_cycle_stats(raw_df, boundaries, config_with_optional, test_config)

    assert result.loc[0, "fz_bw4"] == pytest.approx(5.0)
    assert result.loc[1, "volts_avg"] == pytest.approx(120.0)
    assert result.loc[1, "volts_min"] == pytest.approx(120.0)
    assert result.loc[1, "volts_max"] == pytest.approx(120.0)


# ---------------------------------------------------------------------------
# Duty cycle: hand-traced expected values
# ---------------------------------------------------------------------------

def test_duty_cycle_matches_hand_computed_value(channel_config, test_config):
    """
    Cycle 0: on at row 0, off at row 50, next cycle on at row 100.
    Timestamps are 1 min apart, so:
        on_duration = 50 min
        total_min   = 100 min   (no -1 min adjustment; matches the report's % Runtime)
        duty_cycle  = 50 / 100
    """
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    expected = 50 / 100
    assert result.loc[0, "duty_cycle"] == pytest.approx(expected, rel=1e-6)

def test_last_cycle_duty_cycle_uses_last_scan_time(channel_config, test_config):
    """
    Cycle 2 (the last cycle): on at row 200, off at row 250, and last_scan_time
    is row 299:
        on_duration = 50 min
        total_min   = 299 - 200 = 99 min
        duty_cycle  = 50 / 99
    """
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    expected = 50 / 99
    assert result.loc[2, "duty_cycle"] == pytest.approx(expected, rel=1e-6)

def test_duty_cycle_is_nan_when_total_window_is_not_positive(channel_config, test_config):
    """
    Edge case: cycle 0 and cycle 1 have identical time_on values (zero
    elapsed time between them), so total_min <= 0. Row ranges are still
    valid/non-overlapping so only the duty-cycle math is being exercised,
    not a degenerate window.
    """
    raw_df = make_stats_raw_df(n_rows=10)
    ts = raw_df["Timestamp"]
    boundaries = CycleBoundaries(
        time_on=[ts.iloc[0], ts.iloc[0], ts.iloc[5]],   # cycle 0 and 1 share a timestamp
        time_off=[ts.iloc[1], ts.iloc[3], ts.iloc[8]],
        index_on=[0, 3, 5],
        index_off=[1, 4, 8],
        watt_hours_at_on=[0, 3, 5],
        cyclic=True,
        last_scan_index=9,
        last_scan_time=ts.iloc[9],
    )

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)
    assert pd.isna(result.loc[0, "duty_cycle"])


def test_duty_cycle_never_exceeds_one_on_real_file(real_pipeline):
    _, _, cycle_data = real_pipeline
    duty = cycle_data.cycles["duty_cycle"].dropna()
    assert (duty <= 1.0).all()

    # Cycle 21 (16:53:57 to 17:39:05, just before the defrost): the macro report shows 99.34%.
    cycle_21 = cycle_data.cycles.loc[cycle_data.cycles["cycle_num"] == 21, "duty_cycle"].iloc[0]
    assert cycle_21 == pytest.approx(0.9934, abs=0.001)


# ---------------------------------------------------------------------------
# Failure path
# ---------------------------------------------------------------------------

def test_raises_on_empty_boundaries(channel_config, test_config):
    raw_df = make_stats_raw_df()
    empty_boundaries = CycleBoundaries(
        time_on=[], time_off=[], index_on=[], index_off=[], watt_hours_at_on=[],
        cyclic=True, last_scan_index=len(raw_df) - 1, last_scan_time=raw_df["Timestamp"].iloc[-1],
    )
    with pytest.raises(ValueError, match="No cycle boundaries"):
        _compute_cycle_stats(raw_df, empty_boundaries, channel_config, test_config)