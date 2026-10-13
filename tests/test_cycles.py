# tests/test_cycles.py
import math
import pytest
import pandas as pd

from cycles import _compute_cycle_stats, CycleBoundaries
from models import ChannelConfig, TestConfig


def make_stats_raw_df(n_rows: int = 300) -> pd.DataFrame:
    """
    FF BW 1 varies by cycle (38.0 / 40.0 / 42.0 across three 100-row
    blocks) specifically so per-cycle averaging can be distinguished from
    a global average. Every other channel is constant, so its expected
    average is just that constant, computable without any arithmetic.
    """
    timestamps = pd.date_range("2026-01-27 00:00:00", periods=n_rows, freq="min")
    ff_bw1 = [38.0] * 100 + [40.0] * 100 + [42.0] * 100
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
      total_min   = 100 - 1 = 99 min   (per the "-1 min" adjustment in the formula)
      duty_cycle  = 50 / 99
    """
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    expected = 50 / 99
    assert result.loc[0, "duty_cycle"] == pytest.approx(expected, rel=1e-6)

def test_last_cycle_duty_cycle_uses_last_scan_time(channel_config, test_config):
    """
    Cycle 2 (the last cycle, index 2 of 3): on at row 200, off at row 250,
    and since there's no next cycle, the formula should use last_scan_time
    (row 299) in place of a next cycle's on-time:
      on_duration = 50 min
      total_min   = (299 - 200) - 1 = 98 min
      duty_cycle  = 50 / 98
    """
    raw_df = make_stats_raw_df()
    boundaries = make_three_cycle_boundaries(raw_df)

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)

    expected = 50 / 98
    assert result.loc[2, "duty_cycle"] == pytest.approx(expected, rel=1e-6)

def test_duty_cycle_is_nan_when_total_window_is_not_positive(channel_config, test_config):
    """
    Edge case: if a cycle's 'on' time and the next cycle's 'on' time are
    less than a minute apart, total_min <= 0 and the formula should not
    raise a ZeroDivisionError -- it should produce NaN (our deliberate
    deviation from the VBA, which did not guard against this at all).
    """
    raw_df = make_stats_raw_df(n_rows=10)
    ts = raw_df["Timestamp"]
    boundaries = CycleBoundaries(
        time_on=[ts.iloc[0], ts.iloc[0]],     # zero gap between cycles -> total_min <= 0
        time_off=[ts.iloc[1], ts.iloc[2]],
        index_on=[0, 0],
        index_off=[1, 2],
        watt_hours_at_on=[0, 0],
        cyclic=True,
        last_scan_index=9,
        last_scan_time=ts.iloc[9],
    )

    result = _compute_cycle_stats(raw_df, boundaries, channel_config, test_config)
    assert pd.isna(result.loc[0, "duty_cycle"])


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