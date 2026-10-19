"""
tests/test_checks.py
Tests for checks.py, covering:
  1. Contract tests for the simple aggregation functions (Gradientcheck/AmbientCheck equivalents)
  2. Failure-path tests for every raise condition in the stability search
  3. Ground-truth-style tests using synthetic data with KNOWN stability outcomes
     (perfectly flat = must be stable, linearly drifting = must never be stable)
  4. Explicit skips for behavior we cannot yet validate against a real file
"""
import pytest
import pandas as pd

from checks import (
    check_gradient, check_ambient, check_stability,
    _grow_stability_window, _find_defrost_before, _InsufficientDataError,
)
from tests.factories import make_perfectly_stable_raw_df, make_never_stable_raw_df, make_cycle_data


# ---------------------------------------------------------------------------
# 1. Contract tests: check_gradient / check_ambient
# ---------------------------------------------------------------------------

def test_gradient_check_computes_expected_value(channel_config, test_config):
    raw_df = make_perfectly_stable_raw_df(200)
    result = check_gradient(raw_df, channel_config, start=0, end=100, test_config=test_config)
    # RSOC HI=90.5, RSOC LO=90.0 -> diff 0.5, height_adjusted_in=65 -> 65/12 ft
    expected = (90.5 - 90.0) / (65.0 / 12)
    assert result.rs_grad == pytest.approx(expected)
    assert result.ls_grad == pytest.approx(expected)

def test_gradient_check_zero_when_hi_equals_lo(channel_config, test_config):
    raw_df = make_perfectly_stable_raw_df(50)
    raw_df["RSOC LO"] = raw_df["RSOC HI"]  # no gradient at all
    result = check_gradient(raw_df, channel_config, start=0, end=49, test_config=test_config)
    assert result.rs_grad == pytest.approx(0.0)

def test_ambient_check_returns_min_and_max(channel_config):
    raw_df = make_perfectly_stable_raw_df(50)
    raw_df.loc[10, "RSOC Mid"] = 99.0   # inject a known outlier
    raw_df.loc[20, "RSOC Mid"] = 80.0
    result = check_ambient(raw_df, channel_config, start=0, end=49)
    assert result.rsoc_max == pytest.approx(99.0)
    assert result.rsoc_min == pytest.approx(80.0)


# ---------------------------------------------------------------------------
# 2. _grow_stability_window: contract + every failure branch
# ---------------------------------------------------------------------------

def test_grow_stability_window_finds_valid_j_k_l_m():
    # cycles spaced 60 min apart -> plenty of room for 120/180/120 min gaps
    index_on = list(range(0, 3000, 60))
    j, k, l, m = _grow_stability_window(index_on, j=0)
    assert index_on[k] - index_on[j] >= 120
    assert index_on[l] - index_on[k] >= 180
    assert index_on[m] - index_on[l] >= 120

def test_grow_stability_window_raises_when_k_search_exhausted():
    index_on = [0, 10, 20]  # far too few/close cycles to ever reach a 120-min gap
    with pytest.raises(_InsufficientDataError, match="K search exhausted"):
        _grow_stability_window(index_on, j=0)

def test_grow_stability_window_raises_when_l_search_exhausted():
    # enough cycles to satisfy K, not enough to satisfy L
    index_on = [0, 60, 130, 140, 150]
    with pytest.raises(_InsufficientDataError, match="L search exhausted"):
        _grow_stability_window(index_on, j=0)

def test_grow_stability_window_raises_when_m_search_exhausted():
    # enough cycles to satisfy K and L, not enough for M
    index_on = [0, 60, 130, 320, 330, 340]
    with pytest.raises(_InsufficientDataError, match="M search exhausted"):
        _grow_stability_window(index_on, j=0)


# ---------------------------------------------------------------------------
# 3. _find_defrost_before
# ---------------------------------------------------------------------------

def test_find_defrost_before_detects_defrost_in_range():
    # Defrost injected exactly at the boundary start of cycle index 2
    # (index_on[2] == 200), so it's unambiguous which cycle "contains" it.
    raw_df = make_perfectly_stable_raw_df(500)
    raw_df.loc[200, "Event 2 Defrost"] = 1   # <- exactly at index_on[2]
    cycle_data = make_cycle_data(index_on=[0, 100, 200, 300, 400])

    result = _find_defrost_before(raw_df, cycle_data, start_cycle=0, end_cycle=3)

    # Defrost at row 200 falls in cycle xx=2's range [200, 300) -> expect xx+1 = 3
    assert result == 3

def test_find_defrost_before_excludes_defrost_outside_checked_range():
    raw_df = make_perfectly_stable_raw_df(500)
    raw_df.loc[300, "Event 2 Defrost"] = 1   # exactly at index_on[3], outside end_cycle=3's range
    cycle_data = make_cycle_data(index_on=[0, 100, 200, 300, 400])

    result = _find_defrost_before(raw_df, cycle_data, start_cycle=0, end_cycle=3)
    assert result is None   # row 300 belongs to cycle xx=3, which is excluded by end_cycle=3


# ---------------------------------------------------------------------------
# 4. check_stability: ground-truth-style tests using synthetic data with
#    KNOWN correct outcomes -- these are the highest-value tests in this
#    file, since they check the actual AHAM stability logic end to end.
# ---------------------------------------------------------------------------

def test_perfectly_flat_data_is_judged_stable_noncyclic(channel_config, test_config):
    raw_df = make_perfectly_stable_raw_df(2000)
    cycle_data = make_cycle_data(index_on=[1], cyclic=False)
    result = check_stability(raw_df, cycle_data, channel_config, test_config)
    assert result.is_stable is True
    assert result.reason is None

def test_linearly_drifting_data_is_never_stable_noncyclic(channel_config, test_config):
    raw_df = make_never_stable_raw_df(2000)
    cycle_data = make_cycle_data(index_on=[1], cyclic=False)
    result = check_stability(raw_df, cycle_data, channel_config, test_config)
    assert result.is_stable is False
    assert result.reason is not None

def test_perfectly_flat_data_is_judged_stable_cyclic(channel_config, test_config):
    raw_df = make_perfectly_stable_raw_df(3000)
    index_on = list(range(0, 2900, 60))  # evenly spaced cycles, 60 min apart
    cycle_data = make_cycle_data(index_on=index_on, cyclic=True)
    result = check_stability(raw_df, cycle_data, channel_config, test_config)
    assert result.is_stable
    assert result.ff_avg == pytest.approx(38.0)
    assert result.fz_avg == pytest.approx(0.0)

def test_fz_only_ignores_ff_drift(channel_config, test_config_fz_only):
    """
    With fz_only=True, only FZ channels should be checked -- so even if FF
    channels are wildly unstable, the test should still pass as long as FZ
    channels are flat. This directly tests _select_channels' effect.
    """
    raw_df = make_perfectly_stable_raw_df(2000)
    raw_df["FF BW 1"] = [i * 0.1 for i in range(2000)]  # FF drifting heavily
    cycle_data = make_cycle_data(index_on=[1], cyclic=False)
    result = check_stability(raw_df, cycle_data, channel_config, test_config_fz_only)
    assert result.is_stable is True

def test_not_fz_only_fails_on_ff_drift(channel_config, test_config):
    """Sanity check / contrast case for the test above: with fz_only=False,
    the same FF drift SHOULD cause instability."""
    raw_df = make_perfectly_stable_raw_df(2000)
    raw_df["FF BW 1"] = [i * 0.1 for i in range(2000)]
    cycle_data = make_cycle_data(index_on=[1], cyclic=False)
    result = check_stability(raw_df, cycle_data, channel_config, test_config)
    assert result.is_stable is False

def test_ff_drift_fails_full_check_but_not_fz_only(channel_config, test_config, test_config_fz_only):
    raw_df = make_perfectly_stable_raw_df(3000)
    raw_df["FF BW 1"] = [38.0 + i * 0.01 for i in range(len(raw_df))]   # FF drifts, FZ flat
    cycle_data = make_cycle_data(index_on=list(range(0, 2900, 60)), cyclic=True)

    assert not check_stability(raw_df, cycle_data, channel_config, test_config).is_stable
    assert check_stability(raw_df, cycle_data, channel_config, test_config_fz_only).is_stable


# ---------------------------------------------------------------------------
# 5. Not yet validated against real data -- explicitly flagged, not silently
#    assumed correct.
# ---------------------------------------------------------------------------

@pytest.mark.skip(reason=(
    "The defrost-interrupts-stability restart path (GoTo ReCheckStability "
    "equivalent) has only been exercised with a single injected defrost "
    "event on synthetic data. It has not been validated against a real "
    "multi-day file where a defrost genuinely occurs mid-search."
))
def test_defrost_during_search_restarts_and_still_finds_stability():
    ...

@pytest.mark.skip(reason=(
    "check_stability's cyclic path has not been validated against the "
    "actual cycle boundaries produced by cycles.py on the real sample "
    "CSV -- only against hand-built synthetic CycleData so far. Needs an "
    "end-to-end test chaining io_import -> cycles -> checks once cycles.py "
    "itself is validated."
))
def test_cyclic_stability_matches_real_sample_file():
    ...
