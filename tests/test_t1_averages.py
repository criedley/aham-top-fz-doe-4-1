# tests/test_t1_averages.py
import pandas as pd
import pytest

from checks import check_stability, _t1_window_averages


def test_t1_window_averages_hand_computed(channel_config):
    # Row 0 is outside the window; rows 1-2 are inside (inclusive end).
    cycles = pd.DataFrame({
        "ff_bw1": [100.0, 30.0, 32.0],
        "ff_bw2": [100.0, 40.0, 42.0],
        "ff_bw3": [100.0, 50.0, 52.0],
        "fz_bw1": [9.0, -5.0, -7.0],
        "fz_bw2": [9.0, -3.0, -5.0],
        "fz_bw3": [9.0, -1.0, -3.0],
        "fz_bw4": [0.0, 0.0, 0.0],   # unconfigured channel: must be ignored
        "fz_bw5": [0.0, 0.0, 0.0],
    })
    ff_avg, fz_avg = _t1_window_averages(cycles, start=1, end=2, config=channel_config)
    assert ff_avg == pytest.approx(41.0)   # channel means 31, 41, 51
    assert fz_avg == pytest.approx(-4.0)   # channel means -6, -4, -2 (would be -3 if bw4 leaked in)


def test_stability_averages_match_macro_on_7272_file(real_pipeline, real_test_config):
    imp, config, cycle_data = real_pipeline
    r = check_stability(imp.raw_df, cycle_data, config, real_test_config)

    assert r.is_stable
    assert r.stable_cycle == 6          # cycle 7
    assert r.start_t1_index == 2        # cycle 3
    # Macro report (Book.xlsx): FF Avg 37.64, FZ Avg -0.6
    assert r.ff_avg == pytest.approx(37.64, abs=0.01)
    assert r.fz_avg == pytest.approx(-0.6, abs=0.05)