# tests/test_report.py
from pathlib import Path

import pandas as pd
import pytest

from main import run_analysis
from report import (ambient_ok, build_report, gradient_maxes, gradient_ok,
                    ice_ok, pre_defrost_cycles, optional_columns)

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"


def _cycles(**cols) -> pd.DataFrame:
    return pd.DataFrame(cols)


# --- gradients -------------------------------------------------------------

def test_gradient_maxes_hand_computed_and_last_cycle_excluded():
    # height 5 ft. Lo-Mid = (avg - lo)/5, Mid-Hi = (hi - avg)/5.
    cycles = _cycles(
        lsoc_avg=[90.0, 90.0, 90.0], lsoc_lo_avg=[89.5, 89.0, 80.0], lsoc_hi_avg=[91.0, 91.5, 99.0],
        rsoc_avg=[90.0, 90.0, 90.0], rsoc_lo_avg=[89.0, 89.0, 80.0], rsoc_hi_avg=[91.0, 91.0, 99.0],
    )
    ll, lh, rl, rh = gradient_maxes(cycles, height_ft=5.0)
    assert ll == pytest.approx(0.2)    # cycle 2: (90 - 89)/5 ; cycle 1 gives 0.1
    assert lh == pytest.approx(0.3)    # cycle 2: (91.5 - 90)/5 ; cycle 1 gives 0.2
    assert rl == pytest.approx(0.2)
    assert rh == pytest.approx(0.2)    # the huge last-cycle values are ignored


def test_gradient_maxes_single_cycle_is_zero():
    cycles = _cycles(lsoc_avg=[90.0], lsoc_lo_avg=[89.0], lsoc_hi_avg=[91.0],
                     rsoc_avg=[90.0], rsoc_lo_avg=[89.0], rsoc_hi_avg=[91.0])
    assert gradient_maxes(cycles, 5.0) == (0.0, 0.0, 0.0, 0.0)


# --- flags -----------------------------------------------------------------

def test_ambient_flag_only_considers_cycles_up_to_stable_cycle():
    cycles = _cycles(lsoc_avg=[90.0, 90.0, 95.0], rsoc_avg=[90.0, 90.0, 90.0])
    assert ambient_ok(cycles, stable_cycle=1)
    assert not ambient_ok(cycles, stable_cycle=2)


def test_ambient_flag_fails_on_either_side():
    assert not ambient_ok(_cycles(lsoc_avg=[90.0], rsoc_avg=[91.06]), 0)
    assert not ambient_ok(_cycles(lsoc_avg=[88.9], rsoc_avg=[90.0]), 0)


def test_gradient_flag_uses_limit_and_stable_cycle():
    cycles = _cycles(l_grad=[0.1, 0.1, 3.0], r_grad=[0.1, 0.1, 0.1])
    assert gradient_ok(cycles, stable_cycle=1, limit=2.5)
    assert not gradient_ok(cycles, stable_cycle=2, limit=2.5)


def test_ice_flag_is_none_without_channel_and_checks_32f_with_one(channel_config):
    table = _cycles(ice_bw=[30.0, 33.0])
    assert ice_ok(table, channel_config) is None            # fixture has ice_bw_col=None
    channel_config.ice_bw_col = "ICE BW"
    assert ice_ok(table, channel_config) is False
    assert ice_ok(_cycles(ice_bw=[30.0, 32.0]), channel_config) is True


# --- table slice -----------------------------------------------------------

def test_pre_defrost_cycles_stop_at_defrost_and_leave_room_for_last_cycle():
    base = pd.Timestamp("2026-01-27")
    cycles = _cycles(cycle_num=[1, 2, 3, 4],
                     time_on=[base + pd.Timedelta(hours=h) for h in (0, 1, 2, 3)])
    table = pre_defrost_cycles(cycles, defrost_start=base + pd.Timedelta(hours=2, minutes=30))
    assert table["cycle_num"].tolist() == [1, 2, 3]
    table = pre_defrost_cycles(cycles, defrost_start=base + pd.Timedelta(hours=10))
    assert table["cycle_num"].tolist() == [1, 2, 3]         # capped at len - 1


# --- real file vs. the macro report (Book.xlsx) ----------------------------

@pytest.fixture
def report(channel_map, real_test_config):
    analysis = run_analysis(FIXTURE, channel_map, real_test_config)
    return build_report(analysis, real_test_config)


def test_report_summary_matches_macro(report):
    assert report.test_duration_min == pytest.approx(673.7, abs=0.001)
    assert report.energy_kwh_day == pytest.approx(1.677, abs=0.001)
    assert report.energy_corr_kwh_day == pytest.approx(1.677, abs=0.001)
    assert report.ep1_kwh == pytest.approx(0.458, abs=0.0005)
    assert report.t1_min == pytest.approx(431.8, abs=0.001)
    assert report.ep2_kwh == pytest.approx(0.447, abs=0.0005)
    assert report.t2_min == pytest.approx(303.5, abs=0.001)
    assert report.ct == 10
    assert report.k_factor == 1


def test_report_periods_match_macro(report):
    # T1: the macro shows minutes only, so allow a minute either way.
    assert abs((report.t1_start - pd.Timestamp("2026-01-27 14:55")).total_seconds()) <= 60
    assert abs((report.t1_end - pd.Timestamp("2026-01-27 22:06")).total_seconds()) <= 60
    assert report.t2_start == pd.Timestamp("2026-01-28 16:22:22")
    assert report.t2_end == pd.Timestamp("2026-01-28 21:25:53")


def test_report_averages_and_ambient_match_macro(report):
    assert report.ff_avg == pytest.approx(37.64, abs=0.001)
    assert report.fz_avg == pytest.approx(-0.6, abs=0.001)
    assert report.lsoc_avg == pytest.approx(90.4, abs=0.001)
    assert report.rsoc_avg == pytest.approx(90.26, abs=0.001)


def test_report_gradients_match_macro(report):
    assert report.grad_left_lo_mid == pytest.approx(0.043, abs=0.0005)
    assert report.grad_left_mid_hi == pytest.approx(0.031, abs=0.0005)
    assert report.grad_right_lo_mid == pytest.approx(0.197, abs=0.0005)
    assert report.grad_right_mid_hi == pytest.approx(0.027, abs=0.0005)


def test_report_table_and_flags(report):
    assert report.stable_cycle == 6
    assert report.cycle_table["cycle_num"].iloc[-1] == 21   # last cycle before the defrost
    assert report.ambient_ok and report.gradient_ok
    assert report.ice_ok is None                            # 7272 has no ICE channel


def test_optional_columns_only_for_assigned_channels(channel_config):
    assert optional_columns(channel_config) == ()       # fixture has all four unassigned
    channel_config.volts_col = "Volts"
    channel_config.fz_bw4_col = "FZ BW 4"
    assert [k for k, _ in optional_columns(channel_config)] == [
        "fz_bw4", "volts_avg", "volts_min", "volts_max"]