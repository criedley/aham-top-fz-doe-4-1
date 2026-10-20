# tests/test_t2_table.py
from pathlib import Path

import pandas as pd
import pytest

from main import run_analysis
from report import build_report

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"

# Ground truth: the "T2 Data" table of the macro-generated report for the 7272 file.
EXPECTED = {
    "A Cycle":  dict(ff=(37.32, 38.48, 37.31), fz=(-3.17, 0.98, 0.6),   volts=(115.24, 115.18, 115.31),
                     lsoc=90.47, rsoc=90.31, grad=(0.052, 0.113), duration=87.67, on=55.52, off=32.15,
                     runtime=0.6333, wh=92,  start="01/28 14:54", end="01/28 16:22"),
    "Precool":  dict(ff=(37.26, 38.7, 37.58),  fz=(-1.79, 1.9, 1.62),   volts=(115.23, 115.18, 115.3),
                     lsoc=90.42, rsoc=90.25, grad=(0.055, 0.108), duration=76.72, on=45.13, off=31.58,
                     runtime=0.5883, wh=78,  start="01/28 16:22", end="01/28 17:39"),
    "Defrost":  dict(ff=(40.12, 39.06, 37.96), fz=(2.96, 0.31, -0.33),  volts=(115.16, 115.11, 115.32),
                     lsoc=90.42, rsoc=90.28, grad=(0.052, 0.111), duration=16.15, on=15.85, off=0.3,
                     runtime=0.9814, wh=82,  start="01/28 17:39", end="01/28 17:55"),
    "Recovery": dict(ff=(36.57, 38.83, 37.78), fz=(5.19, 7.16, 5.65),   volts=(115.2, 115.08, 115.33),
                     lsoc=90.41, rsoc=90.3,  grad=(0.053, 0.2),   duration=121.32, on=111.3, off=10.02,
                     runtime=0.9174, wh=194, start="01/28 17:55", end="01/28 19:56"),
    "Post 1":   dict(ff=(37.09, 38.32, 37.12), fz=(-3.34, 0.66, 0.27),  volts=(115.26, 115.19, 115.35),
                     lsoc=90.43, rsoc=90.27, grad=(0.059, 0.124), duration=89.33, on=55.75, off=33.58,
                     runtime=0.6241, wh=93,  start="01/28 19:56", end="01/28 21:25"),
    "B Cycle":  dict(ff=(37.21, 38.4, 37.21),  fz=(-3.33, 0.91, 0.51),  volts=(115.26, 115.2, 115.34),
                     lsoc=90.39, rsoc=90.22, grad=(0.053, 0.11),  duration=86.48, on=55.27, off=31.22,
                     runtime=0.639,  wh=92,  start="01/28 21:25", end="01/28 22:52"),
}


@pytest.fixture
def t2(channel_map, real_test_config):
    analysis = run_analysis(FIXTURE, {**channel_map, "Volts": "Volts"}, real_test_config)
    return build_report(analysis, real_test_config).t2_table.set_index("label")


def test_t2_rows_are_in_macro_order(t2):
    assert t2.index.tolist() == list(EXPECTED)


# Defrost LSOC is an exact tie (90.415): Excel ROUND gives 90.42, Python round gives 90.41.
KNOWN_ROUNDING_TIES = {("Defrost", "lsoc_avg")}

SHIFTED = {"Defrost", "Recovery"}   # boundary is 2 s later than the macro's (see test below)

@pytest.mark.parametrize("label", list(EXPECTED))
def test_t2_row_matches_macro(t2, label):
    exp, row = EXPECTED[label], t2.loc[label]
    assert [row["ff_bw1"], row["ff_bw2"], row["ff_bw3"]] == pytest.approx(exp["ff"], abs=0.0051)
    assert [row["fz_bw1"], row["fz_bw2"], row["fz_bw3"]] == pytest.approx(exp["fz"], abs=0.0051)
    assert [row["volts_avg"], row["volts_min"], row["volts_max"]] == pytest.approx(exp["volts"], abs=0.0051)

    lsoc_tol = 0.0101 if (label, "lsoc_avg") in KNOWN_ROUNDING_TIES else 0.0051
    
    assert row["lsoc_avg"] == pytest.approx(exp["lsoc"], abs=lsoc_tol)
    assert row["rsoc_avg"] == pytest.approx(exp["rsoc"], abs=0.0051)
    assert [row["l_grad"], row["r_grad"]] == pytest.approx(exp["grad"], abs=0.0005)

    tol, rtol = (0.05, 0.0005) if label in SHIFTED else (0.0051, 0.00006)

    assert row["duration_min"] == pytest.approx(exp["duration"], abs=tol)
    assert row["on_min"] == pytest.approx(exp["on"], abs=tol)
    assert row["off_min"] == pytest.approx(exp["off"], abs=tol)
    assert row["runtime"] == pytest.approx(exp["runtime"], abs=rtol)
    assert row["wh"] == pytest.approx(exp["wh"], abs=0.5)
    assert row["start"].strftime("%m/%d %H:%M") == exp["start"]
    assert row["end"].strftime("%m/%d %H:%M") == exp["end"]
    assert row["fz_bw4"] == row["fz_bw5"] == row["ice_bw"] == 0     # NA channels print 0, as in the macro


def test_defrost_recovery_boundary_is_the_real_unit_off_event(t2):
    # Macro: 17:55:14. Its import step rewrites that defrost-off row to Unit On = 0 and
    # drops the 17:55:16 row. The port keeps the actual compressor-off at 17:55:16.
    assert t2.loc["Defrost", "end"] == pd.Timestamp("2026-01-28 17:55:16")
    assert t2.loc["Recovery", "start"] == pd.Timestamp("2026-01-28 17:55:16")