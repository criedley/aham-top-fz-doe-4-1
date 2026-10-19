# tests/test_main.py
from pathlib import Path

import pytest

from main import run_analysis

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"


def test_end_to_end_matches_macro_on_7272_file(channel_map, real_test_config):
    r = run_analysis(FIXTURE, channel_map, real_test_config)

    # Ground truth: Book.xlsx (macro report)
    assert r.energy.ep1 == pytest.approx(0.458, abs=0.0005)
    assert r.energy.ep2 == pytest.approx(0.447, abs=0.0005)
    assert r.energy.t1 == pytest.approx(431.8, abs=0.1)
    assert r.energy.t2 == pytest.approx(303.5, abs=0.1)
    assert r.energy.et == pytest.approx(1.677, abs=0.002)
    assert r.energy.et_corr == pytest.approx(1.677, abs=0.002)
    assert r.stability.test_duration_min == pytest.approx(673.7, abs=0.1)