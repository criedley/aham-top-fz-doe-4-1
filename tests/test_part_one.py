# tests/test_part_one.py
import pandas as pd
import pytest

from part_one import calc_et, calc_ep1_fixed_window


def test_calc_et_hand_computed():
    # 1440*1/720 = 2.0 ; (0.5 - 1*60/720) = 0.416667 ; 12/24 = 0.5
    r = calc_et(ep1=1.0, t1=720, ep2=0.5, t2=60, ct=24,
                fz_only=False, fz_only_kfactor=1.0)
    assert r.et == pytest.approx(2.0 + 0.416667 * 0.5, rel=1e-5)   # 2.208333
    assert r.et_corr == pytest.approx(r.et)


def test_ct_of_12_makes_defrost_term_unscaled():
    r = calc_et(1.0, 720, 0.5, 60, ct=12, fz_only=False, fz_only_kfactor=1.0)
    assert r.et == pytest.approx(2.0 + 0.416667, rel=1e-5)           # 2.416667


def test_k_factor_applied_only_when_fz_only():
    on = calc_et(1.0, 720, 0.5, 60, 96, fz_only=True, fz_only_kfactor=0.85)
    off = calc_et(1.0, 720, 0.5, 60, 96, fz_only=False, fz_only_kfactor=0.85)
    assert on.et_corr == pytest.approx(on.et * 0.85)
    assert off.et_corr == pytest.approx(off.et)       # ECalc ignores K when not FZ-only


@pytest.mark.parametrize("t1, ct", [(0, 96), (-5, 96), (720, 0), (720, -1)])
def test_calc_et_rejects_nonpositive_t1_or_ct(t1, ct):
    with pytest.raises(ValueError):
        calc_et(1.0, t1, 0.5, 60, ct, False, 1.0)


def _watts_df(n=500):
    return pd.DataFrame({"Watts": [100.0] * n, "Watt-Hours": list(range(n))})


def test_ep1_window_matches_vba_179_interval_quirk(channel_config):
    r = calc_ep1_fixed_window(_watts_df(), channel_config, end_idx=300)
    assert r.ep1 == pytest.approx(0.179)      # 300 - 121 = 179 Wh, not 180
    assert r.t1 == 180.0
    assert r.avg_watts == pytest.approx(100.0)


def test_ep1_raises_when_window_does_not_fit(channel_config):
    with pytest.raises(ValueError):
        calc_ep1_fixed_window(_watts_df(), channel_config, end_idx=100)    # start < 0
    with pytest.raises(ValueError):
        calc_ep1_fixed_window(_watts_df(500), channel_config, end_idx=500)  # past the end

def test_calc_et_reproduces_macro_on_7272_file():
    # Ground truth from the original macro on the 7272 sample file.
    # EP2 is only known to 3 decimals, so allow a small tolerance.
    t1 = 25907 / 60                      # 431.7833 min, matches the checks.py result
    r = calc_et(ep1=0.458, t1=t1, ep2=0.447, t2=303.5, ct=10,
                fz_only=False, fz_only_kfactor=1.0)
    assert r.et == pytest.approx(1.677, abs=0.002)
    assert r.et_corr == pytest.approx(1.677, abs=0.002)