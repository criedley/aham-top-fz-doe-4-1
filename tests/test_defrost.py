# tests/test_defrost.py
import pandas as pd
import pytest

from checks import check_stability
from defrost import calc_ep2

BASE = pd.Timestamp("2026-01-27 00:00:00")
OFFS = [0, 60, 120, 180, 240, 256, 316, 376, 436, 496, 556]
DEFROST_IDX = 4          # defrost starts at the off event at minute 240


def _ts(m: int) -> pd.Timestamp:
    return BASE + pd.Timedelta(minutes=m)


def make_events(offs=OFFS, d=DEFROST_IDX, with_defrost=True) -> pd.DataFrame:
    rows = [(offs[0] - 20, 1, 0)]
    for i, t in enumerate(offs):
        nxt = offs[i + 1] if i + 1 < len(offs) else t + 60
        if with_defrost and i == d:
            rows += [(t, 0, 0), (t, 0, 1), (t + 1, 1, 1)]   # off, defrost, on+defrost
        else:
            rows += [(t, 0, 0), (t + (nxt - t) * 2 // 3, 1, 0)]
    return pd.DataFrame({
        "Timestamp": [_ts(m) for m, _, _ in rows],
        "Watt-Hours": [10.0 * m for m, _, _ in rows],
        "Event 1 Unit On": [u for _, u, _ in rows],
        "Event 2 Defrost": [f for _, _, f in rows],
    })


def make_raw(bad_ranges=()) -> pd.DataFrame:
    minutes = range(0, 620)
    fz = [2.0 if any(lo <= m < hi for lo, hi in bad_ranges) else -0.6 for m in minutes]
    return pd.DataFrame({
        "Timestamp": [_ts(m) for m in minutes],
        "FF BW 1": 37.6, "FF BW 2": 37.6, "FF BW 3": 37.6,
        "FZ BW 1": fz, "FZ BW 2": fz, "FZ BW 3": fz,
    })


def run(channel_config, bad_ranges=(), **kw):
    return calc_ep2(make_raw(bad_ranges), make_events(**kw), channel_config,
                    ff_avg=37.6, fz_avg=-0.6)


def test_base_case_t2_spans_precool_to_cycle_b(channel_config):
    r = run(channel_config)
    assert r.t2_start == _ts(180)               # off[3]
    assert r.t2_end == _ts(376)                 # off[7]
    assert r.t2_minutes == pytest.approx(196)
    assert r.ep2_kwh == pytest.approx((3760 - 1800) / 1000)
    assert r.segments["Defrost"].start == _ts(240)
    assert r.segments["Defrost"].end == _ts(256)


def test_cycle_a_retry_moves_t2_start_back(channel_config):
    r = run(channel_config, bad_ranges=[(120, 180)])   # A = 120..180 fails
    assert r.segments["A"].start == _ts(60)
    assert r.t2_start == _ts(120)                       # end of the accepted A


def test_cycle_b_retry_moves_t2_end_forward(channel_config):
    r = run(channel_config, bad_ranges=[(376, 436)])   # first B fails
    assert r.t2_end == _ts(436)
    assert r.segments["B"].start == _ts(436)


def test_raises_when_cycle_a_never_stabilizes(channel_config):
    with pytest.raises(ValueError, match="Cycle A"):
        run(channel_config, bad_ranges=[(0, 180)])


def test_raises_when_cycle_b_never_stabilizes(channel_config):
    with pytest.raises(ValueError, match="Cycle B"):
        run(channel_config, bad_ranges=[(376, 620)])


def test_raises_without_defrost(channel_config):
    with pytest.raises(ValueError, match="defrost"):
        run(channel_config, with_defrost=False)


def test_matches_macro_on_7272_file(real_pipeline, real_test_config):
    imp, config, cycle_data = real_pipeline
    stab = check_stability(imp.raw_df, cycle_data, config, real_test_config)
    assert stab.ff_avg is not None and stab.fz_avg is not None

    r = calc_ep2(imp.raw_df, imp.events_df, config, stab.ff_avg, stab.fz_avg)

    # Ground truth: Book.xlsx (macro report)
    assert r.t2_start == pd.Timestamp("2026-01-28 16:22:22")
    assert r.t2_end == pd.Timestamp("2026-01-28 21:25:53")
    assert r.ep2_kwh == pytest.approx(0.447, abs=0.0005)
    assert r.t2_minutes == pytest.approx(303.5, abs=0.1)
    for name, wh in {"A": 92, "Precool": 78, "Defrost": 82, "Recovery": 194, "Post": 93}.items():
        assert r.segments[name].energy_wh == pytest.approx(wh, abs=1), name