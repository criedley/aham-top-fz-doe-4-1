# tests/test_checks_real_file.py
from pathlib import Path
import pytest

from io_import import load_datacart_csv
from setup import build_channel_config
from cycles import detect_cycles
from checks import check_stability
from models import TestConfig

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw.csv"

def test_stability_matches_original_macro_on_7272_file(alias_map, channel_map):
    # alias_map and channel_map: the dicts from scratch_run.py, moved into conftest.py
    # CABINET_HEIGHT in the file header is 68.5, +10" t/c offset
    test_config = TestConfig(height_in=68.5, height_adjusted_in=78.5, ct=96.0,
                             fz_only=False, fz_only_kfactor=1.0, adjusted_start=False)
    imp = load_datacart_csv(FIXTURE, alias_map)
    config = build_channel_config(imp, channel_map)
    cycle_data = detect_cycles(imp, config, test_config)

    result = check_stability(imp.raw_df, cycle_data, config, test_config)

    assert result.is_stable
    assert result.stable_cycle == 6          # macro: stabcycle = 7 (1-indexed)
    assert result.start_t1_index == 2        # macro: StartT1 = 3
    assert result.ep1 == pytest.approx(0.458)
    assert result.t1_minutes == pytest.approx(431.78, abs=0.01)
    assert result.test_duration_min == pytest.approx(673.7, abs=0.05)