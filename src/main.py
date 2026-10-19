"""
src/main.py

Orchestrates the cyclic-path pipeline: import -> channel config -> cycles ->
stability (EP1 / T1) -> defrost (EP2 / T2) -> ET.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from channel_defaults import ALIAS_MAP
from checks import check_stability
from cycles import detect_cycles
from defrost import DefrostResult, calc_ep2
from io_import import load_datacart_csv
from models import (ChannelConfig, CycleData, EnergyResult, ImportResult,
                    StabilityResult, TestConfig)
from part_one import calc_et
from setup import build_channel_config


@dataclass
class AnalysisResult:
    energy: EnergyResult
    stability: StabilityResult
    defrost: DefrostResult
    cycle_data: CycleData
    import_result: ImportResult
    channel_config: ChannelConfig


def run_analysis(csv_path: str | Path, channel_map: dict[str, str],
                 test_config: TestConfig,
                 alias_map: dict[str, list[str]] = ALIAS_MAP) -> AnalysisResult:
    imp = load_datacart_csv(csv_path, alias_map)
    config = build_channel_config(imp, channel_map)
    cycle_data = detect_cycles(imp, config, test_config)

    if not cycle_data.cyclic:
        # Needs the watt-band EP2 path and calc_ep1_fixed_window, both unvalidated.
        raise NotImplementedError("Non-cyclic (fixed-window) tests are not supported yet")

    stab = check_stability(imp.raw_df, cycle_data, config, test_config)
    if not stab.is_stable:
        raise ValueError(f"Test is not stable: {stab.reason}")
    assert stab.ep1 is not None and stab.t1_minutes is not None
    assert stab.ff_avg is not None and stab.fz_avg is not None

    defrost = calc_ep2(imp.raw_df, imp.events_df, config, stab.ff_avg, stab.fz_avg)

    energy = calc_et(
        ep1=stab.ep1, t1=stab.t1_minutes,
        ep2=defrost.ep2_kwh, t2=defrost.t2_minutes,
        ct=test_config.ct,
        fz_only=test_config.fz_only, fz_only_kfactor=test_config.fz_only_kfactor,
    )
    return AnalysisResult(energy, stab, defrost, cycle_data, imp, config)