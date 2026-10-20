"""
report.py

Builds the numbers shown on the macro's "AHAM Report" sheet (TestSummary.bas
AHAMSummary), as plain values. Rendering is a separate step.

Values are rounded where the macro rounds them for display (Energy to 3
decimals, T1 / T2 to 1, averages to 2). Full-precision values stay in
AnalysisResult.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

import pandas as pd

from defrost import Segment, build_t2_table
from models import ChannelConfig, TestConfig

if TYPE_CHECKING:
    from main import AnalysisResult

AMBIENT_LOW_F = 88.95       # AHAMSummary: LSOC / RSOC cycle averages must be inside
AMBIENT_HIGH_F = 91.05      # (88.95, 91.05), exclusive
ICE_MAX_F = 32.0
MAX_TABLE_CYCLES = 100      # AHAMSummary: "Limits number of cycles shown to 100"


@dataclass(frozen=True)
class ReportData:
    test_duration_min: float
    test_duration_hr: float
    energy_kwh_day: float
    ep1_kwh: float
    t1_min: float
    t1_start: pd.Timestamp
    t1_end: pd.Timestamp
    ep2_kwh: float
    t2_min: float
    t2_start: pd.Timestamp
    t2_end: pd.Timestamp
    ct: float
    ff_avg: float | None            # None when freezer-only (macro prints "NA")
    fz_avg: float
    k_factor: float
    energy_corr_kwh_day: float
    lsoc_avg: float
    rsoc_avg: float
    grad_left_lo_mid: float
    grad_left_mid_hi: float
    grad_right_lo_mid: float
    grad_right_mid_hi: float
    ambient_ok: bool
    gradient_ok: bool
    ice_ok: bool | None             # None when no ICE BW channel (macro prints "NA")
    t2_segments: dict[str, Segment]
    t2_table: pd.DataFrame
    stable_cycle: int               # 0-based; highlight this row in the cycle table
    cycle_table: pd.DataFrame       # cycles printed in the macro's main table


def gradient_maxes(cycles: pd.DataFrame, height_ft: float) -> tuple[float, float, float, float]:
    """
    Left Lo-Mid, Left Mid-Hi, Right Lo-Mid, Right Mid-Hi: the maximum over cycles
    1..J-1 (the last cycle is excluded, as in Setup.bas) of
    |round((upper - lower) / height_ft, 3)| using the rounded per-cycle averages.
    """
    body = cycles.iloc[:-1]
    if body.empty:
        return 0.0, 0.0, 0.0, 0.0

    def peak(upper: str, lower: str) -> float:
        return float(((body[upper] - body[lower]) / height_ft).round(3).abs().max())

    return (
        peak("lsoc_avg", "lsoc_lo_avg"),
        peak("lsoc_hi_avg", "lsoc_avg"),
        peak("rsoc_avg", "rsoc_lo_avg"),
        peak("rsoc_hi_avg", "rsoc_avg"),
    )


def ambient_ok(cycles: pd.DataFrame, stable_cycle: int) -> bool:
    """
    AmbientFLG: LSOC and RSOC cycle averages inside (88.95, 91.05), for cycles up to
    and including stable_cycle in the main cycle table.

    Deliberate difference from the macro: AHAMSummary also tests the T2 segment rows
    with `If I <= stabcycle`, comparing the T2 row number with the stable cycle number.
    Those are unrelated counters, so we don't reproduce that.
    """
    window = cycles.iloc[: stable_cycle + 1]
    inside = (window["lsoc_avg"].between(AMBIENT_LOW_F, AMBIENT_HIGH_F, inclusive="neither")
              & window["rsoc_avg"].between(AMBIENT_LOW_F, AMBIENT_HIGH_F, inclusive="neither"))
    return bool(inside.all())


def gradient_ok(cycles: pd.DataFrame, stable_cycle: int, limit: float) -> bool:
    """GradientFLG: left and right gradients within the limit, cycles up to stable_cycle."""
    window = cycles.iloc[: stable_cycle + 1]
    return bool(((window["l_grad"].abs() <= limit) & (window["r_grad"].abs() <= limit)).all())


def ice_ok(cycle_table: pd.DataFrame, config: ChannelConfig) -> bool | None:
    """ICEPass: every printed cycle's ICE BW average is <= 32 F. None if no ICE channel."""
    if config.ice_bw_col is None:
        return None
    return bool((cycle_table["ice_bw"] <= ICE_MAX_F).all())


def pre_defrost_cycles(cycles: pd.DataFrame, defrost_start: pd.Timestamp) -> pd.DataFrame:
    """
    Cycles shown in the macro's main table (its FullCycle): those that start before
    the defrost, capped at 100 and at (number of cycles - 1).
    """
    limit = min(MAX_TABLE_CYCLES, len(cycles) - 1)
    return cycles[cycles["time_on"] < defrost_start].head(limit).reset_index(drop=True)


def _mean_of_window(raw_df: pd.DataFrame, cols: list[str],
                    start: pd.Timestamp, end: pd.Timestamp) -> float:
    """Mean over every cell of the channels for rows with start <= Timestamp < end."""
    window = raw_df[(raw_df["Timestamp"] >= start) & (raw_df["Timestamp"] < end)]
    return float(pd.Series(window[cols].to_numpy().ravel()).mean())


def build_report(analysis: AnalysisResult, test_config: TestConfig) -> ReportData:
    stab = analysis.stability
    if (stab.stable_cycle is None or stab.start_t1_index is None
            or stab.test_duration_min is None):
        raise ValueError("A report needs a stable cyclic analysis")

    raw_df = analysis.import_result.raw_df
    config = analysis.channel_config
    cycles = analysis.cycle_data.cycles
    energy = analysis.energy
    defrost = analysis.defrost

    t1_start = cycles["time_on"].iloc[stab.start_t1_index]
    t1_end = cycles["time_on"].iloc[stab.stable_cycle + 1]

    # FF / FZ Avg: raw-row mean over the T1 window (v1.9f), not the mean of cycle means.
    ff_cols = [config.ff_bw1_col, config.ff_bw2_col, config.ff_bw3_col]
    fz_cols = [c for c in (config.fz_bw1_col, config.fz_bw2_col, config.fz_bw3_col,
                           config.fz_bw4_col, config.fz_bw5_col) if c is not None]
    ff_avg = None if test_config.fz_only else round(_mean_of_window(raw_df, ff_cols, t1_start, t1_end), 2)
    fz_avg = round(_mean_of_window(raw_df, fz_cols, t1_start, t1_end), 2)

    # LSOC / RSOC Avg: from the cycle after the stable one to the end of T2.
    lsoc_avg = round(_mean_of_window(raw_df, [config.lsoc_col], t1_end, defrost.t2_end), 2)
    rsoc_avg = round(_mean_of_window(raw_df, [config.rsoc_col], t1_end, defrost.t2_end), 2)

    height_ft = test_config.height_adjusted_in / 12
    ll, lh, rl, rh = gradient_maxes(cycles, height_ft)

    table = pre_defrost_cycles(cycles, defrost.segments["Defrost"].start)
    t2_table = build_t2_table(raw_df, analysis.import_result.events_df, config, defrost, height_ft)


    return ReportData(
        test_duration_min=round(stab.test_duration_min, 1),
        test_duration_hr=round(stab.test_duration_min / 60, 1),
        energy_kwh_day=round(energy.et, 3),
        ep1_kwh=energy.ep1,
        t1_min=round(energy.t1, 1),
        t1_start=t1_start,
        t1_end=t1_end,
        ep2_kwh=energy.ep2,
        t2_min=round(energy.t2, 1),
        t2_start=defrost.t2_start,
        t2_end=defrost.t2_end,
        ct=test_config.ct,
        ff_avg=ff_avg,
        fz_avg=fz_avg,
        k_factor=test_config.fz_only_kfactor,
        energy_corr_kwh_day=round(energy.et_corr, 3),
        lsoc_avg=lsoc_avg,
        rsoc_avg=rsoc_avg,
        grad_left_lo_mid=ll,
        grad_left_mid_hi=lh,
        grad_right_lo_mid=rl,
        grad_right_mid_hi=rh,
        ambient_ok=ambient_ok(cycles, stab.stable_cycle),
        gradient_ok=gradient_ok(cycles, stab.stable_cycle, test_config.gradient_limit()),
        ice_ok=ice_ok(table, config),
        t2_segments=defrost.segments,
        t2_table=t2_table,
        stable_cycle=stab.stable_cycle,
        cycle_table=table,
    )