"""
cycles.py

Python equivalent of Setup.bas's GetCycles + Module1.bas's Run100.
Detects compressor on/off cycle boundaries, dispatching between
event-based detection (normal case) and fixed-window stability-based
detection (100%-duty-cycle / no-events fallback), then computes
per-cycle averages and gradients.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from models import ImportResult, ChannelConfig, TestConfig, CycleData


@dataclass
class CycleBoundaries:
    """
    Private intermediate result shared by both detection strategies.
    Never leaves this module -- detect_cycles consumes it and returns
    the public CycleData instead.
    """
    time_on: list[pd.Timestamp]
    time_off: list[pd.Timestamp]
    index_on: list[int]
    index_off: list[int]
    watt_hours_at_on: list[float]
    cyclic: bool
    last_scan_index: int
    last_scan_time: pd.Timestamp


def detect_cycles(import_result: ImportResult, config: ChannelConfig,
                   test_config: TestConfig) -> CycleData:
    """Main entry point. Dispatches to event-based or fixed-window detection,
    then computes per-cycle stats, matching GetCycles' two-path structure
    with Run100 as the 'eventcount < 1' fallback."""
    if len(import_result.events_df) < 1:
        boundaries = _find_boundaries_fixed_window(import_result.raw_df, config)
        method = "fixed_window"
    else:
        boundaries = _find_boundaries_from_events(
            import_result.events_df, import_result.raw_df, config, test_config
        )
        method = "event_based"

    cycles_df = _compute_cycle_stats(import_result.raw_df, boundaries, config, test_config)

    return CycleData(
        cycles=cycles_df,
        detection_method=method,
        cyclic=boundaries.cyclic,
        last_scan_index=boundaries.last_scan_index,
        last_scan_time=boundaries.last_scan_time,
    )


# ---------------------------------------------------------------------------
# Event-based detection (GetCycles' main loop)
# ---------------------------------------------------------------------------

def _find_boundaries_from_events(events_df: pd.DataFrame, raw_df: pd.DataFrame,
                                  config: ChannelConfig,
                                  test_config: TestConfig) -> CycleBoundaries:
    unit_on = events_df[config.unit_on_col]
    timestamps = events_df["Timestamp"]

    if test_config.adjusted_start:
        start_mask = timestamps >= test_config.time_to_start
        if not start_mask.any():
            raise ValueError("No events found after requested start time")
        events_df = events_df.loc[start_mask.idxmax():]
        unit_on = events_df[config.unit_on_col]
        timestamps = events_df["Timestamp"]

    # Only treat a row as an on/off boundary if Unit On ITSELF changed on
    # that row -- not merely if its current value happens to be 1/0.
    # events_df can contain rows triggered by Defrost/Comp changes where
    # Unit On's value is just carried over from the previous state.
    unit_on_changed = unit_on.diff().fillna(unit_on.iloc[0] if len(unit_on) else 0) != 0
    on_rows = events_df[unit_on_changed & (unit_on == 1)]
    off_rows = events_df[unit_on_changed & (unit_on == 0)]

    time_on, time_off, wh_at_on = [], [], []
    off_positions = off_rows.index.to_numpy()
    cyclic = True

    for on_idx, on_row in on_rows.iterrows():
        later_offs = off_positions[off_positions > on_idx]
        if len(later_offs) == 0:
            cyclic = len(time_on) > 0
            break
        off_idx = later_offs[0]
        time_on.append(on_row["Timestamp"])
        time_off.append(off_rows.loc[off_idx, "Timestamp"])
        wh_at_on.append(on_row[config.watt_hours_col])

    if not time_on:
        raise ValueError("No Cycle On/Off events detected, test failed")

    raw_times = raw_df["Timestamp"].to_numpy()
    index_on = np.searchsorted(raw_times, pd.to_datetime(time_on)).tolist()
    index_off = np.searchsorted(raw_times, pd.to_datetime(time_off)).tolist()

    return CycleBoundaries(
        time_on=time_on, time_off=time_off,
        index_on=index_on, index_off=index_off,
        watt_hours_at_on=wh_at_on, cyclic=cyclic,
        last_scan_index=len(raw_df) - 1,
        last_scan_time=raw_df["Timestamp"].iloc[-1],
    )


# ---------------------------------------------------------------------------
# Fixed-window detection (Run100 fallback, for 100%-duty-cycle tests)
# ---------------------------------------------------------------------------

def _find_boundaries_fixed_window(raw_df: pd.DataFrame,
                                   config: ChannelConfig) -> CycleBoundaries:
    # always builds bw_cols from both FF and FZ channels, with no fz_only branching at all.
    timestamps = raw_df["Timestamp"]
    bw_cols = [c for c in (
        config.ff_bw1_col, config.ff_bw2_col, config.ff_bw3_col,
        config.fz_bw1_col, config.fz_bw2_col, config.fz_bw3_col,
        config.fz_bw4_col, config.fz_bw5_col,
    ) if c is not None]

    start_a = _find_stability_window(raw_df, timestamps, bw_cols)

    window_s = 1799
    index_on, time_on = [], []
    index_off, time_off, watt_hours = [], [], []

    i = start_a
    while i < len(raw_df):
        index_on.append(i)
        time_on.append(timestamps.iloc[i])

        off_target = timestamps.iloc[i] + pd.Timedelta(seconds=window_s)
        off_mask = timestamps.iloc[i:] >= off_target
        if not off_mask.any():
            break
        off_idx = off_mask.idxmax()
        index_off.append(off_idx)
        time_off.append(timestamps.iloc[off_idx])
        watt_hours.append(raw_df[config.watt_hours_col].iloc[off_idx - 1])

        next_on_target = timestamps.iloc[off_idx] + pd.Timedelta(seconds=window_s)
        next_mask = timestamps.iloc[off_idx:] >= next_on_target
        if not next_mask.any():
            break
        i = next_mask.idxmax()

    return CycleBoundaries(
        time_on=time_on[: len(index_off)],      # <- now truncated to match, was unguarded before
        time_off=time_off,
        index_on=index_on[: len(index_off)],
        index_off=index_off,
        watt_hours_at_on=watt_hours, cyclic=False,
        last_scan_index=len(raw_df) - 1,
        last_scan_time=timestamps.iloc[-1],
    )


def _find_stability_window(raw_df: pd.DataFrame, timestamps: pd.Series,
                            bw_cols: list[str]) -> int:
    start_a = 1
    n = len(raw_df)

    while start_a < n:
        end_a_mask = timestamps.iloc[start_a:] >= timestamps.iloc[start_a] + pd.Timedelta(seconds=7200)
        if not end_a_mask.any():
            raise ValueError("Stability not found")
        end_a = end_a_mask.idxmax()

        end_gap_mask = timestamps.iloc[end_a:] >= timestamps.iloc[end_a] + pd.Timedelta(seconds=10800)
        if not end_gap_mask.any():
            raise ValueError("Stability not found")
        end_gap = end_gap_mask.idxmax()

        start_b = end_gap
        end_b_mask = timestamps.iloc[start_b:] >= timestamps.iloc[start_b] + pd.Timedelta(seconds=7200)
        if not end_b_mask.any():
            raise ValueError("Stability not found")
        end_b = end_b_mask.idxmax()

        stab_time_s = (
            (timestamps.iloc[end_b] - timestamps.iloc[start_b]).total_seconds() / 2
            + (timestamps.iloc[end_gap] - timestamps.iloc[end_a]).total_seconds()
            + (timestamps.iloc[end_a] - timestamps.iloc[start_a]).total_seconds() / 2
        )
        threshold = 0.042 * (stab_time_s / 3600)

        stable = all(
            abs(raw_df[col].iloc[start_a:end_a].mean() - raw_df[col].iloc[start_b:end_b].mean()) <= threshold
            for col in bw_cols
        )
        if stable:
            return start_a

        start_a += 1

    raise ValueError("Stability not found")


# ---------------------------------------------------------------------------
# Shared per-cycle averaging (the "Run100NextLine" convergence point)
# ---------------------------------------------------------------------------

def window_stats(window: pd.DataFrame, config: ChannelConfig, height_ft: float) -> dict[str, float]:
    """Channel averages (2 dp), volts min/max and gradients (3 dp) for one window of raw rows."""
    def avg(col: str | None) -> float:
        """Matches the VBA's 'If ch <> NA Then average Else 0' pattern."""
        if col is None:
            return 0.0
        return round(float(window[col].mean()), 2)

    if config.volts_col is not None:
        volts = window[config.volts_col]
        volts_avg = round(float(volts.mean()), 2)
        volts_min = round(float(volts.min()), 2)
        volts_max = round(float(volts.max()), 2)
    else:
        volts_avg = volts_min = volts_max = 0.0

    rsoc_hi_avg, rsoc_lo_avg = avg(config.rsoc_hi_col), avg(config.rsoc_lo_col)
    lsoc_hi_avg, lsoc_lo_avg = avg(config.lsoc_hi_col), avg(config.lsoc_lo_col)

    return {
        "ff_bw1": avg(config.ff_bw1_col), "ff_bw2": avg(config.ff_bw2_col), "ff_bw3": avg(config.ff_bw3_col),
        "fz_bw1": avg(config.fz_bw1_col), "fz_bw2": avg(config.fz_bw2_col), "fz_bw3": avg(config.fz_bw3_col),
        "fz_bw4": avg(config.fz_bw4_col), "fz_bw5": avg(config.fz_bw5_col),
        "ice_bw": avg(config.ice_bw_col),
        "volts_avg": volts_avg, "volts_min": volts_min, "volts_max": volts_max,
        "rsoc_avg": avg(config.rsoc_col), "lsoc_avg": avg(config.lsoc_col),
        "rsoc_hi_avg": rsoc_hi_avg, "rsoc_lo_avg": rsoc_lo_avg,
        "lsoc_hi_avg": lsoc_hi_avg, "lsoc_lo_avg": lsoc_lo_avg,
        "r_grad": abs(round((rsoc_hi_avg - rsoc_lo_avg) / height_ft, 3)),
        "l_grad": abs(round((lsoc_hi_avg - lsoc_lo_avg) / height_ft, 3)),
    }


def _compute_cycle_stats(raw_df: pd.DataFrame, boundaries: CycleBoundaries,
                          config: ChannelConfig, test_config: TestConfig) -> pd.DataFrame:
    """
    Computes per-cycle averages and gradients, matching GetCycles' averaging
    loop (FFBW1-3, FZBW1-5, ICEBW, Volts, RSOC/LSOC avg+hi+lo, gradients),
    using the already-resolved cycle boundaries from either detection
    strategy. This is the "Run100NextLine" convergence point both
    detection paths share in the original VBA.
    """
    n = len(boundaries.index_on)
    if n == 0:
        raise ValueError("No cycle boundaries available to compute statistics from")

    height_ft = test_config.height_adjusted_in / 12
    records = []

    for i in range(n):
        start = boundaries.index_on[i]
        end = (boundaries.index_on[i + 1] - 1) if i + 1 < n else boundaries.last_scan_index
        stats = window_stats(raw_df.iloc[start:end + 1], config, height_ft)

        # Duty cycle: unchanged from your current code
        time_on_i = boundaries.time_on[i]
        time_off_i = boundaries.time_off[i]
        next_time_on = boundaries.time_on[i + 1] if i + 1 < n else boundaries.last_scan_time
        on_duration_min = (time_off_i - time_on_i).total_seconds() / 60
        total_min = (next_time_on - time_on_i).total_seconds() / 60
        duty_cycle = on_duration_min / total_min if total_min > 0 else float("nan")

        records.append({
            "cycle_num": i + 1,
            "time_on": time_on_i,
            "time_off": time_off_i,
            "index_on": start,
            "index_off": boundaries.index_off[i],
            "watt_hours_at_on": boundaries.watt_hours_at_on[i],
            "duty_cycle": duty_cycle,
            **stats,
        })
    return pd.DataFrame.from_records(records)