# cycles.py
import numpy as np
import pandas as pd
from dataclasses import dataclass
from .models import CycleData, ChannelConfig, TestConfig

@dataclass
class CycleBoundaries:
    time_on: list[pd.Timestamp]      # TimeCycleOn()
    time_off: list[pd.Timestamp]     # TimeCycleOff()
    index_on: list[int]              # IndexCycleOn()
    index_off: list[int]             # IndexCycleOff()
    watt_hours_at_on: list[float]    # WattHoursCount()
    cyclic: bool                     # cyclic flag
    last_scan_index: int             # LastScanIndex
    last_scan_time: pd.Timestamp     # LastScan

# def compute_cycle_stats(raw_df, boundaries, config) -> CycleData: return

# def detect_cycles(raw_df, events_df, config) -> CycleData: return
    

def _find_boundaries_from_events(events_df: pd.DataFrame, raw_df: pd.DataFrame,
                                  config: "ChannelConfig",
                                  test_config: "TestConfig") -> CycleBoundaries:
    """
    Event-based cycle detection. Equivalent to the main loop in Setup.bas's
    GetCycles: pairs each Unit-On "1" event with the next "0" event to
    define one compressor cycle, replacing the VBA .Find() row-by-row scan.
    """
    unit_on = events_df[config.unit_on_col]
    timestamps = events_df["Timestamp"]

    # Mirrors AdjustedStart: optionally skip to a technician-specified start time
    if test_config.adjusted_start:
        start_mask = timestamps >= test_config.time_to_start
        if not start_mask.any():
            raise ValueError("No events found after requested start time")
        events_df = events_df.loc[start_mask.idxmax():]
        unit_on = events_df[config.unit_on_col]
        timestamps = events_df["Timestamp"]

    on_rows = events_df[unit_on == 1]
    off_rows = events_df[unit_on == 0]

    time_on, time_off, wh_at_on = [], [], []
    off_positions = off_rows.index.to_numpy()
    cyclic = True

    for on_idx, on_row in on_rows.iterrows():
        # find the first "off" event after this "on" event
        later_offs = off_positions[off_positions > on_idx]
        if len(later_offs) == 0:
            # no matching off -> incomplete trailing cycle, matches
            # GetCycles' "Start > x.Row" / cyclic = False exit condition
            cyclic = len(time_on) > 0
            break
        off_idx = later_offs[0]
        time_on.append(on_row["Timestamp"])
        time_off.append(off_rows.loc[off_idx, "Timestamp"])
        wh_at_on.append(on_row[config.watt_hours_col])

    if not time_on:
        raise ValueError("No Cycle On/Off events detected, test failed")

    # Resolve event timestamps to row positions in raw_df (vectorized,
    # replaces GetCycles' per-cycle "Do While DateDiff(...) <= 0" scan)
    raw_times = raw_df["Timestamp"].to_numpy()
    index_on = np.searchsorted(raw_times, pd.to_datetime(time_on)).tolist()
    index_off = np.searchsorted(raw_times, pd.to_datetime(time_off)).tolist()

    return CycleBoundaries(
        time_on=time_on,
        time_off=time_off,
        index_on=index_on,
        index_off=index_off,
        watt_hours_at_on=wh_at_on,
        cyclic=cyclic,
        last_scan_index=len(raw_df) - 1,
        last_scan_time=raw_df["Timestamp"].iloc[-1],
    )

def _find_stability_window(raw_df: pd.DataFrame, timestamps: pd.Series,
                            bw_cols: list[str]) -> int:
    """
    Equivalent of Run100's stability search: slides a 2hr/3hr/2hr
    (A / gap / B) window across the data until every BW channel's average
    in window A matches window B within the AHAM threshold
    (0.042 degF per hour of total elapsed span).
    """
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