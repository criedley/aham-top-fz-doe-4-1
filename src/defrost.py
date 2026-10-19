"""
defrost.py

Python equivalent of DefrostPortion.bas EP2calc (cyclic path only).

Segments are bounded by compressor-off events, so the whole calculation
works on events_df. raw_df is only used for the +/-0.5 degF average checks.
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from models import ChannelConfig

AVG_TOLERANCE = 0.5   # degF, VBA's "StabAvg - 0.5 ... StabAvg + 0.5" checks


@dataclass(frozen=True)
class Segment:
    start: pd.Timestamp
    end: pd.Timestamp
    wh_start: float
    wh_end: float

    @property
    def energy_wh(self) -> float:
        return self.wh_end - self.wh_start


@dataclass(frozen=True)
class DefrostResult:
    t2_start: pd.Timestamp
    t2_end: pd.Timestamp
    ep2_kwh: float
    t2_minutes: float
    segments: dict[str, Segment]


def _off_events(events_df: pd.DataFrame, config: ChannelConfig) -> pd.DataFrame:
    """Rows where Unit On goes 1 -> 0, with their row position in events_df."""
    unit_on = events_df[config.unit_on_col]
    mask = ((unit_on == 0) & (unit_on.shift(1) == 1)).to_numpy()
    offs = events_df.loc[mask, ["Timestamp", config.watt_hours_col]]
    offs = offs.rename(columns={config.watt_hours_col: "wh"}).reset_index(drop=True)
    offs["event_pos"] = np.flatnonzero(mask)
    return offs


def _first_defrost_pos(events_df: pd.DataFrame, config: ChannelConfig) -> int:
    hits = np.flatnonzero((events_df[config.defrost_on_col] == 1).to_numpy())
    if len(hits) == 0:
        raise ValueError("No freezer defrost in data set. Test failed.")
    return int(hits[0])


def _boundary_index(offs: pd.DataFrame, defrost_pos: int) -> int:
    """
    Index (into offs) of the off event that starts the defrost segment:
    the last off event at or before the first Defrost = 1 event ROW.
    Compared by row position, not timestamp, because the off / defrost-on /
    on+defrost events share one timestamp.
    """
    before = np.flatnonzero((offs["event_pos"] <= defrost_pos).to_numpy())
    if len(before) == 0:
        raise ValueError("No compressor-off event before the first defrost")
    return int(before[-1])


def _channel_cols(config: ChannelConfig) -> tuple[list[str], list[str]]:
    ff = [config.ff_bw1_col, config.ff_bw2_col, config.ff_bw3_col]
    fz = [c for c in (config.fz_bw1_col, config.fz_bw2_col, config.fz_bw3_col,
                      config.fz_bw4_col, config.fz_bw5_col) if c is not None]
    return ff, fz


def _mean_temp(raw_df: pd.DataFrame, seg: Segment, cols: list[str]) -> float:
    """Mean over every cell in the window, like Excel's AVERAGE on a 2-D range."""
    window = raw_df[(raw_df["Timestamp"] >= seg.start) & (raw_df["Timestamp"] < seg.end)]
    return float(pd.Series(window[cols].to_numpy().ravel()).mean())


def _passes(raw_df: pd.DataFrame, seg: Segment, config: ChannelConfig,
            ff_avg: float, fz_avg: float) -> bool:
    ff_cols, fz_cols = _channel_cols(config)
    ff = _mean_temp(raw_df, seg, ff_cols)
    fz = _mean_temp(raw_df, seg, fz_cols)
    return abs(fz - fz_avg) <= AVG_TOLERANCE and abs(ff - ff_avg) <= AVG_TOLERANCE


def calc_ep2(raw_df: pd.DataFrame, events_df: pd.DataFrame, config: ChannelConfig,
             ff_avg: float, fz_avg: float) -> DefrostResult:
    """
    Cyclic-path EP2 / T2. ff_avg and fz_avg are the T1-window averages from
    StabilityResult.

    Layout (off-event indices into `offs`):
        a      cycle A  : offs[a]   -> offs[a+1]     (retried backwards)
        ...    precool  : offs[a+1] -> offs[d]
        d      defrost  : offs[d]   -> offs[d+1]
        d+1    recovery : offs[d+1] -> offs[d+2]
        ...    post     : offs[d+2] -> offs[b]
        b      cycle B  : offs[b]   -> offs[b+1]     (retried forwards)
    T2 runs from offs[a+1] to offs[b].
    """
    offs = _off_events(events_df, config)
    d = _boundary_index(offs, _first_defrost_pos(events_df, config))
    if d < 2:
        raise ValueError("Not enough cycles before defrost for a precool and an A cycle")

    def seg(i: int, j: int | None = None) -> Segment:
        j = i + 1 if j is None else j
        return Segment(start=offs["Timestamp"].iloc[i], end=offs["Timestamp"].iloc[j],
                       wh_start=float(offs["wh"].iloc[i]), wh_end=float(offs["wh"].iloc[j]))

    # Cycle A: step back one cycle at a time until its averages are within tolerance.
    a = d - 2
    while a >= 0 and not _passes(raw_df, seg(a), config, ff_avg, fz_avg):
        a -= 1
    if a < 0:
        raise ValueError("Part 2 Cycle A (before defrost) never stabilizes. Test failed.")

    # Cycle B: first cycle after post-defrost cycle 1 whose averages are within tolerance.
    b = d + 3
    while True:
        if b + 1 >= len(offs):
            raise ValueError("Part 2 Cycle B (after defrost) never stabilizes. Test failed.")
        if _passes(raw_df, seg(b), config, ff_avg, fz_avg):
            break
        b += 1

    start = a + 1
    t2 = seg(start, b)
    return DefrostResult(
        t2_start=t2.start,
        t2_end=t2.end,
        ep2_kwh=t2.energy_wh / 1000,
        t2_minutes=(t2.end - t2.start).total_seconds() / 60,
        segments={
            "A": seg(a),
            "Precool": seg(start, d),
            "Defrost": seg(d),
            "Recovery": seg(d + 1),
            "Post": seg(d + 2, b),
            "B": seg(b),
        },
    )