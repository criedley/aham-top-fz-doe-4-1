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
from cycles import window_stats

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
    pre_cycles: tuple[Segment, ...] = ()        # failed Cycle A candidates, oldest first ("Pre 1", ...)
    post_cycles: tuple[Segment, ...] = ()       # failed Cycle B candidates ("Post 2", ...)
    defrost_start: pd.Timestamp | None = None   # first event row with Defrost = 1


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
    offs = _off_events(events_df, config)
    defrost_pos = _first_defrost_pos(events_df, config)
    d = _boundary_index(offs, defrost_pos)
    if d < 2:
        raise ValueError("Not enough cycles before defrost for a precool and an A cycle")

    def seg(i: int, j: int | None = None) -> Segment:
        j = i + 1 if j is None else j
        return Segment(start=offs["Timestamp"].iloc[i], end=offs["Timestamp"].iloc[j],
                       wh_start=float(offs["wh"].iloc[i]), wh_end=float(offs["wh"].iloc[j]))

    a = d - 2
    while a >= 0 and not _passes(raw_df, seg(a), config, ff_avg, fz_avg):
        a -= 1
    if a < 0:
        raise ValueError("Part 2 Cycle A (before defrost) never stabilizes. Test failed.")

    b = d + 3
    while True:
        if b + 1 >= len(offs):
            raise ValueError("Part 2 Cycle B (after defrost) never stabilizes. Test failed.")
        if _passes(raw_df, seg(b), config, ff_avg, fz_avg):
            break
        b += 1

    t2 = seg(a + 1, b)
    return DefrostResult(
        t2_start=t2.start,
        t2_end=t2.end,
        ep2_kwh=t2.energy_wh / 1000,
        t2_minutes=(t2.end - t2.start).total_seconds() / 60,
        segments={
            "A": seg(a),
            "Precool": seg(d - 1),
            "Defrost": seg(d),
            "Recovery": seg(d + 1),
            "Post": seg(d + 2),
            "B": seg(b),
        },
        pre_cycles=tuple(seg(i) for i in range(a + 1, d - 1)),
        post_cycles=tuple(seg(i) for i in range(d + 3, b)),
        defrost_start=events_df["Timestamp"].iloc[defrost_pos],
    )


def _on_event_ts(events_df: pd.DataFrame, config: ChannelConfig, seg: Segment) -> pd.Timestamp:
    """First Unit On = 1 event inside the segment (the compressor start after its off event)."""
    hits = events_df.loc[
        (events_df["Timestamp"] >= seg.start) & (events_df["Timestamp"] <= seg.end)
        & (events_df[config.unit_on_col] == 1), "Timestamp"]
    if hits.empty:
        raise ValueError(f"No compressor-on event between {seg.start} and {seg.end}")
    return hits.iloc[0]


def build_t2_table(raw_df: pd.DataFrame, events_df: pd.DataFrame, config: ChannelConfig,
                   result: DefrostResult, height_ft: float) -> pd.DataFrame:
    """
    The macro's "T2 Data" table (DefrostPortion.bas CycleData). One row per segment, in
    time order: A Cycle, Pre n, Precool, Defrost, Recovery, Post n, B Cycle.

    Timing columns come from the segment's events. Temperatures are averaged over raw scan
    rows; the row window differs by segment, copied from the macro's row walks:
      A / Pre / Post / B : whole cycle (first scan after the off event .. last scan before the next)
      Precool            : the on-run plus the scan before it
      Defrost            : first Defrost = 1 scan .. first scan after the defrost ends (inclusive)
      Recovery           : the on-run .. first scan after it turns off (inclusive)
    """
    if result.defrost_start is None:
        raise ValueError("DefrostResult has no defrost_start")

    # The macro's Raw Data sheet has event-only rows removed; keep real scans only.
    scan = raw_df[raw_df[config.watts_col].notna()].reset_index(drop=True)

    def lb(t: pd.Timestamp) -> int:
        """Index of the first scan at or after t."""
        return int(scan["Timestamp"].searchsorted(t))

    segs = result.segments
    precool, defrost, recovery = segs["Precool"], segs["Defrost"], segs["Recovery"]

    def whole(s: Segment) -> tuple[int, int]:
        return lb(s.start), lb(s.end)

    specs: list[tuple[str, Segment, int, int]] = [("A Cycle", segs["A"], *whole(segs["A"]))]
    specs += [(f"Pre {n}", s, *whole(s)) for n, s in enumerate(result.pre_cycles, 1)]
    specs += [
        ("Precool", precool, lb(_on_event_ts(events_df, config, precool)) - 1, lb(precool.end)),
        ("Defrost", defrost, lb(result.defrost_start), lb(defrost.end) + 1),
        ("Recovery", recovery, lb(_on_event_ts(events_df, config, recovery)), lb(recovery.end) + 1),
        ("Post 1", segs["Post"], *whole(segs["Post"])),
    ]
    specs += [(f"Post {n}", s, *whole(s)) for n, s in enumerate(result.post_cycles, 2)]
    specs.append(("B Cycle", segs["B"], *whole(segs["B"])))

    records = []
    for label, s, lo, hi in specs:
        on_ts = _on_event_ts(events_df, config, s)
        off_min = round((on_ts - s.start).total_seconds() / 60, 2)
        on_min = round((s.end - on_ts).total_seconds() / 60, 2)
        records.append({
            "label": label,
            **window_stats(scan.iloc[lo:hi], config, height_ft),
            "duration_min": round((s.end - s.start).total_seconds() / 60, 2),
            "on_min": on_min,
            "off_min": off_min,
            "runtime": on_min / (on_min + off_min) if on_min + off_min > 0 else float("nan"),
            "wh": s.energy_wh,
            "start": s.start,
            "end": s.end,
        })
    return pd.DataFrame.from_records(records)