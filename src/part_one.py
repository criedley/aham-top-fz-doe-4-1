"""
part_one.py

Python equivalent of PartOne.bas (EP1Calc, ECalc). Pure functions: no
sheet access and no module-level state.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from models import ChannelConfig, EnergyResult

EP1_WINDOW_ROWS = 180      # EP1Calc looks back 180 scans
EP1_T1_MINUTES = 180.0     # and hard-codes t1 = 180 min (assumes ~1-min scans)


@dataclass
class Ep1Result:
    ep1: float          # kWh
    t1: float           # minutes
    avg_watts: float


def calc_ep1_fixed_window(raw_df: pd.DataFrame, config: ChannelConfig,
                          end_idx: int) -> Ep1Result:
    """
    Equivalent of EP1Calc. end_idx is the 0-based iloc position of the last
    row in the window (VBA's EndRow). Used only on the non-cyclic path.
    """
    start_idx = end_idx - (EP1_WINDOW_ROWS - 1)
    if start_idx < 0 or end_idx >= len(raw_df):
        raise ValueError(
            f"EP1 window [{start_idx}, {end_idx}] does not fit in "
            f"{len(raw_df)} rows of data"
        )

    watt_hours = raw_df[config.watt_hours_col]
    sum_wh = float(watt_hours.iloc[end_idx]) - float(watt_hours.iloc[start_idx])
    avg_watts = float(raw_df[config.watts_col].iloc[start_idx:end_idx + 1].mean())

    return Ep1Result(ep1=sum_wh / 1000, t1=EP1_T1_MINUTES, avg_watts=avg_watts)


def calc_et(ep1: float, t1: float, ep2: float, t2: float, ct: float,
            fz_only: bool, fz_only_kfactor: float) -> EnergyResult:
    """Equivalent of ECalc (AHAM spec 8.7.2.1.2)."""
    if t1 <= 0:
        raise ValueError(f"t1 must be positive, got {t1}")
    if ct <= 0:
        raise ValueError(f"CT must be positive, got {ct}")

    et = (1440 * ep1 / t1) + (ep2 - (ep1 * t2 / t1)) * (12 / ct)
    et_corr = et * fz_only_kfactor    # K is a product-class factor (ETP 8.6 definition of K); 1.0 for refrigerator-freezers
    return EnergyResult(ep1=ep1, ep2=ep2, t1=t1, t2=t2, et=et, et_corr=et_corr)