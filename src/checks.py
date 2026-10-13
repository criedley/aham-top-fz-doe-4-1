"""
checks.py

Python equivalent of Checks.bas: GradientCheck, AmbientCheck,
StabilityCheck, StabilityCheckFZOnly.

Scope: determines whether a test has reached AHAM-required thermal
stability and identifies the T1 energy-measurement window. Does NOT
perform EP2/ET calculation or report generation -- those remain the
responsibility of defrost.py / part_one.py / report.py, orchestrated
by main.py.
"""

from __future__ import annotations

import logging
import pandas as pd

from models import ChannelConfig, TestConfig, CycleData, GradientResult, AmbientResult, StabilityResult

logger = logging.getLogger(__name__)

STABILITY_RATE_PER_HOUR = 0.042  # AHAM standard drift threshold, degF/hr
WINDOW_ROWS = 120                 # ~2 hours at 1-min scans, matches VBA's 119-row ranges
MIN_K_GAP_ROWS = 120
MIN_L_GAP_ROWS = 180
MIN_M_GAP_ROWS = 120
MIN_EP1_WINDOW_MIN = 420          # 7 hours, matches "seven hour period for EP1, T1"


def check_gradient(raw_df: pd.DataFrame, config: ChannelConfig,
                    start: int, end: int, test_config: TestConfig) -> GradientResult:
    """Equivalent of Gradientcheck."""
    rsoc_hi_avg = raw_df[config.rsoc_hi_col].iloc[start:end + 1].mean()
    rsoc_lo_avg = raw_df[config.rsoc_lo_col].iloc[start:end + 1].mean()
    lsoc_hi_avg = raw_df[config.lsoc_hi_col].iloc[start:end + 1].mean()
    lsoc_lo_avg = raw_df[config.lsoc_lo_col].iloc[start:end + 1].mean()

    height_ft = test_config.height_adjusted_in / 12
    return GradientResult(
        rs_grad=(rsoc_hi_avg - rsoc_lo_avg) / height_ft,
        ls_grad=(lsoc_hi_avg - lsoc_lo_avg) / height_ft,
    )


def check_ambient(raw_df: pd.DataFrame, config: ChannelConfig,
                   start: int, end: int) -> AmbientResult:
    """Equivalent of AmbientCheck."""
    window = raw_df.iloc[start:end + 1]
    return AmbientResult(
        rsoc_min=window[config.rsoc_col].min(),
        rsoc_max=window[config.rsoc_col].max(),
        lsoc_min=window[config.lsoc_col].min(),
        lsoc_max=window[config.lsoc_col].max(),
    )


def check_stability(raw_df: pd.DataFrame, cycle_data: CycleData,
                     config: ChannelConfig, test_config: TestConfig) -> StabilityResult:
    """
    Dispatches to the cyclic or non-cyclic stability search, and to the
    FZ-only (freezer channels only) or full (FF+FZ channels) variant,
    matching StabilityCheck / StabilityCheckFZOnly's branching.
    """
    bw_channels = _select_channels(config, test_config.fz_only)

    if not cycle_data.cyclic:
        return _check_stability_fixed_window(raw_df, bw_channels, cycle_data, test_config)
    return _check_stability_cyclic(raw_df, bw_channels, cycle_data, test_config)


def _select_channels(config: ChannelConfig, fz_only: bool) -> list[str]:
    """FZOnly checks only freezer channels; otherwise FF + FZ."""
    fz_cols = [c for c in (config.fz_bw1_col, config.fz_bw2_col, config.fz_bw3_col,
                           config.fz_bw4_col, config.fz_bw5_col) if c is not None]
    if fz_only:
        return fz_cols
    ff_cols = [c for c in (config.ff_bw1_col, config.ff_bw2_col, config.ff_bw3_col) if c is not None]
    return ff_cols + fz_cols


def _check_stability_fixed_window(raw_df: pd.DataFrame, bw_channels: list[str],
                                   cycle_data: CycleData, test_config: TestConfig) -> StabilityResult:
    """
    Non-cyclic path: slides a 120-row window forward by Delta=300 rows,
    comparing averages, matching StabilityCheck's "cyclic = False" branch
    (and structurally identical to Run100's search in cycles.py).
    """
    delta = 300
    threshold = (delta / 60) * STABILITY_RATE_PER_HOUR
    start = cycle_data.cycles["index_on"].iloc[0] if len(cycle_data.cycles) else 1
    n = len(raw_df)

    r = start
    while r + delta + WINDOW_ROWS < n:
        window_a = raw_df[bw_channels].iloc[r: r + WINDOW_ROWS - 1]
        window_b = raw_df[bw_channels].iloc[r + delta: r + delta + WINDOW_ROWS - 1]

        diffs = (window_a.mean() - window_b.mean()).abs()
        if (diffs <= threshold).all():
            end_row = r + delta + WINDOW_ROWS
            test_duration = (raw_df["Timestamp"].iloc[end_row] - raw_df["Timestamp"].iloc[1]).total_seconds() / 60
            return StabilityResult(
                is_stable=True, stable_cycle=None, start_t1_index=None, end_t1_index=None,
                t1_minutes=None, ep1=None, test_duration_min=test_duration,
            )
        r += 1

    return StabilityResult(
        is_stable=False, stable_cycle=None, start_t1_index=None, end_t1_index=None,
        t1_minutes=None, ep1=None, test_duration_min=None,
        reason="Test did not reach stability within the available data",
    )


def _check_stability_cyclic(raw_df: pd.DataFrame, bw_channels: list[str],
                             cycle_data: CycleData, test_config: TestConfig) -> StabilityResult:
    """
    Cyclic path: grows three cycle-index gaps (J->K >= 120min, K->L >= 180min,
    L->M >= 120min) and compares channel averages between the J:K window and
    the L:M window, matching StabilityCheck's "cyclic = True" branch.

    Restarts the search (equivalent of VBA's GoTo ReCheckStability) if a
    defrost event is found inside the stability window before it was
    confirmed stable.
    """
    cycles = cycle_data.cycles
    index_on = cycles["index_on"].to_list()
    n_cycles = len(index_on)

    j = 0
    while j < n_cycles:
        try:
            j, k, l, m = _grow_stability_window(index_on, j)
        except _InsufficientDataError as e:
            return StabilityResult(
                is_stable=False, stable_cycle=None, start_t1_index=None,
                end_t1_index=None, t1_minutes=None, ep1=None, test_duration_min=None,
                reason=str(e),
            )

        window_jk = raw_df[bw_channels].iloc[index_on[j]: index_on[k] - 1]
        window_lm = raw_df[bw_channels].iloc[index_on[l]: index_on[m] - 1]

        mid1 = raw_df["Timestamp"].iloc[index_on[j]] + (raw_df["Timestamp"].iloc[index_on[k] - 1] - raw_df["Timestamp"].iloc[index_on[j]]) / 2
        mid2 = raw_df["Timestamp"].iloc[index_on[l]] + (raw_df["Timestamp"].iloc[index_on[m] - 1] - raw_df["Timestamp"].iloc[index_on[l]]) / 2
        delta_min = (mid2 - mid1).total_seconds() / 60
        threshold = (delta_min / 60) * STABILITY_RATE_PER_HOUR

        diffs = (window_jk.mean() - window_lm.mean()).abs()
        stable = bool((diffs <= threshold).all())

        if stable:
            # check for a defrost event inside [j, m) before declaring success
            defrost_idx = _find_defrost_before(raw_df, cycle_data, j, m)
            if defrost_idx is not None:
                logger.info("Defrost occurred before stability confirmed; restarting search at cycle %d", defrost_idx)
                j = defrost_idx  # equivalent of "GoTo ReCheckStability" with defcheckcounter advanced
                continue
            return _finalize_stability(raw_df, cycle_data, stable_cycle=m - 1)

        j += 1

    return StabilityResult(
        is_stable=False, stable_cycle=None, start_t1_index=None, end_t1_index=None,
        t1_minutes=None, ep1=None, test_duration_min=None,
        reason="Test is not stable (exhausted all cycles)",
    )


class _InsufficientDataError(Exception):
    pass


def _grow_stability_window(index_on: list[int], j: int) -> tuple[int, int, int, int]:
    """Finds k, l, m per the minimum-gap rules, raising if data runs out."""
    n = len(index_on)

    k = j
    while k < n and (index_on[k] - index_on[j]) < MIN_K_GAP_ROWS:
        k += 1
    if k >= n:
        raise _InsufficientDataError(f"Test not stable, aborting at cycle {j} (K search exhausted)")

    l = k
    while l < n and (index_on[l] - index_on[k]) < MIN_L_GAP_ROWS:
        l += 1
    if l >= n:
        raise _InsufficientDataError(f"Test not stable, aborting at cycle {j} (L search exhausted)")

    m = l
    while m < n and (index_on[m] - index_on[l]) < MIN_M_GAP_ROWS:
        m += 1
    if m >= n:
        raise _InsufficientDataError(f"Test not stable, aborting at cycle {j} (M search exhausted)")

    return j, k, l, m


def _find_defrost_before(raw_df: pd.DataFrame, cycle_data: CycleData, start_cycle: int, end_cycle: int) -> int | None:
    """Checks for a defrost event between two cycle boundaries."""
    cycles = cycle_data.cycles
    for xx in range(start_cycle, end_cycle):
        on_idx = cycles["index_on"].iloc[xx]
        off_idx = cycles["index_on"].iloc[xx + 1] if xx + 1 < len(cycles) else len(raw_df) - 1
        if (raw_df["Event 2 Defrost"].iloc[on_idx:off_idx] == 1).any():
            return xx + 1
    return None


def _finalize_stability(raw_df: pd.DataFrame, cycle_data: CycleData, stable_cycle: int) -> StabilityResult:
    """
    Finds the 7-hour (420 min) EP1/T1 window ending at the stable cycle,
    matching "Find seven hour period for EP1, T1".
    """
    cycles = cycle_data.cycles
    time_on = cycles["time_on"].to_list()
    watt_hours = cycles["watt_hours_at_on"].to_list()

    j = stable_cycle
    while j > 0 and (time_on[stable_cycle + 1] - time_on[j]).total_seconds() / 60 < MIN_EP1_WINDOW_MIN:
        j -= 1
    if j == stable_cycle:
        j -= 1

    ep1 = (watt_hours[stable_cycle + 1] - watt_hours[j]) / 1000
    t1_minutes = (time_on[stable_cycle + 1] - time_on[j]).total_seconds() / 60
    test_duration = (time_on[stable_cycle + 1] - raw_df["Timestamp"].iloc[1]).total_seconds() / 60

    return StabilityResult(
        is_stable=True,
        stable_cycle=stable_cycle,
        start_t1_index=j,
        end_t1_index=stable_cycle,
        t1_minutes=t1_minutes,
        ep1=ep1,
        test_duration_min=test_duration,
    )

