# tests/factories.py
import pandas as pd

from models import CycleData

def make_perfectly_stable_raw_df(n_rows: int = 2000) -> pd.DataFrame:
    """
    Builds a synthetic raw_df where every BW channel holds a constant
    value (zero drift) -- this should always be judged "stable" by any
    correct implementation, regardless of window size or threshold,
    since the difference between any two windows is exactly 0.0.
    """
    timestamps = pd.date_range("2026-01-27 00:00:00", periods=n_rows, freq="min")
    return pd.DataFrame({
        "Timestamp": timestamps,
        "FF BW 1": [38.0] * n_rows, "FF BW 2": [38.0] * n_rows, "FF BW 3": [38.0] * n_rows,
        "FZ BW 1": [0.0] * n_rows, "FZ BW 2": [0.0] * n_rows, "FZ BW 3": [0.0] * n_rows,
        "RSOC HI": [90.5] * n_rows, "RSOC LO": [90.0] * n_rows, "RSOC Mid": [90.25] * n_rows,
        "LSOC HI": [90.5] * n_rows, "LSOC LO": [90.0] * n_rows, "LSOC Mid": [90.25] * n_rows,
        "Event 2 Defrost": [0] * n_rows,
        "Watt-Hours": list(range(n_rows)),
    })

def make_cycle_data(
    index_on: list[int],
    cyclic: bool = True,
    ff_avg: float = 38.0,
    fz_avg: float = 0.0,
) -> CycleData:
    """Minimal synthetic CycleData -- just enough for the stability search to run.

    Per-cycle channel averages are constant, mirroring the columns that
    cycles._compute_cycle_stats produces. fz_bw4/fz_bw5 are 0.0, as in
    production for unconfigured channels.
    """
    base_time = pd.Timestamp("2026-01-27 00:00:00")
    cycles_df = pd.DataFrame({
        "index_on": index_on,
        "time_on": [base_time + pd.Timedelta(minutes=i) for i in index_on],
        "watt_hours_at_on": [float(i) for i in index_on],
    })
    for col in ("ff_bw1", "ff_bw2", "ff_bw3"):
        cycles_df[col] = ff_avg
    for col in ("fz_bw1", "fz_bw2", "fz_bw3"):
        cycles_df[col] = fz_avg
    for col in ("fz_bw4", "fz_bw5"):
        cycles_df[col] = 0.0
    return CycleData(
        cycles=cycles_df, detection_method="event_based", cyclic=cyclic,
        last_scan_index=index_on[-1] + 100,
        last_scan_time=base_time + pd.Timedelta(minutes=index_on[-1] + 100),
    )

def make_never_stable_raw_df(n_rows: int = 2000) -> pd.DataFrame:
    """
    Builds a synthetic raw_df where FZ BW 1 drifts linearly and never
    plateaus -- this should never be judged stable, since the
    window-to-window difference always exceeds the AHAM threshold.
    """
    timestamps = pd.date_range("2026-01-27 00:00:00", periods=n_rows, freq="min")
    drift = [i * 0.01 for i in range(n_rows)]  # far exceeds 0.042/hr
    return pd.DataFrame({
        "Timestamp": timestamps,
        "FF BW 1": [38.0] * n_rows, "FF BW 2": [38.0] * n_rows, "FF BW 3": [38.0] * n_rows,
        "FZ BW 1": drift, "FZ BW 2": [0.0] * n_rows, "FZ BW 3": [0.0] * n_rows,
        "RSOC HI": [90.5] * n_rows, "RSOC LO": [90.0] * n_rows, "RSOC Mid": [90.25] * n_rows,
        "LSOC HI": [90.5] * n_rows, "LSOC LO": [90.0] * n_rows, "LSOC Mid": [90.25] * n_rows,
        "Event 2 Defrost": [0] * n_rows,
        "Watt-Hours": list(range(n_rows)),
    })