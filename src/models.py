# src/aham/models.py
from dataclasses import dataclass
from datetime import datetime
import pandas as pd

@dataclass
class ImportResult:
    raw_df: pd.DataFrame        # the test data itself
    events_df: pd.DataFrame     # extracted event transitions
    test_info: dict             # header metadata (serial #, date, room, etc.)
    watts_col: str
    column_map: dict[str, str]
    
@dataclass
class ChannelConfig:
    ff_bw1_col: str
    ff_bw2_col: str
    ff_bw3_col: str
    fz_bw1_col: str
    fz_bw2_col: str
    fz_bw3_col: str
    fz_bw4_col: str | None
    fz_bw5_col: str | None
    ice_bw_col: str | None
    volts_col: str | None 
    rsoc_col: str
    lsoc_col: str
    rsoc_hi_col: str
    rsoc_lo_col: str
    lsoc_hi_col: str
    lsoc_lo_col: str
    unit_on_col: str
    defrost_on_col: str
    watts_col: str
    watt_hours_col: str

@dataclass
class TestConfig:
    # --- Cabinet / technician-entered physical parameters ---
    height_in: float              # hght, raw technician input (cabinet height, inches)
    height_adjusted_in: float     # hght + 10" t/c offset adjustment, computed not entered
    ct: float                     # CT, compressor "cycles per test" or similar AHAM param
    ff_ct: float = 0.0            # FFCT, legacy fresh-food CT (unused post-v1.6 per VBA comment)

    # --- Freezer-only test handling ---
    fz_only: bool = False
    fz_only_kfactor: float = 1.0  # must be one of {1, 0.85, 0.7} per original dropdown validation

    # --- Test start adjustment ---
    adjusted_start: bool = False
    time_to_start: datetime | None = None   # only meaningful if adjusted_start=True

    # --- Stability-check constants (AHAM standard, not technician-editable) ---
    stability_window_rows: int = 120        # the 119-row averaging window in StabilityCheck
    stability_delta_rows: int = 300         # Delta: row offset between window A and window B
    stability_rate_per_hour: float = 0.042  # degF/hr drift allowed, used in both StabilityCheck and Run100

    # --- Diagnostics / behavior flags ---
    debug_mode: bool = False

    def __post_init__(self):
        # only three K-factors are valid, per AHAM spec for different freezer-only configurations
        if self.fz_only and self.fz_only_kfactor not in (1.0, 0.85, 0.7):
            raise ValueError(
                f"fz_only_kfactor must be 1, 0.85, or 0.7 (per AHAM spec), got {self.fz_only_kfactor}"
            )
        if self.ct <= 0:
            raise ValueError("CT value must be entered and greater than zero")  # replaces CheckUserInput
        if self.height_in <= 0:
            raise ValueError("Unit height must be greater than zero")
        if self.adjusted_start and self.time_to_start is None:
            raise ValueError("adjusted_start=True requires time_to_start to be set")
        if self.fz_only_kfactor not in (1.0, 0.85, 0.7):
            raise ValueError(
                f"fz_only_kfactor must be 1, 0.85, or 0.7 (per AHAM spec), got {self.fz_only_kfactor}"
            )

    def gradient_limit(self) -> float:
        """Equivalent to GradientLimit = 0.5 * hght / 12 (feet)."""
        return 0.5 * self.height_adjusted_in / 12

    def stability_threshold(self, delta_rows: int | None = None) -> float:
        """Equivalent to Difference = (Delta / 60) * 0.042 in StabilityCheck."""
        delta = delta_rows if delta_rows is not None else self.stability_delta_rows
        return (delta / 60) * self.stability_rate_per_hour
    
@dataclass
class CycleData:
    cycles: pd.DataFrame
    detection_method: str
    cyclic: bool
    last_scan_index: int
    last_scan_time: pd.Timestamp

@dataclass
class GradientResult:
    rs_grad: float
    ls_grad: float

@dataclass
class AmbientResult:
    rsoc_min: float
    rsoc_max: float
    lsoc_min: float
    lsoc_max: float

@dataclass
class StabilityResult:
    is_stable: bool
    stable_cycle: int | None        # stabcycle
    start_t1_index: int | None      # StartT1 / J
    end_t1_index: int | None        # EndT1 / stabcycle
    t1_minutes: float | None        # t1
    ep1: float | None               # EP1, computed here since it's derived from the same window search
    test_duration_min: float | None # TestDuration
    reason: str | None = None       # replaces the MsgBox failure messages
    end_row_index: int | None = None
    end_row_index: int | None = None
    ff_avg: float | None = None     # FFAvg: mean of per-cycle FF channel averages over the T1 window
    fz_avg: float | None = None     # FZAvg: same for FZ channels

@dataclass
class EnergyResult:
    ep1: float       # kWh
    ep2: float       # kWh
    t1: float        # minutes
    t2: float        # minutes
    et: float        # kWh/day
    et_corr: float   # kWh/day, after the freezer-only K-factor