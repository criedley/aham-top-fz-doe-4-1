# models.py
from dataclasses import dataclass
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
class CycleBoundaries:
    time_on: list[pd.Timestamp]      # TimeCycleOn()
    time_off: list[pd.Timestamp]     # TimeCycleOff()
    index_on: list[int]              # IndexCycleOn()
    index_off: list[int]             # IndexCycleOff()
    watt_hours_at_on: list[float]    # WattHoursCount()
    cyclic: bool                     # cyclic flag
    last_scan_index: int             # LastScanIndex
    last_scan_time: pd.Timestamp     # LastScan

@dataclass
class CycleData:
    cycles: pd.DataFrame
    detection_method: str
    cyclic: bool
    last_scan_index: int
    last_scan_time: pd.Timestamp