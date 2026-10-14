# scratch_run.py — run this from the project root
import logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s [%(name)s] %(message)s")

import sys
sys.path.insert(0, "src")

from io_import import load_datacart_csv
from setup import build_channel_config
from cycles import detect_cycles
from models import TestConfig

# Normalizes raw header spelling -> the exact names io_import.py's
# _validate_required_columns and REQUIRED_COLUMNS expect.
ALIAS_MAP = {
    "FF BW 1": ["FF BW1"],
    "FF BW 2": ["FF BW2"],
    "FF BW 3": ["FF BW3"],
    "FZ BW 1": ["FZ BW1"],
    "FZ BW 2": ["FZ BW2"],
    "FZ BW 3": ["FZ BW3"],
    "RSOC HI": ["RSOC UP"],
    "RSOC Mid": ["RSOC MD"],
    "RSOC LO": ["RSOC LO"],
    "LSOC HI": ["LSOC UP"],
    "LSOC Mid": ["LSOC MD"],
    "LSOC LO": ["LSOC LO"],
    "Watts": ["Watts"],
    "Watt-Hours": ["Watt-Hours"],
    "Event 1 Unit On": ["Event 1 Unit On"],
    "Event 2 Defrost": ["Event 2 Defrost"],
    "Event 3 Comp": ["Event 3 Comp"],
}

# Bridges REQUIRED_ROLES (role names) -> the now-normalized column names above.
CHANNEL_MAP = {
    "Watts": "Watts",
    "Watt-Hours": "Watt-Hours",
    "Unit On": "Event 1 Unit On",
    "Defrost On": "Event 2 Defrost",
    "RSOC Mid": "RSOC Mid",
    "LSOC Mid": "LSOC Mid",
    "RSOC HI": "RSOC HI",
    "RSOC LO": "RSOC LO",
    "LSOC HI": "LSOC HI",
    "LSOC LO": "LSOC LO",
    "FF BW 1": "FF BW 1",
    "FF BW 2": "FF BW 2",
    "FF BW 3": "FF BW 3",
    "FZ BW 1": "FZ BW 1",
    "FZ BW 2": "FZ BW 2",
    "FZ BW 3": "FZ BW 3",
}

TEST_CONFIG = TestConfig(
    height_in=55.0,
    height_adjusted_in=65.0,
    ct=96.0,
    fz_only=False,
    fz_only_kfactor=1.0,
    adjusted_start=False,
)

CSV_PATH = "data/7272_RAW-DATA___008_27JAN20261051_T25013475_001.csv"  # adjust if needed

print("=" * 70)
print("STEP 1: load_datacart_csv")
print("=" * 70)
import_result = load_datacart_csv(CSV_PATH, ALIAS_MAP)
print(f"raw_df shape: {import_result.raw_df.shape}")
print(f"raw_df columns: {list(import_result.raw_df.columns)}")
print(f"events_df shape: {import_result.events_df.shape}")
print(import_result.raw_df[["Timestamp", "Watts", "Event 1 Unit On", "Event 2 Defrost"]].head(10))

print("\n" + "=" * 70)
print("STEP 2: build_channel_config")
print("=" * 70)
config = build_channel_config(import_result, CHANNEL_MAP)
print(config)

print("\n" + "=" * 70)
print("STEP 3: detect_cycles")
print("=" * 70)
cycle_data = detect_cycles(import_result, config, TEST_CONFIG)
print(f"detection_method: {cycle_data.detection_method}")
print(f"cyclic: {cycle_data.cyclic}")
print(f"number of cycles detected: {len(cycle_data.cycles)}")
print(cycle_data.cycles[["cycle_num", "time_on", "time_off", "index_on", "index_off", "duty_cycle"]].to_string())

print("\n" + "=" * 70)
print("STEP 4: spot-check known transitions")
print("=" * 70)
known_transitions = [
    "2026-01-27 17:16:20",
    "2026-01-27 18:42:47",
    "2026-01-28 23:23:53",
    "2026-01-29 08:57:28",
]
import pandas as pd
for ts_str in known_transitions:
    ts = pd.Timestamp(ts_str)
    nearby_on = cycle_data.cycles[(cycle_data.cycles["time_on"] - ts).abs().dt.total_seconds() < 120]
    nearby_off = cycle_data.cycles[(cycle_data.cycles["time_off"] - ts).abs().dt.total_seconds() < 120]
    match = "ON match" if not nearby_on.empty else ("OFF match" if not nearby_off.empty else "NO MATCH")
    print(f"{ts_str}: {match}")