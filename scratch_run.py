# scratch_run.py — run this from the project root
import pandas as pd
import logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s [%(name)s] %(message)s")

import sys
sys.path.insert(0, "src")

from io_import import load_datacart_csv
from setup import build_channel_config
from cycles import detect_cycles
from models import TestConfig
from checks import check_stability
from channel_defaults import ALIAS_MAP, CHANNEL_MAP

TEST_CONFIG = TestConfig(
    height_in=71.0,
    height_adjusted_in=81.0,
    ct=10.0,
    fz_only=False,
    fz_only_kfactor=1.0,
    adjusted_start=False,
)

CSV_PATH = "tests/fixtures/sample_raw.csv"  # adjust if needed

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
for ts_str in known_transitions:
    ts = pd.Timestamp(ts_str)
    nearby_on = cycle_data.cycles[(cycle_data.cycles["time_on"] - ts).abs().dt.total_seconds() < 120]
    nearby_off = cycle_data.cycles[(cycle_data.cycles["time_off"] - ts).abs().dt.total_seconds() < 120]
    match = "ON match" if not nearby_on.empty else ("OFF match" if not nearby_off.empty else "NO MATCH")
    print(f"{ts_str}: {match}")

print("\n" + "=" * 70)
print("STEP 5: Check stability")
print("=" * 70)
result = check_stability(import_result.raw_df, cycle_data, config, TEST_CONFIG)
print(result)