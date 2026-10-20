import sys
sys.path.insert(0, "src")
sys.path.insert(0, "tests")

import pandas as pd
from channel_defaults import CHANNEL_MAP
from main import run_analysis
from models import TestConfig
from report import build_report
from test_t2_table import EXPECTED

pd.set_option("display.width", 200)
pd.set_option("display.max_columns", 20)

cfg = TestConfig(height_in=71.0, height_adjusted_in=81.0, ct=10.0)
a = run_analysis("tests/fixtures/sample_raw.csv", {**CHANNEL_MAP, "Volts": "Volts"}, cfg)
ch = a.channel_config
raw = a.import_result.raw_df
ev = a.import_result.events_df

# 1. Event rows around the defrost end, with seconds and counters
print("EVENT ROWS 17:38 - 18:07")
show = ["Timestamp", ch.watt_hours_col, ch.unit_on_col, ch.defrost_on_col,
        "Event Count 1 Unit On", "Event Count 2 Defrost"]
show = [c for c in show if c in ev.columns]
print(ev[(ev["Timestamp"] >= "2026-01-28 17:38") & (ev["Timestamp"] <= "2026-01-28 18:07")][show].to_string())

# 2. Scan rows in the same period (seconds of each scan, and state at each scan)
scan = raw[raw[ch.watts_col].notna()].reset_index(drop=True)
print("\nSCAN ROWS 17:38 - 17:58")
print(scan[(scan["Timestamp"] >= "2026-01-28 17:38") & (scan["Timestamp"] <= "2026-01-28 17:58")]
      [["Timestamp", ch.unit_on_col, ch.defrost_on_col, ch.lsoc_col]].to_string())

# 3. Every column of the Defrost and Recovery rows against the macro
t2 = build_report(a, cfg).t2_table.set_index("label")
def expected_flat(e):
    out = {}
    for name, vals in (("ff_bw", e["ff"]), ("fz_bw", e["fz"])):
        for i, v in enumerate(vals, 1):
            out[f"{name}{i}"] = v
    out.update(volts_avg=e["volts"][0], volts_min=e["volts"][1], volts_max=e["volts"][2],
               lsoc_avg=e["lsoc"], rsoc_avg=e["rsoc"], l_grad=e["grad"][0], r_grad=e["grad"][1],
               duration_min=e["duration"], on_min=e["on"], off_min=e["off"],
               runtime=e["runtime"], wh=e["wh"])
    return out

for label in ("Defrost", "Recovery"):
    print(f"\n{label}: ours vs macro")
    for k, want in expected_flat(EXPECTED[label]).items():
        got = float(t2.loc[label, k])
        flag = "" if abs(got - want) <= 0.0051 else "   <-- DIFF"
        print(f"  {k:13s} ours={got:<10.4f} macro={want:<10.4f}{flag}")
    print("  start", t2.loc[label, "start"], " end", t2.loc[label, "end"])

# 4. Unrounded Defrost means, and the effect of moving the window by one scan
d = a.defrost
lb = lambda t: int(scan["Timestamp"].searchsorted(t))
lo, hi = lb(d.defrost_start), lb(d.segments["Defrost"].end) + 1
print("\nDEFROST WINDOW (macro: lsoc 90.42, ff1 40.12, fz1 2.96)")
for dlo, dhi in [(0, 0), (-1, 0), (1, 0), (0, -1), (0, 1)]:
    w = scan.iloc[lo + dlo: hi + dhi]
    print(f"  lo{dlo:+d} hi{dhi:+d}  n={len(w)}  lsoc={w[ch.lsoc_col].mean():.6f}"
          f"  ff1={w[ch.ff_bw1_col].mean():.4f}  fz1={w[ch.fz_bw1_col].mean():.4f}")