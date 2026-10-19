import sys
sys.path.insert(0, "src")

from channel_defaults import CHANNEL_MAP
from main import run_analysis
from models import TestConfig

cfg = TestConfig(height_in=68.5, height_adjusted_in=78.5, ct=10.0)
a = run_analysis("tests/fixtures/sample_raw.csv", CHANNEL_MAP, cfg)
c = a.cycle_data.cycles
h = cfg.height_adjusted_in / 12
print("cycles:", len(c))

pairs = {
    "Left Lo-Mid  (macro 0.043)": ("lsoc_avg", "lsoc_lo_avg"),
    "Left Mid-Hi  (macro 0.031)": ("lsoc_hi_avg", "lsoc_avg"),
    "Right Lo-Mid (macro 0.197)": ("rsoc_avg", "rsoc_lo_avg"),
    "Right Mid-Hi (macro 0.027)": ("rsoc_hi_avg", "rsoc_avg"),
}
for name, (upper, lower) in pairs.items():
    s = ((c[upper] - c[lower]) / h).round(3).abs()
    print(f"{name}: max[:-1]={s.iloc[:-1].max()}  max[:-2]={s.iloc[:-2].max()}  last 3={s.tail(3).tolist()}")