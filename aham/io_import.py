# io_import.py
import pandas as pd
import numpy as np
import io

def parse_thermal_csv(filepath):
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    # parse metadata block
    metadata = {}
    data_marker_idx = 0
    for i, line in enumerate(lines):
        if line.startswith("DATA::"):
            data_marker_idx = i
            n_header_rows = int(line.strip().split("::")[1])
            break
        if ":" in line:
            key, _, rest = line.partition(":")
            metadata[key.strip()] = rest.strip().strip(",").strip('"')

    # parse the 5-row header block
    header_start = data_marker_idx + 1
    header_rows = [
        [c.strip().strip('"') for c in lines[header_start + i].strip().split(",")]
        for i in range(n_header_rows)
    ]
    channel_ids, labels, units, scale, offset = header_rows

    # parse the data rows
    data_start = header_start + n_header_rows
    data_text = "".join(lines[data_start:])

    n_named_cols = len(labels)  # 27, including "Timestamp"
    df = pd.read_csv(
        io.StringIO(data_text),
        header=None,
        usecols=range(n_named_cols),   # drop the trailing empty column
        names=labels,
        na_values=["", " "],
    )

    # clean up timestamp & numeric columns 
    df["Timestamp"] = pd.to_datetime(
        df["Timestamp"].str.strip(), format="%d/%b/%Y %H:%M:%S" # use fixed C-level parser for speed
    )
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Replace logger's "no sensor" sentinel with NaN
    df.replace(-9.9e37, np.nan, inplace=True)

    return metadata, {"units": units, "scale": scale, "offset": offset}, df


metadata, channel_info, df = parse_thermal_csv(
    "data/7272_RAW-DATA___008_27JAN20261051_T25013475_001.csv"
)

# print(metadata["MODEL"], metadata["TEST_TITLE"])
# print(df.head())
# print(df.dtypes)
