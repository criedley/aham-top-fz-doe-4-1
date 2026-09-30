# io_import.py
import csv
import numpy as np
import pandas as pd

def load_datacart_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as file:
        rows = list(csv.reader(file))

    marker_idx = next(
        i for i, row in enumerate(rows)
        if row and row[0].startswith("DATA::")
    )
    n_header_rows = int(rows[marker_idx][0].split("::")[1])

    metadata = [
        {"key": row[0].rstrip(":"), "values": row[1:]}
        for row in rows[:marker_idx]
        if row
    ]
    metadata_df = pd.DataFrame(metadata)

    channel_names = rows[marker_idx + 2]
    data_start = marker_idx + n_header_rows + 1
    data_df = pd.DataFrame([row[:len(channel_names)] for row in rows[data_start:]], columns=channel_names)
    # data_df = data_df.loc[:, ~(data_df.iloc[0] == -9.9e37)]  # drop open-TC columns

    return metadata_df, data_df

path = "7272_RAW-DATA___008_27JAN20261051_T25013475_001.csv"
df = load_datacart_csv(path)[1]
print(df)