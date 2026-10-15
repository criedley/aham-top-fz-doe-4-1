"""
src/aham/io_import.py

Python equivalent of Module1.bas's ImportData / backupimport / CorrectNames.
Responsible for loading a raw data-cart CSV export, cleaning it, extracting
event transitions, normalizing column names, and returning a single
ImportResult object for the rest of the pipeline to consume.
"""

from __future__ import annotations

import logging
import io
from pathlib import Path

import pandas as pd
import numpy as np

from models import ImportResult

logger = logging.getLogger(__name__)

OPEN_TC_SENTINEL = -9.9e37           
MAX_HEADER_SCAN_ROWS = 50           # matches VBA's "If isel > 50" bail-out
MIN_REQUIRED_SCANS = 600            # matches RowCol's "insufficient scans" check
DUPLICATE_EVENT_WINDOW_S = 120      # matches the "DateDiff(...) < 120" dedup logic
UNEVEN_SCAN_GAP_MIN = 10            # matches "DateDiff('n', ...) > 10" check

REQUIRED_COLUMNS = [
    "Timestamp",
    "Watts",
    "Watt-Hours",
    "Event 1 Unit On",
    "Event 2 Defrost",
]

def load_datacart_csv(path: str | Path, alias_map: dict[str, list[str]],
                       on_missing_event_row: str = "approximate") -> ImportResult:
    metadata, header_rows, raw_df = _read_raw_csv(path)

    raw_df = _drop_leading_event_only_rows(raw_df, max_rows=5)
    raw_df = _drop_open_thermocouple_columns(raw_df)
    raw_df = _coerce_numeric(raw_df)

    _check_minimum_scans(raw_df)
    _check_uneven_scans(raw_df, on_missing_event_row)
    
    events_df = _extract_events(raw_df)
    events_df = _dedupe_events(events_df)

    raw_df = _correct_column_names(raw_df, alias_map)
    events_df = _correct_column_names(events_df, alias_map)

    _validate_required_columns(raw_df)

    return ImportResult(
        raw_df=raw_df,
        events_df=events_df,
        test_info=metadata,
        watts_col="Watts",
        column_map={c: c for c in raw_df.columns},
    )

# ---------------------------------------------------------------------------
# CSV structure parsing (replaces the "DATA::" marker scan)
# ---------------------------------------------------------------------------

def _read_raw_csv(path: str | Path) -> tuple[dict, int, pd.DataFrame]:
    """
    Two-phase parse. The metadata block above 'DATA::n' has a ragged
    (inconsistent) number of fields per line and cannot be loaded with
    pd.read_csv directly -- only the uniform data block below the header
    rows is loaded that way.
    """
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    metadata: dict[str, str] = {}
    marker_idx = None
    header_rows = 0

    for i, line in enumerate(lines):
        if line.startswith("DATA::"):
            marker_idx = i
            header_rows = int(line.strip().split("::")[1])
            break
        if ":" in line:
            key, _, rest = line.partition(":")
            metadata[key.strip()] = rest.strip().strip(",").strip('"')

    if marker_idx is None or marker_idx > MAX_HEADER_SCAN_ROWS:
        raise ImportError(
            "Data format is invalid: 'DATA::' marker not detected or header corrupted"
        )

    header_start = marker_idx + 1
    header_block = [
        [c.strip().strip('"') for c in lines[header_start + i].strip().split(",")]
        for i in range(header_rows)
    ]
    labels = header_block[1] if len(header_block) > 1 else header_block[0]
    n_named_cols = len(labels)

    data_start = header_start + header_rows
    data_text = "".join(lines[data_start:])

    raw = pd.read_csv(
        io.StringIO(data_text),
        header=None,
        usecols=range(n_named_cols),   # drop the trailing empty column
        names=labels,
        na_values=["", " "],
        dtype=str,
    )
    return metadata, header_rows, raw

# ---------------------------------------------------------------------------
# Cleanup steps
# ---------------------------------------------------------------------------

def _drop_leading_event_only_rows(df: pd.DataFrame, max_rows: int = 5) -> pd.DataFrame:
    """
    Matches:
        isel = 2
        Do While Cells(isel, 2) = "" And isel <= 5
            Rows(isel).Delete
        Loop
    i.e. discard any event-only rows (blank second column) at the very
    start of the file, up to a small limit.
    """
    second_col = df.columns[1]
    drop_idx = []
    for i in range(min(max_rows, len(df))):
        if pd.isna(df.iloc[i][second_col]) or df.iloc[i][second_col] == "":
            drop_idx.append(i)
        else:
            break
    return df.drop(index=drop_idx).reset_index(drop=True)

def _drop_open_thermocouple_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Matches ImportData's open-TC column/row removal:
        - drop columns that are entirely sentinel/empty in the first few rows
        - drop any row where a surviving column reports the sentinel value
    """
    keep_cols = [df.columns[0]]  # always keep Timestamp
    for col in df.columns[1:]:
        numeric = pd.to_numeric(df[col], errors="coerce")
        is_open = ((numeric == OPEN_TC_SENTINEL) | df[col].isna() | (df[col] == "")).all()
        if not is_open:
            keep_cols.append(col)
    return df[keep_cols]

def _coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    """Parse timestamp column and coerce all other columns to numeric where possible."""
    df = df.copy()
    ts_col = df.columns[0]
    df[ts_col] = pd.to_datetime(df[ts_col].astype(str).str.strip(), errors="coerce")
    df = df.rename(columns={ts_col: "Timestamp"})

    for col in df.columns[1:]:
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().any():
            # at least some values parsed -> this is a numeric column,
            # keep it with NaN in place of any stray non-numeric entries
            df[col] = converted
        else:
            # nothing parsed -> genuinely non-numeric column (e.g. "Edited"
            # marker text), leave the original values untouched
            logger.debug("Column '%s' left as non-numeric (no values parsed)", col)

    return df

def _check_minimum_scans(df: pd.DataFrame) -> None:
    """Equivalent of RowCol's 'insufficient scans' check."""
    if len(df) < MIN_REQUIRED_SCANS:
        raise ImportError(
            f"Insufficient scans: {len(df)} rows found, {MIN_REQUIRED_SCANS} minimum required"
        )

def _check_uneven_scans(df: pd.DataFrame, on_missing_event_row: str) -> None:
    """
    Equivalent of the 'detect uneven scans (more than 2 min apart)' check
    near Skip100Run. Replaces the interactive MsgBox Yes/No with an explicit
    mode flag, since this pipeline has no UI to prompt.
    """
    gaps = df["Timestamp"].diff().dt.total_seconds() / 60
    bad = gaps[gaps > UNEVEN_SCAN_GAP_MIN]
    if not bad.empty:
        msg = f"Inconsistent scan intervals found at {len(bad)} point(s); largest gap {bad.max():.1f} min"
        if on_missing_event_row == "raise":
            raise ImportError(msg)
        logger.warning(msg)


# ---------------------------------------------------------------------------
# Event extraction (diff-based -- see prior discussion on why this replaces
# the VBA's "insert synthetic row for missed transition" workaround)
# ---------------------------------------------------------------------------

def _event_state_columns(df: pd.DataFrame) -> list[str]:
    """
    Columns representing actual event STATE (e.g. "Event 1 Unit On",
    "Event 2 Defrost", "Event 3 Comp") -- deliberately excludes the
    "Event Count N ..." columns, which are monotonically-incrementing
    counters, not state flags, and must never be used to judge whether
    two event rows represent the same state.
    """
    return [c for c in df.columns if c.lower().startswith("event") and "count" not in c.lower()]

def _extract_events(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detects every row where a state column changes value, using diff()
    rather than scanning for blank-marker rows. Naturally handles both
    dedicated event rows emitted by the data cart and transitions that
    happen to fall between two regular scans.
    """
    state_cols = _event_state_columns(df)
    if not state_cols:
        return df.iloc[0:0]
    changed = (df[state_cols].diff().fillna(0) != 0).any(axis=1)
    return df.loc[changed].reset_index(drop=True)


def _dedupe_events(events_df: pd.DataFrame) -> pd.DataFrame:
    """
    Matches the false/duplicate-event removal logic: events within
    DUPLICATE_EVENT_WINDOW_S seconds of each other that don't represent a
    genuine state change are collapsed to one.
    """
    if events_df.empty:
        return events_df

    state_cols = _event_state_columns(events_df)
    keep = [0]
    for i in range(1, len(events_df)):
        prev_idx = keep[-1]
        gap_s = (events_df["Timestamp"].iloc[i] - events_df["Timestamp"].iloc[prev_idx]).total_seconds()
        same_state = events_df.iloc[i][state_cols].equals(events_df.iloc[prev_idx][state_cols])
        if gap_s < DUPLICATE_EVENT_WINDOW_S and same_state:
            continue
        keep.append(i)

    return events_df.iloc[keep].reset_index(drop=True)


# ---------------------------------------------------------------------------
# CorrectNames equivalent
# ---------------------------------------------------------------------------

def _correct_column_names(df: pd.DataFrame, alias_map: dict[str, list[str]]) -> pd.DataFrame:
    """
    Normalizes column headers to canonical names using a known-aliases
    table, replacing the old 'ParameterNames' sheet lookup.
    """
    canonical_names = set(alias_map.keys())
    rename_map = {}
    for col in df.columns:
        if col in canonical_names or col == "Timestamp":
            continue
        for canonical, aliases in alias_map.items():
            if col in aliases:
                rename_map[col] = canonical
                break
        else:
            logger.warning("Column '%s' did not match any known canonical name or alias", col)
    return df.rename(columns=rename_map)


def _validate_required_columns(df: pd.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ImportError(f"Raw data missing required columns after normalization: {missing}")
