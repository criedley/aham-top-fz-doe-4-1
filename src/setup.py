"""
setup.py

Python equivalent of Setup.bas's SetupVars/AvailChannel/RowCol.
Resolves technician-assigned channel roles (Watts, FF BW 1, RSOC HI, etc.)
to actual column names in the imported raw data, and validates that the
assignment is complete and consistent with what was actually imported.
"""

from __future__ import annotations

import yaml
import logging
from pathlib import Path

from models import ImportResult, ChannelConfig

logger = logging.getLogger(__name__)

MIN_REQUIRED_SCANS = 600   # matches RowCol's "insufficient scans" check

# Logical roles that MUST be assigned to a real column -- no "NA" allowed.
# Matches CheckUserInput's blank-field validation.
REQUIRED_ROLES = [
    "Watts", "Watt-Hours", "Unit On", "Defrost On",
    "RSOC Mid", "LSOC Mid", "RSOC HI", "RSOC LO", "LSOC HI", "LSOC LO",
    "FF BW 1", "FF BW 2", "FF BW 3",
    "FZ BW 1", "FZ BW 2", "FZ BW 3",
]

# Roles that are allowed to be "NA" (optional sensors), matching the
# dropdowns in AddDataSheet that default to "NA" when not present.
OPTIONAL_ROLES = ["FZ BW 4", "FZ BW 5", "ICE BW", "Volts"]


def build_channel_config(import_result: ImportResult,
                          technician_channel_map: dict[str, str]) -> ChannelConfig:
    available = set(import_result.raw_df.columns)

    _check_minimum_scans(import_result.raw_df)
    _check_complete_assignment(technician_channel_map)

    def resolve_required(role: str) -> str:
        assigned = technician_channel_map.get(role)
        if assigned is None or assigned == "NA":
            raise ValueError(f"Required channel role '{role}' was not assigned")
        if assigned not in available:
            raise ValueError(
                f"Channel '{assigned}' assigned to role '{role}' not found in raw data. "
                f"Available columns: {sorted(available)}"
            )
        return assigned

    def resolve_optional(role: str) -> str | None:
        assigned = technician_channel_map.get(role)
        if assigned is None or assigned == "NA":
            return None
        if assigned not in available:
            raise ValueError(
                f"Channel '{assigned}' assigned to role '{role}' not found in raw data. "
                f"Available columns: {sorted(available)}"
            )
        return assigned

    return ChannelConfig(
        watts_col=resolve_required("Watts"),
        watt_hours_col=resolve_required("Watt-Hours"),
        unit_on_col=resolve_required("Unit On"),
        defrost_on_col=resolve_required("Defrost On"),
        rsoc_col=resolve_required("RSOC Mid"),
        lsoc_col=resolve_required("LSOC Mid"),
        rsoc_hi_col=resolve_required("RSOC HI"),
        rsoc_lo_col=resolve_required("RSOC LO"),
        lsoc_hi_col=resolve_required("LSOC HI"),
        lsoc_lo_col=resolve_required("LSOC LO"),
        ff_bw1_col=resolve_required("FF BW 1"),
        ff_bw2_col=resolve_required("FF BW 2"),
        ff_bw3_col=resolve_required("FF BW 3"),
        fz_bw1_col=resolve_required("FZ BW 1"),
        fz_bw2_col=resolve_required("FZ BW 2"),
        fz_bw3_col=resolve_required("FZ BW 3"),
        fz_bw4_col=resolve_optional("FZ BW 4"),
        fz_bw5_col=resolve_optional("FZ BW 5"),
        ice_bw_col=resolve_optional("ICE BW"),
        volts_col=resolve_optional("Volts"),
    )


def _check_minimum_scans(raw_df) -> None:
    """Equivalent of RowCol's 'insufficient scans' MsgBox + ExitFLG."""
    if len(raw_df) < MIN_REQUIRED_SCANS:
        raise ValueError(
            f"Insufficient scans: {len(raw_df)} rows found, "
            f"{MIN_REQUIRED_SCANS} minimum required (increase test duration)"
        )


def _check_complete_assignment(technician_channel_map: dict[str, str]) -> None:
    """
    Equivalent of CheckUserInput's loop:
        For I = 1 To 21
            If Cells(I, 2).Value = "" Then
                MsgBox ("Please select proper channel or NA if channel is not available")
    Fails fast and lists every missing role at once, rather than one at a
    time the way the VBA's MsgBox loop effectively did.
    """
    missing = [role for role in REQUIRED_ROLES if not technician_channel_map.get(role)]
    if missing:
        raise ValueError(
            f"Missing required channel assignments: {missing}. "
            f"Please assign a column or 'NA' for optional roles."
        )

def load_technician_input(config_path: Path = Path("config/channel_map.yaml")) -> tuple[dict, dict]:
    """Reads the config written by the HTML form submission."""
    if not config_path.exists():
        raise FileNotFoundError(
            f"No technician configuration found at {config_path}. "
            f"Please complete the setup form first."
        )
    data = yaml.safe_load(config_path.read_text())
    return data["channel_map"], data["test_config"]