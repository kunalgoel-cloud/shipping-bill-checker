"""
checker_core.py
----------------
All non-UI logic for the Shipping Bill Checker:
  - reading the courier working sheet (auto header detection)
  - reading the commercial agreement workbook (zone matrix + zone map + charges)
  - deriving zones from state names
  - running the approve / dispute checks
  - building the consolidated output workbook

Kept separate from app.py so the logic can be unit-tested / reused without
Streamlit running.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

APPOINTMENT = "Appointment based Delivery"
NON_APPOINTMENT = "Non Appointment based Delivery"

STATUS_APPROVED = "Approved"
STATUS_DISPUTED = "Disputed"
STATUS_NOT_CHECKED = "Not Checked"

# Column names we ask the user to map from their working sheet.
REQUIRED_FIELDS = {
    "awb": "AWB Number",
    "chargeable_weight": "Chargeable Weight (Billed)",
    "rate_per_kg": "Rate per KG (Billed)",
    "pickup_state": "Pickup State",
    "drop_state": "Drop State",
}

OPTIONAL_FIELDS = {
    "courier": "Courier Name",
    "customer": "Customer / Client",
    "drop_location": "Delivery Location (City)",
    "appointment_billed": "Appointment Charges (Billed amount)",
    "total_billed": "Total Charges (Billed)",
}

ALL_FIELDS = {**REQUIRED_FIELDS, **OPTIONAL_FIELDS}

# Fuzzy hints used to auto-suggest a column mapping from the uploaded sheet.
AUTO_HINTS = {
    "awb": ["awb"],
    "chargeable_weight": ["chargeable weight"],
    "rate_per_kg": ["rate per kg", "rate/kg", "per kg rate"],
    "pickup_state": ["pickup state", "origin state", "from state"],
    "drop_state": ["drop state", "destination state", "to state"],
    "courier": ["courier"],
    "customer": ["client name", "customer"],
    "drop_location": ["drop city", "destination city", "delivery city", "drop location"],
    "appointment_billed": ["appointment charges (billing)", "appointment charges billing"],
    "total_billed": ["total charges (billing)", "total charges"],
}


# --------------------------------------------------------------------------
# Working sheet loading
# --------------------------------------------------------------------------

def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    """Read a csv/xlsx/xls upload with tolerant header detection.

    Many courier working sheets have a summary/total row above the real
    header, so we scan the first 10 rows for the one most likely to be a
    header (the row with the most non-empty text-looking cells that isn't
    mostly numeric).
    """
    name = (uploaded_file.name or "").lower()
    raw_bytes = uploaded_file.read()
    buf = io.BytesIO(raw_bytes)

    if name.endswith(".csv"):
        preview = pd.read_csv(io.BytesIO(raw_bytes), header=None, nrows=10, dtype=str)
    else:
        preview = pd.read_excel(io.BytesIO(raw_bytes), header=None, nrows=10, dtype=str)

    header_row = _guess_header_row(preview)

    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw_bytes), header=header_row)
    else:
        df = pd.read_excel(io.BytesIO(raw_bytes), header=header_row)

    # Drop fully-empty columns/rows and normalise column names to strings.
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)
    return df


def _guess_header_row(preview: pd.DataFrame) -> int:
    best_row, best_score = 0, -1
    for i in range(len(preview)):
        row = preview.iloc[i]
        non_null = row.notna().sum()
        text_like = sum(
            1 for v in row if isinstance(v, str) and not _looks_numeric(v)
        )
        score = non_null + text_like  # reward rows with lots of text labels
        if score > best_score:
            best_score = score
            best_row = i
    return best_row


def _looks_numeric(v: str) -> bool:
    v = v.strip().replace(",", "")
    if v == "":
        return False
    try:
        float(v)
        return True
    except ValueError:
        return False


def suggest_column_mapping(columns: list[str]) -> dict[str, Optional[str]]:
    """Best-effort auto mapping from sheet columns to our internal field names."""
    norm_cols = {c: re.sub(r"\s+", " ", c.strip().lower()) for c in columns}
    mapping: dict[str, Optional[str]] = {}
    for field_key, hints in AUTO_HINTS.items():
        found = None
        for col, norm in norm_cols.items():
            if any(h in norm for h in hints):
                found = col
                break
        mapping[field_key] = found
    return mapping


# --------------------------------------------------------------------------
# Commercial agreement loading
# --------------------------------------------------------------------------

@dataclass
class Agreement:
    courier_name: str
    zone_matrix: pd.DataFrame  # index = from-zone, columns = to-zone, values = rate/kg
    zone_map: dict[str, str]  # normalised state/location -> zone
    charges: dict[str, float] = field(default_factory=dict)
    metro_locations: list[str] = field(default_factory=list)

    def rate_for(self, from_zone: str, to_zone: str) -> Optional[float]:
        try:
            return float(self.zone_matrix.loc[from_zone, to_zone])
        except (KeyError, ValueError, TypeError):
            return None

    def zone_for(self, state_or_city: Optional[str]) -> Optional[str]:
        if not state_or_city or pd.isna(state_or_city):
            return None
        key = _normalise(state_or_city)
        return self.zone_map.get(key)


def _normalise(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def load_agreement(uploaded_file) -> Agreement:
    """Parse the agreement workbook (see build_template() for expected layout)."""
    raw = uploaded_file.read()
    xls = pd.ExcelFile(io.BytesIO(raw))

    matrix_df = pd.read_excel(xls, sheet_name="ZoneMatrix", index_col=0)
    matrix_df.index = [str(i).strip() for i in matrix_df.index]
    matrix_df.columns = [str(c).strip() for c in matrix_df.columns]

    mapping_df = pd.read_excel(xls, sheet_name="ZoneMapping")
    zone_map: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        zone = str(row.get("Zone", "")).strip()
        locations = str(row.get("States / Locations (comma separated)", ""))
        if not zone or locations.lower() == "nan":
            continue
        for loc in locations.split(","):
            loc = loc.strip()
            if loc:
                zone_map[_normalise(loc)] = zone

    charges_df = pd.read_excel(xls, sheet_name="Charges")
    charges: dict[str, float] = {}
    metro_locations: list[str] = []
    for _, row in charges_df.iterrows():
        param = str(row.get("Parameter", "")).strip()
        value = row.get("Value", "")
        if param == "Metro Locations (comma separated)":
            metro_locations = [x.strip() for x in str(value).split(",") if x.strip()]
        elif param:
            try:
                charges[param] = float(value)
            except (ValueError, TypeError):
                pass

    courier_name = "Unknown"
    try:
        meta_df = pd.read_excel(xls, sheet_name="Charges", header=None, nrows=1)
    except Exception:
        pass

    return Agreement(
        courier_name=courier_name,
        zone_matrix=matrix_df,
        zone_map=zone_map,
        charges=charges,
        metro_locations=metro_locations,
    )


# --------------------------------------------------------------------------
# Default (Safexpress) template — used both to seed the downloadable
# template and as a ready-to-go sample agreement.
# --------------------------------------------------------------------------

_ZONES = ["North-1", "North-2", "East", "North-East", "West-1", "West-2",
          "South-1", "South-2", "Central"]

_MATRIX_VALUES = [
    [6.48, 6.48, 10.8, 16.2, 7.56, 8.64, 8.64, 10.8, 7.56],
    [6.48, 6.48, 10.8, 16.2, 8.64, 8.64, 10.8, 10.8, 7.56],
    [8.64, 10.8, 6.48, 7.56, 8.64, 10.8, 8.64, 10.8, 7.56],
    [8.64, 10.8, 7.56, 6.48, 10.8, 10.8, 10.8, 16.2, 8.64],
    [7.56, 8.64, 10.8, 16.2, 6.48, 6.48, 8.64, 10.8, 7.56],
    [8.64, 10.8, 10.8, 16.2, 6.48, 6.48, 7.56, 10.8, 7.56],
    [8.64, 10.8, 10.8, 16.2, 8.64, 7.56, 6.48, 7.56, 7.56],
    [10.8, 10.8, 10.8, 16.2, 8.64, 8.64, 6.48, 6.48, 7.56],
    [7.56, 8.64, 10.8, 16.2, 6.48, 7.56, 7.56, 10.8, 6.48],
]

_ZONE_LOCATIONS = {
    "North-1": "Delhi, Uttar Pradesh, Haryana, Rajasthan",
    "North-2": "Chandigarh, Punjab, Himachal Pradesh, Uttarakhand, Jammu & Kashmir, Ladakh",
    "East": "West Bengal, Odisha, Bihar, Jharkhand, Chhattisgarh",
    "North-East": "Assam, Meghalaya, Tripura, Arunachal Pradesh, Mizoram, Manipur, Nagaland, Sikkim",
    "West-1": "Gujarat, Daman & Diu, Dadra & Nagar Haveli",
    "West-2": "Maharashtra, Goa",
    "South-1": "Andhra Pradesh, Telangana, Karnataka, Tamil Nadu, Puducherry",
    "South-2": "Kerala",
    "Central": "Madhya Pradesh",
}

_DEFAULT_CHARGES = {
    "Docket Charge (Flat Rs)": 100,
    "FOV Percent (%)": 0.1,
    "FOV Minimum (Rs)": 100,
    "FSC Percent (%)": 20,
    "Min Chargeable Weight (Kg)": 15,
    "Min Chargeable Freight (Rs)": 400,
    "Metro Congestion Charge (Rs)": 100,
    "Appointment Charge Amount (Rs, 0 = any positive value counts)": 0,
}

_METRO_LOCATIONS = "Ahmedabad, Bengaluru, Chennai, Delhi, Hyderabad, Kolkata, Mumbai, Pune"


def build_template(prefill: bool = True) -> bytes:
    """Build the commercial-agreement Excel template as bytes.

    If prefill=True, the workbook is pre-populated with the Safexpress rate
    card supplied by the ops team, ready to use as-is or edit for another
    courier / lane.
    """
    matrix_df = pd.DataFrame(
        _MATRIX_VALUES if prefill else [[None] * len(_ZONES) for _ in _ZONES],
        index=_ZONES, columns=_ZONES,
    )
    matrix_df.index.name = "From \\ To"

    mapping_rows = [
        {"Zone": z, "States / Locations (comma separated)": _ZONE_LOCATIONS[z] if prefill else ""}
        for z in _ZONES
    ]
    mapping_df = pd.DataFrame(mapping_rows)

    charge_rows = [
        {"Parameter": k, "Value": (v if prefill else "")}
        for k, v in _DEFAULT_CHARGES.items()
    ]
    charge_rows.append({
        "Parameter": "Metro Locations (comma separated)",
        "Value": _METRO_LOCATIONS if prefill else "",
    })
    charges_df = pd.DataFrame(charge_rows)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        matrix_df.to_excel(writer, sheet_name="ZoneMatrix")
        mapping_df.to_excel(writer, sheet_name="ZoneMapping", index=False)
        charges_df.to_excel(writer, sheet_name="Charges", index=False)

        notes = pd.DataFrame({
            "Instructions": [
                "ZoneMatrix: per-kg rate from the zone in the row to the zone in the column.",
                "ZoneMapping: list every state / location that falls in each zone.",
                "Charges: flat/percentage charge parameters used by the optional extra checks.",
                "Replace the sample Safexpress figures with your own agreement, or use as-is.",
                "Do not rename the sheet tabs (ZoneMatrix / ZoneMapping / Charges) or the tool cannot read them.",
            ]
        })
        notes.to_excel(writer, sheet_name="Read Me", index=False)

    return out.getvalue()


# --------------------------------------------------------------------------
# Checking logic
# --------------------------------------------------------------------------

def _num(x) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        return float(x)
    except (ValueError, TypeError):
        return None


def run_checks(
    df: pd.DataFrame,
    mapping: dict[str, Optional[str]],
    agreement: Agreement,
    weight_tolerance: float = 0.0,
    rate_tolerance: float = 0.01,
    extra_checks: Optional[dict[str, bool]] = None,
) -> pd.DataFrame:
    """Return a new dataframe with computed/expected columns, per-check pass
    flags, an overall Status, and a Dispute Reasons string.

    `df` must already contain the editable "Ideal Weight (Kg)" and
    "Delivery Type" columns added by the app.
    """
    extra_checks = extra_checks or {}
    out = df.copy()

    def col(field_key):
        name = mapping.get(field_key)
        return out[name] if name and name in out.columns else pd.Series([None] * len(out))

    pickup_state = col("pickup_state")
    drop_state = col("drop_state")
    billed_weight = col("chargeable_weight").map(_num)
    billed_rate = col("rate_per_kg").map(_num)

    pickup_zone, drop_zone, expected_rate = [], [], []
    for ps, ds in zip(pickup_state, drop_state):
        pz = agreement.zone_for(ps)
        dz = agreement.zone_for(ds)
        pickup_zone.append(pz)
        drop_zone.append(dz)
        expected_rate.append(agreement.rate_for(pz, dz) if pz and dz else None)

    out["Computed Pickup Zone"] = pickup_zone
    out["Computed Drop Zone"] = drop_zone
    out["Agreement Rate/KG"] = expected_rate

    min_weight = agreement.charges.get("Min Chargeable Weight (Kg)")

    ideal_weight = out.get("Ideal Weight (Kg)", pd.Series([None] * len(out))).map(_num)
    if min_weight is not None:
        expected_weight = ideal_weight.map(
            lambda w: max(w, min_weight) if w is not None else None
        )
    else:
        expected_weight = ideal_weight
    out["Expected Chargeable Weight"] = expected_weight

    delivery_type = out.get("Delivery Type", pd.Series([None] * len(out)))
    appt_billed_col = mapping.get("appointment_billed")
    appt_billed = out[appt_billed_col].map(_num) if appt_billed_col and appt_billed_col in out.columns else pd.Series([None] * len(out))

    statuses, reasons = [], []
    rate_pass_l, weight_pass_l, appt_pass_l = [], [], []

    for i in range(len(out)):
        fail_reasons = []

        # --- Rate check ---
        er, br = expected_rate[i], billed_rate.iloc[i]
        if er is None:
            rate_pass = None  # not checked - zone/rate unavailable
        else:
            rate_pass = (br is not None) and (abs(br - er) <= rate_tolerance)
            if not rate_pass:
                fail_reasons.append(
                    f"Rate mismatch (billed {br if br is not None else 'NA'} vs agreement {er})"
                )
        rate_pass_l.append(rate_pass)

        # --- Weight check ---
        ew, bw = expected_weight.iloc[i], billed_weight.iloc[i]
        if ew is None:
            weight_pass = None  # ideal weight not entered yet
        else:
            weight_pass = (bw is not None) and (abs(bw - ew) <= weight_tolerance)
            if not weight_pass:
                fail_reasons.append(
                    f"Weight mismatch (billed {bw if bw is not None else 'NA'} vs expected {ew})"
                )
        weight_pass_l.append(weight_pass)

        # --- Appointment check ---
        dt = delivery_type.iloc[i] if i < len(delivery_type) else None
        if not dt or appt_billed_col is None:
            appt_pass = None  # not selected / no billed appointment column mapped
        else:
            ab = appt_billed.iloc[i]
            ab_val = ab if ab is not None else 0.0
            fixed_amt = agreement.charges.get(
                "Appointment Charge Amount (Rs, 0 = any positive value counts)"
            )
            if dt == APPOINTMENT:
                if fixed_amt:
                    appt_pass = abs(ab_val - fixed_amt) <= 0.01
                else:
                    appt_pass = ab_val > 0
            else:
                appt_pass = ab_val == 0
            if not appt_pass:
                fail_reasons.append(
                    f"Appointment charge mismatch (billed {ab_val}, expected {'>0' if dt == APPOINTMENT else '0'})"
                )
        appt_pass_l.append(appt_pass)

        # Overall status: a row can only be Approved once EVERY applicable
        # check has actually run and passed. A check that is still None
        # (couldn't run — e.g. no ideal weight entered yet, or no
        # appointment-billing column mapped) must never be silently treated
        # as "doesn't count" — otherwise a row could be marked Approved
        # while one of its checks was never actually verified.
        any_failed = len(fail_reasons) > 0
        any_pending = any(c is None for c in (rate_pass, weight_pass, appt_pass))
        if any_failed:
            statuses.append(STATUS_DISPUTED)
        elif any_pending:
            statuses.append(STATUS_NOT_CHECKED)
        else:
            statuses.append(STATUS_APPROVED)
        reasons.append("; ".join(fail_reasons) if fail_reasons else "")

    out["Rate Match"] = rate_pass_l
    out["Weight Match"] = weight_pass_l
    out["Appointment Match"] = appt_pass_l
    out["Status"] = statuses
    out["Dispute Reasons"] = reasons

    return out


def to_download_bytes(df: pd.DataFrame, fmt: str) -> bytes:
    if fmt == "csv":
        return df.to_csv(index=False).encode("utf-8")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All AWBs", index=False)
        if "Status" in df.columns:
            df[df["Status"] == STATUS_APPROVED].to_excel(writer, sheet_name="Approved", index=False)
            df[df["Status"] == STATUS_DISPUTED].to_excel(writer, sheet_name="Disputed", index=False)
    return out.getvalue()
