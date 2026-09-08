"""
cloud_backup.py
----------------
Optional Supabase-backed persistence for the per-AWB review workflow.

The app works completely fine without this configured — it just falls
back to the original session-only behaviour (everything disappears on
refresh). When Streamlit secrets contain SUPABASE_URL and SUPABASE_KEY,
every Calculate / Approve / Reject / Remarks action for an AWB is also
upserted into a Supabase table (keyed by AWB), and re-uploading the same
working sheet after a refresh pulls that saved state back in — so the ops
team's decisions and remarks survive a refresh even though nothing else
does (the raw working sheet and commercial agreement are still session-only,
same as before; only the per-AWB review state is backed up).

Kept separate from app.py so it can be unit-tested / reused, and separate
from checker_core.py so that file stays free of any network dependency.
"""

from __future__ import annotations

import math

TABLE_NAME = "awb_workflow"

# Fields from the app's per-AWB workflow dict that get persisted. Anything
# not listed here (the in-memory "calculated" breakdown, "calc_inputs")
# stays local — it's cheap to recompute and doesn't need to round-trip.
PERSISTED_FIELDS = [
    "ideal_weight", "delivery_type", "decision", "final_price",
    "price_source", "status", "reason", "remarks",
]


def get_client(url: str, key: str):
    """Create a Supabase client. Raises if the `supabase` package or the
    given credentials are invalid — callers should catch and degrade
    gracefully rather than let this crash the app."""
    from supabase import create_client
    return create_client(url, key)


def _native(v):
    """Coerce pandas/numpy scalars to plain JSON-serializable Python types.
    (numpy.int64/float64 etc. aren't accepted by the standard json encoder
    the Supabase client uses under the hood, and pandas NaN needs to become
    None, not the float nan.)"""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if hasattr(v, "item"):  # numpy scalar (int64, float64, bool_, ...)
        v = v.item()
        if isinstance(v, float) and math.isnan(v):
            return None
        return v
    if isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def save_awb_state(client, awb: str, context: dict, workflow: dict) -> None:
    """Upsert one AWB's current review state.

    `context` is read-only descriptive info from the working sheet (courier,
    consignee, states, ...) kept alongside the workflow fields purely so the
    Supabase table is readable on its own, without needing to cross-reference
    the original working sheet.
    """
    row = {"awb": str(awb)}
    for k, v in context.items():
        row[k] = _native(v)
    for f in PERSISTED_FIELDS:
        row[f] = _native(workflow.get(f))
    client.table(TABLE_NAME).upsert(row, on_conflict="awb").execute()


def load_awb_states(client, awbs: list[str]) -> dict[str, dict]:
    """Fetch saved state for the given AWBs. Returns {awb: {field: value}}
    for whichever of them have a saved row; AWBs with no saved row are
    simply absent from the result (nothing to resume)."""
    if not awbs:
        return {}
    resp = client.table(TABLE_NAME).select("*").in_("awb", [str(a) for a in awbs]).execute()
    return {row["awb"]: row for row in (resp.data or [])}


def build_context(row, mapping: dict) -> dict:
    """The descriptive (non-workflow) columns saved alongside each AWB."""
    return {
        "courier": row[mapping["courier"]] if mapping.get("courier") else None,
        "consignee": row[mapping["consignee"]] if mapping.get("consignee") else None,
        "pickup_state": row[mapping["pickup_state"]],
        "drop_state": row[mapping["drop_state"]],
        "delivery_location": row[mapping["drop_location"]],
    }
