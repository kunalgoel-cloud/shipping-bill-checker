"""
Shipping Bill Checker
======================
Ops tool to audit courier working sheets (billed AWBs) against a commercial
rate agreement. Nothing is written to a database — everything lives in the
browser session and disappears when the tab is closed / refreshed.
"""

import numpy as np
import pandas as pd
import streamlit as st

import checker_core as core

st.set_page_config(page_title="Shipping Bill Checker", page_icon="\U0001F69A", layout="wide")

# --------------------------------------------------------------------------
# Session state defaults
# --------------------------------------------------------------------------
for key, default in {
    "working_df": None,
    "mapping": {},
    "agreement": None,
    "editable_df": None,
    "results_df": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# The Safexpress commercial agreement is bundled directly in checker_core.py
# (see default_agreement()) so the app is active and usable with zero setup.
# Uploading a different agreement workbook in the sidebar overrides this for
# the session; refreshing the tab resets back to this built-in default,
# same as everything else — nothing persists.
if st.session_state.agreement is None:
    st.session_state.agreement = core.default_agreement()

st.title("\U0001F69A Shipping Bill Checker")
st.caption(
    "Upload a courier working sheet, check billed AWBs against your commercial "
    "agreement, add expected weight & delivery type, and get an approve/dispute "
    "sheet — nothing is stored server-side."
)

# ==========================================================================
# SIDEBAR — Commercial agreement
# ==========================================================================
with st.sidebar:
    st.header("1. Commercial Agreement")

    active = st.session_state.agreement
    st.success(f"**Active: {active.courier_name}**\n\n{active.source}")
    st.caption("See the full rate card in the **Commercial Agreement** tab →")

    st.write(
        "Upload a different agreement workbook to override this for the "
        "session. Don't have one yet? Grab the template below — it comes "
        "pre-filled with the built-in Safexpress rate card, ready to "
        "overwrite for another courier."
    )

    st.download_button(
        "\U0001F4E5 Download agreement template (.xlsx)",
        data=core.build_template(prefill=True),
        file_name="commercial_agreement_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    agreement_file = st.file_uploader(
        "Upload commercial agreement (.xlsx)", type=["xlsx"], key="agreement_upload"
    )

    if agreement_file is not None:
        try:
            st.session_state.agreement = core.load_agreement(agreement_file)
            st.success(
                f"Agreement loaded — {len(st.session_state.agreement.zone_matrix)} zones, "
                f"{len(st.session_state.agreement.zone_map)} locations mapped."
            )
            st.rerun()
        except Exception as e:
            st.error(f"Could not read that agreement file: {e}")

    if active.source.startswith("Uploaded"):
        if st.button("↺ Reset to built-in Safexpress agreement", use_container_width=True):
            st.session_state.agreement = core.default_agreement()
            st.rerun()

    st.divider()
    st.caption(
        "Data handling: files are processed in memory for this session only. "
        "Closing or refreshing the tab clears everything — nothing is saved "
        "to a database or disk."
    )

# ==========================================================================
# TABS — Checker workflow vs. a read-only view of whatever agreement is active
# ==========================================================================
tab_checker, tab_agreement = st.tabs(["\U0001F50D Checker", "\U0001F4C4 Commercial Agreement"])

with tab_agreement:
    ag = st.session_state.agreement
    st.subheader(f"Active agreement: {ag.courier_name}")
    st.caption(
        f"Source: {ag.source}. This is exactly what every rate / weight / "
        "appointment check in the Checker tab runs against right now — "
        "upload a different agreement in the sidebar to change it."
    )

    st.markdown("#### Charges & Parameters")
    if ag.charges:
        charges_df = pd.DataFrame(
            [{"Parameter": k, "Value": v} for k, v in ag.charges.items()]
        )
        st.dataframe(charges_df, use_container_width=True, hide_index=True)
    else:
        st.info("No charge parameters loaded for this agreement.")

    if ag.metro_locations:
        st.markdown("**Metro locations (congestion charge applies):** " + ", ".join(ag.metro_locations))

    st.markdown("#### Zone Mapping — State / Location → Zone")
    if ag.zone_locations:
        zm_df = pd.DataFrame(
            [{"Zone": z, "States / Locations": ", ".join(locs)} for z, locs in ag.zone_locations.items()]
        )
        st.dataframe(zm_df, use_container_width=True, hide_index=True)
    else:
        st.info("No zone mapping loaded for this agreement.")

    st.markdown("#### Zone-to-Zone Rate Matrix (₹ per Kg)")
    st.dataframe(ag.zone_matrix, use_container_width=True)
    st.caption(
        "Row = pickup zone, column = drop zone — direction matters, the "
        "matrix is not symmetric. Every AWB's expected rate is looked up "
        "here using the zone derived from its Pickup/Drop State, never "
        "from any zone column the working sheet itself claims."
    )

with tab_checker:
    # ======================================================================
    # STEP 1 — Upload working sheet
    # ======================================================================
    st.header("2. Upload Working Sheet")
    working_file = st.file_uploader(
        "Courier working sheet (.xlsx, .xls or .csv)", type=["xlsx", "xls", "csv"], key="working_upload"
    )

    if working_file is not None:
        try:
            df = core.read_uploaded_table(working_file)
            st.session_state.working_df = df
            st.success(f"Loaded {len(df)} rows, {len(df.columns)} columns from **{working_file.name}**.")
        except Exception as e:
            st.error(f"Could not read that working sheet: {e}")

    df = st.session_state.working_df

    if df is None:
        st.info("Upload a working sheet above to get started.")
        st.stop()

    with st.expander("Preview raw working sheet", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)

    # ======================================================================
    # STEP 2 — Column mapping
    # ======================================================================
    st.header("3. Map Columns")
    st.caption(
        "We've guessed which column is which — check them, especially the "
        "required ones, before continuing."
    )

    if not st.session_state.mapping:
        st.session_state.mapping = core.suggest_column_mapping(list(df.columns))

    col_options = ["— none —"] + list(df.columns)
    mapping = {}

    req_cols = st.columns(len(core.REQUIRED_FIELDS))
    for c, (field_key, label) in zip(req_cols, core.REQUIRED_FIELDS.items()):
        current = st.session_state.mapping.get(field_key)
        idx = col_options.index(current) if current in col_options else 0
        with c:
            choice = st.selectbox(f"**{label}** (required)", col_options, index=idx, key=f"map_{field_key}")
        mapping[field_key] = None if choice == "— none —" else choice

    with st.expander("Optional columns (used for extra checks & grouping)", expanded=False):
        opt_cols = st.columns(len(core.OPTIONAL_FIELDS))
        for c, (field_key, label) in zip(opt_cols, core.OPTIONAL_FIELDS.items()):
            current = st.session_state.mapping.get(field_key)
            idx = col_options.index(current) if current in col_options else 0
            with c:
                choice = st.selectbox(label, col_options, index=idx, key=f"map_{field_key}")
            mapping[field_key] = None if choice == "— none —" else choice

    st.session_state.mapping = mapping

    missing_required = [core.REQUIRED_FIELDS[k] for k, v in mapping.items() if k in core.REQUIRED_FIELDS and not v]
    if missing_required:
        st.warning(f"Please map the required column(s): {', '.join(missing_required)}")
        st.stop()

    # ======================================================================
    # STEP 3 — Courier filter (only relevant if multiple couriers in the sheet)
    # ======================================================================
    courier_col = mapping.get("courier")
    selected_courier = None
    work_df = df.copy()

    if courier_col:
        couriers = sorted([c for c in work_df[courier_col].dropna().unique()])
        if len(couriers) > 1:
            st.header("4. Courier")
            selected_courier = st.selectbox(
                "This agreement applies to which courier? "
                "(other couriers' AWBs will still be listed, just marked 'Not Checked')",
                couriers,
            )
        elif len(couriers) == 1:
            selected_courier = couriers[0]

    # ======================================================================
    # STEP 4 — Build the working table with editable columns
    # ======================================================================
    step_num = 5 if courier_col else 4
    st.header(f"{step_num}. Review AWBs — add ideal weight & delivery type")

    display_cols = {
        mapping["awb"]: "AWB",
        mapping["chargeable_weight"]: "Billed Chargeable Weight",
        mapping["rate_per_kg"]: "Billed Rate/KG",
        mapping["pickup_state"]: "Pickup State",
        mapping["drop_state"]: "Drop State",
    }
    if mapping.get("customer"):
        display_cols[mapping["customer"]] = "Customer"
    if mapping.get("drop_location"):
        display_cols[mapping["drop_location"]] = "Delivery Location"
    if courier_col:
        display_cols[courier_col] = "Courier"

    base = work_df[list(display_cols.keys())].rename(columns=display_cols)

    # Carry the original (unrenamed) optional billing columns through so run_checks
    # can find them by their *original* header names via `mapping`.
    for extra_key in ("appointment_billed", "total_billed"):
        col_name = mapping.get(extra_key)
        if col_name and col_name not in base.columns:
            base[col_name] = work_df[col_name]

    # Preserve previous edits (Ideal Weight / Delivery Type) across reruns when
    # the row count/order hasn't changed; otherwise start fresh.
    prior = st.session_state.editable_df
    if prior is not None and len(prior) == len(base):
        base["Ideal Weight (Kg)"] = prior.get("Ideal Weight (Kg)", pd.NA)
        base["Delivery Type"] = prior.get(
            "Delivery Type", pd.Series([core.NON_APPOINTMENT] * len(base))
        )
    else:
        base["Ideal Weight (Kg)"] = np.nan
        # Delivery Type is deliberately NOT pre-filled from the courier's billed
        # appointment amount: this is meant to be the ops user's own independent
        # record of what was actually requested, so it can be compared against
        # what was billed. Defaulting it from the billed data itself would let
        # every row auto-agree with the courier and mask real disputes.
        base["Delivery Type"] = core.NON_APPOINTMENT

    edited = st.data_editor(
        base,
        use_container_width=True,
        num_rows="fixed",
        height=420,
        column_config={
            "Ideal Weight (Kg)": st.column_config.NumberColumn(
                "Ideal Weight (Kg)", help="Actual/expected weight for this shipment", min_value=0.0, step=0.5
            ),
            "Delivery Type": st.column_config.SelectboxColumn(
                "Delivery Type", options=[core.APPOINTMENT, core.NON_APPOINTMENT], required=True
            ),
        },
        disabled=[c for c in base.columns if c not in ("Ideal Weight (Kg)", "Delivery Type")],
        key="awb_editor",
    )
    st.session_state.editable_df = edited

    fill_col1, fill_col2 = st.columns(2)
    with fill_col1:
        bulk_weight = st.number_input("Bulk-fill Ideal Weight for blank rows", min_value=0.0, step=0.5, value=0.0)
        if st.button("Apply to blank rows", key="bulk_weight_btn") and bulk_weight > 0:
            mask = edited["Ideal Weight (Kg)"].isna()
            edited.loc[mask, "Ideal Weight (Kg)"] = bulk_weight
            st.session_state.editable_df = edited
            st.rerun()
    with fill_col2:
        bulk_delivery = st.selectbox("Bulk-set Delivery Type for all rows", [core.APPOINTMENT, core.NON_APPOINTMENT], key="bulk_delivery_sel")
        if st.button("Apply to all rows", key="bulk_delivery_btn"):
            edited["Delivery Type"] = bulk_delivery
            st.session_state.editable_df = edited
            st.rerun()

    # ======================================================================
    # STEP 5 — Run the check
    # ======================================================================
    st.header(f"{step_num + 1}. Run Check")

    # Re-attach the original mapped columns (by their real names) onto the
    # edited table so run_checks can look them up via `mapping`.
    run_input = edited.copy()
    rename_back = {v: k for k, v in display_cols.items()}
    run_input = run_input.rename(columns=rename_back)

    if courier_col and selected_courier:
        checkable_mask = run_input[courier_col] == selected_courier
    else:
        checkable_mask = pd.Series([True] * len(run_input))

    if st.button("✅ Run Check", type="primary", use_container_width=True):
        checkable = run_input[checkable_mask].reset_index(drop=True)
        unchecked = run_input[~checkable_mask].reset_index(drop=True)

        results = core.run_checks(checkable, mapping, st.session_state.agreement)

        if len(unchecked):
            unchecked = unchecked.copy()
            unchecked["Computed Pickup Zone"] = None
            unchecked["Computed Drop Zone"] = None
            unchecked["Agreement Rate/KG"] = None
            unchecked["Expected Chargeable Weight"] = None
            unchecked["Rate Match"] = None
            unchecked["Weight Match"] = None
            unchecked["Appointment Match"] = None
            unchecked["Status"] = core.STATUS_NOT_CHECKED
            unchecked["Dispute Reasons"] = "Different courier — no matching agreement selected"
            results = pd.concat([results, unchecked], ignore_index=True)

        # Friendlier column names for the final sheet.
        results = results.rename(columns={
            mapping["awb"]: "AWB",
            mapping["chargeable_weight"]: "Billed Chargeable Weight",
            mapping["rate_per_kg"]: "Billed Rate/KG",
            mapping["pickup_state"]: "Pickup State",
            mapping["drop_state"]: "Drop State",
            **({mapping["customer"]: "Customer"} if mapping.get("customer") else {}),
            **({mapping["drop_location"]: "Delivery Location"} if mapping.get("drop_location") else {}),
            **({courier_col: "Courier"} if courier_col else {}),
            **({mapping["total_billed"]: "Total Charges (Billed)"} if mapping.get("total_billed") else {}),
        })

        st.session_state.results_df = results

    results = st.session_state.results_df

    # ======================================================================
    # STEP 6 — Results, grouping, download
    # ======================================================================
    if results is not None:
        st.header(f"{step_num + 2}. Consolidated Results")

        total = len(results)
        approved = int((results["Status"] == core.STATUS_APPROVED).sum())
        disputed = int((results["Status"] == core.STATUS_DISPUTED).sum())
        not_checked = total - approved - disputed

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total AWBs", total)
        m2.metric("Approved", approved)
        m3.metric("Disputed", disputed)
        m4.metric("Not Checked", not_checked)

        if "Total Charges (Billed)" in results.columns:
            disputed_value = results.loc[results["Status"] == core.STATUS_DISPUTED, "Total Charges (Billed)"].sum()
            st.caption(f"Billed value under dispute: **₹{disputed_value:,.2f}**")

        group_options = ["None"]
        if "Customer" in results.columns:
            group_options.append("Customer")
        if "Delivery Location" in results.columns:
            group_options.append("Delivery Location")
        group_options.append("Status")

        group_by = st.selectbox("Group view by", group_options)

        status_filter = st.multiselect(
            "Filter by status", [core.STATUS_APPROVED, core.STATUS_DISPUTED, core.STATUS_NOT_CHECKED],
            default=[core.STATUS_APPROVED, core.STATUS_DISPUTED, core.STATUS_NOT_CHECKED],
        )
        view = results[results["Status"].isin(status_filter)]

        if group_by == "None":
            st.dataframe(view, use_container_width=True, height=450)
        else:
            summary = (
                view.groupby(group_by)["Status"]
                .value_counts()
                .unstack(fill_value=0)
                .reset_index()
            )
            st.subheader(f"Summary by {group_by}")
            st.dataframe(summary, use_container_width=True)

            for g, gdf in view.groupby(group_by):
                with st.expander(f"{g} — {len(gdf)} AWB(s)"):
                    st.dataframe(gdf, use_container_width=True)

        st.subheader("Download")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Download CSV",
                data=core.to_download_bytes(results, "csv"),
                file_name="shipping_bill_check_results.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ Download Excel (Approved / Disputed tabs)",
                data=core.to_download_bytes(results, "xlsx"),
                file_name="shipping_bill_check_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
