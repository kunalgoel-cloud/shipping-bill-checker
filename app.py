"""
Shipping Bill Checker
======================
Ops tool to audit courier working sheets (billed AWBs) against a commercial
rate agreement. Nothing is written to a database — everything lives in the
browser session and disappears when the tab is closed / refreshed.

Per-AWB workflow: enter ideal weight + delivery type, Calculate the expected
price from the commercial agreement, then Approve it or Reject and enter the
correct price yourself — either way, that price is compared against the
billed subtotal to decide Approved / Disputed.
"""

import numpy as np
import pandas as pd
import streamlit as st

import checker_core as core

st.set_page_config(page_title="Shipping Bill Checker", page_icon="\U0001F69A", layout="wide")

BILLED_FIELD_KEYS = [
    "freight_billed", "fsc_billed", "fov_billed",
    "docket_billed", "state_charges_billed", "appointment_billed",
]


def get_billed_subtotal(row, mapping):
    return core.billed_subtotal(*[row[mapping[k]] for k in BILLED_FIELD_KEYS])


def new_awb_state():
    return {
        "ideal_weight": None,
        "delivery_type": core.NON_APPOINTMENT,
        "calculated": None,      # breakdown dict from calculate_expected_price
        "calc_inputs": None,     # (ideal_weight, delivery_type) used for that calculation
        "decision": None,        # None | "approved" | "rejected"
        "final_price": None,     # the approved-or-corrected price used for the decision
        "price_source": None,    # None | "system" | "user"
        "status": core.STATUS_NOT_CHECKED,
        "reason": "",
    }


def highlight_disputed(row):
    if row.get("Status") == core.STATUS_DISPUTED:
        return [f"background-color: #{core.DISPUTED_FILL}"] * len(row)
    return [""] * len(row)


# --------------------------------------------------------------------------
# Session state defaults
# --------------------------------------------------------------------------
for key, default in {
    "working_df": None,
    "working_file_id": None,
    "mapping": {},
    "agreement": None,
    "awb_workflow": {},  # {str(AWB): {...see new_awb_state()...}}
    "results_df": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# The Safexpress commercial agreement is bundled directly in checker_core.py
# (see default_agreement()) so the app is active and usable with zero setup.
# Uploading a different agreement workbook (Commercial Agreement tab)
# overrides it for the session; refreshing the tab resets back to this
# built-in default, same as everything else — nothing persists.
if st.session_state.agreement is None:
    st.session_state.agreement = core.default_agreement()

st.title("\U0001F69A Shipping Bill Checker")
st.caption(
    "Upload a courier working sheet, calculate the expected price per AWB from "
    "your commercial agreement, approve or correct it, and get an approve/dispute "
    "sheet — nothing is stored server-side."
)
st.caption(
    "Data handling: files are processed in memory for this session only. "
    "Closing or refreshing the tab clears everything — nothing is saved to "
    "a database or disk."
)

# ==========================================================================
# TABS
#   - written in this order so the Checker tab's upload/mapping state is
#     already finalized for this rerun by the time the other two tabs (which
#     read that state) render — tab *visual* order is set by st.tabs() below
#     and is independent of this code order.
# ==========================================================================
tab_checker, tab_mapping, tab_agreement = st.tabs(
    ["\U0001F50D Checker", "\U0001F5C2️ Column Mapping", "\U0001F4C4 Commercial Agreement"]
)

# ==========================================================================
# CHECKER TAB
# ==========================================================================
with tab_checker:
    st.header("1. Upload Working Sheet")
    working_file = st.file_uploader(
        "Courier working sheet (.xlsx, .xls or .csv)", type=["xlsx", "xls", "csv"], key="working_upload"
    )

    if working_file is not None:
        # file_uploader keeps returning the same file across reruns until the
        # user changes it, so only re-parse (and reset mapping/workflow) when
        # it's actually a *different* file — otherwise every rerun would wipe
        # out everything entered in the per-AWB cards below.
        file_id = (working_file.name, getattr(working_file, "size", None))
        if st.session_state.working_file_id != file_id:
            try:
                new_df = core.read_uploaded_table(working_file)
                st.session_state.working_df = new_df
                st.session_state.working_file_id = file_id
                st.session_state.mapping = {}
                st.session_state.awb_workflow = {}
                st.session_state.results_df = None
                st.success(f"Loaded {len(new_df)} rows, {len(new_df.columns)} columns from **{working_file.name}**.")
            except Exception as e:
                st.error(f"Could not read that working sheet: {e}")

    df = st.session_state.working_df

    if df is None:
        st.info("Upload a working sheet above to get started.")
    else:
        with st.expander("Preview raw working sheet", expanded=False):
            st.dataframe(df.head(20), use_container_width=True)

        # Column mapping lives in its own tab; auto-suggest once here so the
        # checker still works out of the box without a visit there.
        if not st.session_state.mapping:
            st.session_state.mapping = core.suggest_column_mapping(list(df.columns))
        mapping = st.session_state.mapping

        missing_required = [core.REQUIRED_FIELDS[k] for k, v in mapping.items() if k in core.REQUIRED_FIELDS and not v]
        if missing_required:
            st.warning(
                f"Please map the required column(s) in the **Column Mapping** "
                f"tab first: {', '.join(missing_required)}"
            )
        else:
            with st.expander("Current column mapping (edit in the Column Mapping tab)", expanded=False):
                mapped_lines = [f"- **{core.ALL_FIELDS[k]}** → `{v}`" for k, v in mapping.items() if v]
                st.markdown("\n".join(mapped_lines))

            # ------------------------------------------------------------
            # Courier filter (only relevant if multiple couriers in the sheet)
            # ------------------------------------------------------------
            courier_col = mapping.get("courier")
            selected_courier = None
            work_df = df.copy()
            awb_col = mapping["awb"]

            if courier_col:
                couriers = sorted([c for c in work_df[courier_col].dropna().unique()])
                if len(couriers) > 1:
                    st.header("2. Courier")
                    selected_courier = st.selectbox(
                        "This agreement applies to which courier? "
                        "(other couriers' AWBs will still be listed, just marked 'Not Checked')",
                        couriers,
                    )
                elif len(couriers) == 1:
                    selected_courier = couriers[0]

            if courier_col and selected_courier:
                checkable_mask = work_df[courier_col] == selected_courier
            else:
                checkable_mask = pd.Series([True] * len(work_df), index=work_df.index)
            checkable_df = work_df[checkable_mask].reset_index(drop=True)
            other_courier_df = work_df[~checkable_mask].reset_index(drop=True)

            # ------------------------------------------------------------
            # Review AWBs — filters + one card per AWB
            # ------------------------------------------------------------
            step_num = 3 if courier_col else 2
            st.header(f"{step_num}. Review AWBs")
            st.caption(
                "For each AWB: enter the ideal weight and delivery type, Calculate "
                "the expected price from the active commercial agreement, then "
                "Approve it or Reject and enter the correct price yourself."
            )

            consignee_col = mapping.get("consignee")
            drop_state_col = mapping["drop_state"]

            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                if consignee_col:
                    consignee_options = sorted(checkable_df[consignee_col].dropna().unique().tolist())
                    selected_consignees = st.multiselect(
                        "Filter by Consignee Name", consignee_options, key="filter_consignee"
                    )
                else:
                    selected_consignees = []
            with filter_col2:
                drop_state_options = sorted(checkable_df[drop_state_col].dropna().unique().tolist())
                selected_drop_states = st.multiselect(
                    "Filter by Drop State", drop_state_options, key="filter_drop_state"
                )

            view_df = checkable_df
            if selected_consignees:
                view_df = view_df[view_df[consignee_col].isin(selected_consignees)]
            if selected_drop_states:
                view_df = view_df[view_df[drop_state_col].isin(selected_drop_states)]
            view_df = view_df.reset_index(drop=True)

            reviewed = sum(
                1 for a in checkable_df[awb_col]
                if st.session_state.awb_workflow.get(str(a), {}).get("final_price") is not None
            )
            filter_note = f" ({len(view_df)} of {len(checkable_df)} listed under the filters above)" if (selected_consignees or selected_drop_states) else ""
            st.caption(
                f"**{reviewed} of {len(checkable_df)}** AWB(s) reviewed"
                + (f" for **{selected_courier}**" if selected_courier else "")
                + filter_note + "."
            )

            for _, row in view_df.iterrows():
                awb = row[awb_col]
                awb_key = str(awb)
                wf = st.session_state.awb_workflow.setdefault(awb_key, new_awb_state())

                badge = {"Approved": "🟢", "Disputed": "🔴", "Not Checked": "⚪"}[wf["status"]]
                header_bits = [f"AWB {awb}", f"{row[mapping['pickup_state']]} → {row[mapping['drop_state']]}"]
                if consignee_col and pd.notna(row.get(consignee_col)):
                    header_bits.append(str(row[consignee_col]))
                if wf["final_price"] is not None:
                    header_bits.append(f"₹{wf['final_price']:,.2f}")

                with st.expander(f"{badge} " + " · ".join(header_bits), expanded=False):
                    ref_bits = []
                    if mapping.get("chargeable_weight") and pd.notna(row.get(mapping["chargeable_weight"])):
                        ref_bits.append(f"Billed Chargeable Weight: {row[mapping['chargeable_weight']]} Kg")
                    if mapping.get("rate_per_kg") and pd.notna(row.get(mapping["rate_per_kg"])):
                        ref_bits.append(f"Billed Rate/KG: ₹{row[mapping['rate_per_kg']]}")
                    if ref_bits:
                        st.caption("Billed reference (not used in the calculation): " + " · ".join(ref_bits))

                    c1, c2 = st.columns(2)
                    with c1:
                        iw = st.number_input(
                            "Ideal Weight (Kg)", min_value=0.0, step=0.5,
                            value=float(wf["ideal_weight"] or 0.0), key=f"iw_{awb_key}",
                        )
                    with c2:
                        dt_idx = 0 if wf["delivery_type"] == core.APPOINTMENT else 1
                        dt = st.selectbox(
                            "Delivery Type", [core.APPOINTMENT, core.NON_APPOINTMENT],
                            index=dt_idx, key=f"dt_{awb_key}",
                        )

                    wf["ideal_weight"] = iw if iw > 0 else None
                    wf["delivery_type"] = dt

                    if wf["calculated"] is not None and wf["calc_inputs"] != (wf["ideal_weight"], wf["delivery_type"]):
                        st.info("Inputs changed since the last calculation — click Calculate again.")
                        wf["calculated"] = None
                        wf["decision"] = None
                        wf["final_price"] = None
                        wf["price_source"] = None
                        wf["status"] = core.STATUS_NOT_CHECKED
                        wf["reason"] = ""

                    if st.button("🧮 Calculate", key=f"calc_{awb_key}"):
                        breakdown = core.calculate_expected_price(
                            pickup_state=row[mapping["pickup_state"]],
                            drop_state=row[mapping["drop_state"]],
                            drop_city=row[mapping["drop_location"]],
                            ideal_weight=wf["ideal_weight"],
                            delivery_type=wf["delivery_type"],
                            invoice_value=row[mapping["invoice_value"]],
                            agreement=st.session_state.agreement,
                        )
                        wf["calculated"] = breakdown
                        wf["calc_inputs"] = (wf["ideal_weight"], wf["delivery_type"])
                        wf["decision"] = None
                        wf["final_price"] = None
                        wf["price_source"] = None
                        wf["status"] = core.STATUS_NOT_CHECKED
                        wf["reason"] = ""
                        st.rerun()

                    bd = wf["calculated"]
                    if bd is not None:
                        if bd["error"]:
                            st.error(bd["error"])
                        else:
                            billed = get_billed_subtotal(row, mapping)
                            st.metric("Calculated Price", f"₹{bd['calculated_total']:,.2f}")

                            with st.expander("🔍 Show the maths", expanded=False):
                                st.markdown(
                                    f"**Zone:** {bd['pickup_zone']} → {bd['drop_zone']}  "
                                    f"(Agreement Rate: ₹{bd['agreement_rate']}/Kg)"
                                )
                                freight_line = f"- Freight = {bd['expected_weight']:.2f} Kg × ₹{bd['agreement_rate']}/Kg = ₹{bd['raw_freight']:,.2f}"
                                if bd["freight_floor_applied"]:
                                    freight_line += f" → floored to **₹{bd['freight']:,.2f}** (Min Chargeable Freight)"
                                else:
                                    freight_line += f" = **₹{bd['freight']:,.2f}**"
                                fov_line = f"- FOV = 0.1% of Invoice Value = ₹{bd['raw_fov']:,.2f}"
                                if bd["fov_floor_applied"]:
                                    fov_line += f" → floored to **₹{bd['fov']:,.2f}** (FOV Minimum)"
                                else:
                                    fov_line += f" = **₹{bd['fov']:,.2f}**"
                                lines = [
                                    f"- Expected Chargeable Weight: **{bd['expected_weight']:.2f} Kg**",
                                    freight_line,
                                    f"- FSC (20% of Freight): **₹{bd['fsc']:,.2f}**",
                                    fov_line,
                                    f"- Docket: **₹{bd['docket']:,.2f}**",
                                    f"- Metro Congestion ({'applies' if bd['metro_applied'] else 'not applicable'}): **₹{bd['metro']:,.2f}**",
                                    f"- Appointment ({'ABD selected' if wf['delivery_type'] == core.APPOINTMENT else 'Non-ABD'}): **₹{bd['appointment']:,.2f}**",
                                    f"- **Calculated Total: ₹{bd['calculated_total']:,.2f}**",
                                ]
                                st.markdown("\n".join(lines))

                                if billed is not None:
                                    matches = abs(bd["calculated_total"] - billed) <= 1.0
                                    if matches:
                                        st.markdown(f"**As per commercial agreement:** ✅ Matches billed (₹{billed:,.2f})")
                                    else:
                                        st.markdown(
                                            f"**As per commercial agreement:** ❌ Does not match billed "
                                            f"(₹{billed:,.2f}, diff ₹{bd['calculated_total'] - billed:,.2f})"
                                        )
                                else:
                                    st.warning("Billed component columns missing for this AWB — can't preview a match.")

                                extra_rows = []
                                for fk in ["dhp_billed", "war_surcharge_billed", "oda_billed", "handling_billed",
                                           "green_tax_billed", "demurrage_billed", "other_charges_billed", "total_billed"]:
                                    col = mapping.get(fk)
                                    if col and pd.notna(row.get(col)):
                                        extra_rows.append((core.OPTIONAL_FIELDS[fk], row[col]))
                                if extra_rows:
                                    st.caption(
                                        "Other billed charges not covered by this agreement "
                                        "(informational only, not part of the match decision):"
                                    )
                                    st.dataframe(
                                        pd.DataFrame(extra_rows, columns=["Charge", "Billed Amount"]),
                                        hide_index=True, use_container_width=True,
                                    )

                            if wf["decision"] is None:
                                ac1, ac2 = st.columns(2)
                                if ac1.button("✅ Approve this price", key=f"appr_{awb_key}", use_container_width=True):
                                    wf["decision"] = "approved"
                                    wf["final_price"] = bd["calculated_total"]
                                    wf["price_source"] = "system"
                                    wf["status"], wf["reason"] = core.decide_status(wf["final_price"], billed)
                                    st.rerun()
                                if ac2.button("✏️ Reject — enter correct price", key=f"rej_{awb_key}", use_container_width=True):
                                    wf["decision"] = "rejected"
                                    st.rerun()

                            if wf["decision"] == "rejected" and wf["final_price"] is None:
                                corrected = st.number_input(
                                    "Correct price (₹)", min_value=0.0, step=1.0, key=f"corr_{awb_key}"
                                )
                                if st.button("Submit corrected price", key=f"subcorr_{awb_key}"):
                                    wf["final_price"] = corrected
                                    wf["price_source"] = "user"
                                    wf["status"], wf["reason"] = core.decide_status(wf["final_price"], billed)
                                    st.rerun()

                            if wf["final_price"] is not None:
                                if wf["status"] == core.STATUS_APPROVED:
                                    st.success(
                                        f"Approved — price used: ₹{wf['final_price']:,.2f} "
                                        f"({'system-calculated' if wf['price_source'] == 'system' else 'your corrected price'})"
                                    )
                                else:
                                    st.error(
                                        f"Disputed — {wf['reason']} "
                                        f"(price used: ₹{wf['final_price']:,.2f}, "
                                        f"{'system-calculated' if wf['price_source'] == 'system' else 'your corrected price'})"
                                    )
                                if st.button("↺ Redo this AWB", key=f"redo_{awb_key}"):
                                    wf["decision"] = None
                                    wf["final_price"] = None
                                    wf["price_source"] = None
                                    wf["status"] = core.STATUS_NOT_CHECKED
                                    wf["reason"] = ""
                                    st.rerun()

            # ------------------------------------------------------------
            # Run the check
            # ------------------------------------------------------------
            st.header(f"{step_num + 1}. Run Check")
            st.caption(
                "Generates the final consolidated sheet from every AWB's current "
                "review state above — covering all AWBs for the selected courier, "
                "not just whichever ones the filters are currently listing."
            )

            if st.button("✅ Run Check / Generate Final Sheet", type="primary", use_container_width=True):
                result_rows = []
                for _, row in checkable_df.iterrows():
                    awb = row[awb_col]
                    wf = st.session_state.awb_workflow.get(str(awb), new_awb_state())
                    billed = get_billed_subtotal(row, mapping)
                    final_price = wf.get("final_price")
                    status, reason = core.decide_status(final_price, billed)
                    bd = wf.get("calculated") or {}
                    result_rows.append({
                        "AWB": awb,
                        "Courier": row[courier_col] if courier_col else "",
                        "Customer": row[mapping["customer"]] if mapping.get("customer") else "",
                        "Consignee Name": row[mapping["consignee"]] if mapping.get("consignee") else "",
                        "Delivery Location": row[mapping["drop_location"]],
                        "Pickup State": row[mapping["pickup_state"]],
                        "Drop State": row[mapping["drop_state"]],
                        "Ideal Weight (Kg)": wf.get("ideal_weight"),
                        "Delivery Type": wf.get("delivery_type"),
                        "Agreement Rate/KG": bd.get("agreement_rate"),
                        "Calculated Price (₹)": bd.get("calculated_total"),
                        "Price Used (₹)": final_price,
                        "Price Source": {"system": "System-calculated", "user": "User-corrected"}.get(wf.get("price_source"), ""),
                        "Billed Subtotal (₹)": billed,
                        "Difference (₹)": (final_price - billed) if (final_price is not None and billed is not None) else None,
                        "Total Charges (Billed, full)": row[mapping["total_billed"]] if mapping.get("total_billed") else None,
                        "Status": status,
                        "Dispute Reasons": reason if status == core.STATUS_DISPUTED else "",
                    })
                for _, row in other_courier_df.iterrows():
                    result_rows.append({
                        "AWB": row[awb_col],
                        "Courier": row[courier_col] if courier_col else "",
                        "Customer": row[mapping["customer"]] if mapping.get("customer") else "",
                        "Consignee Name": row[mapping["consignee"]] if mapping.get("consignee") else "",
                        "Delivery Location": row[mapping["drop_location"]],
                        "Pickup State": row[mapping["pickup_state"]],
                        "Drop State": row[mapping["drop_state"]],
                        "Ideal Weight (Kg)": None,
                        "Delivery Type": None,
                        "Agreement Rate/KG": None,
                        "Calculated Price (₹)": None,
                        "Price Used (₹)": None,
                        "Price Source": "",
                        "Billed Subtotal (₹)": None,
                        "Difference (₹)": None,
                        "Total Charges (Billed, full)": row[mapping["total_billed"]] if mapping.get("total_billed") else None,
                        "Status": core.STATUS_NOT_CHECKED,
                        "Dispute Reasons": "Different courier — no matching agreement selected",
                    })
                st.session_state.results_df = pd.DataFrame(result_rows)

            results = st.session_state.results_df

            # ------------------------------------------------------------
            # Results, grouping, download
            # ------------------------------------------------------------
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

                if "Difference (₹)" in results.columns:
                    disputed_value = results.loc[results["Status"] == core.STATUS_DISPUTED, "Difference (₹)"].abs().sum()
                    st.caption(f"Total ₹ difference across disputed AWBs: **₹{disputed_value:,.2f}**")

                group_options = ["None"]
                for col in ["Customer", "Delivery Location", "Drop State", "Consignee Name"]:
                    if col in results.columns:
                        group_options.append(col)
                group_options.append("Status")

                group_by = st.selectbox("Group view by", group_options)

                status_filter = st.multiselect(
                    "Filter by status", [core.STATUS_APPROVED, core.STATUS_DISPUTED, core.STATUS_NOT_CHECKED],
                    default=[core.STATUS_APPROVED, core.STATUS_DISPUTED, core.STATUS_NOT_CHECKED],
                )
                view = results[results["Status"].isin(status_filter)]
                st.caption("Disputed rows are highlighted below (and in the downloaded Excel); Approved rows are left as-is.")

                if group_by == "None":
                    st.dataframe(view.style.apply(highlight_disputed, axis=1), use_container_width=True, height=450)
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
                            st.dataframe(gdf.style.apply(highlight_disputed, axis=1), use_container_width=True)

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
                        "⬇️ Download Excel (Disputed rows highlighted, separate Approved/Disputed tabs)",
                        data=core.to_download_bytes(results, "xlsx"),
                        file_name="shipping_bill_check_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

# ==========================================================================
# COLUMN MAPPING TAB
# ==========================================================================
with tab_mapping:
    st.header("Column Mapping")

    df_for_mapping = st.session_state.working_df
    if df_for_mapping is None:
        st.info("Upload a working sheet in the **Checker** tab first — mapping options come from its column headers.")
    else:
        st.caption(
            "We've guessed which column is which — check them, especially the "
            "required ones. Changes here apply immediately in the Checker tab."
        )

        if not st.session_state.mapping:
            st.session_state.mapping = core.suggest_column_mapping(list(df_for_mapping.columns))

        col_options = ["— none —"] + list(df_for_mapping.columns)
        new_mapping = {}

        def render_field_selectboxes(fields_dict, chunk_size=4):
            items = list(fields_dict.items())
            for i in range(0, len(items), chunk_size):
                row_items = items[i:i + chunk_size]
                cols = st.columns(len(row_items))
                for c, (field_key, label) in zip(cols, row_items):
                    current = st.session_state.mapping.get(field_key)
                    idx = col_options.index(current) if current in col_options else 0
                    with c:
                        choice = st.selectbox(label, col_options, index=idx, key=f"map_{field_key}")
                    new_mapping[field_key] = None if choice == "— none —" else choice

        st.subheader("Required")
        st.caption("Needed to calculate and compare the expected price — the checker can't run without these.")
        render_field_selectboxes(core.REQUIRED_FIELDS)

        st.subheader("Optional")
        st.caption("Used for grouping, courier filtering, or shown for reference / informational context only.")
        render_field_selectboxes(core.OPTIONAL_FIELDS)

        st.session_state.mapping = new_mapping

        missing_required = [core.REQUIRED_FIELDS[k] for k, v in new_mapping.items() if k in core.REQUIRED_FIELDS and not v]
        if missing_required:
            st.warning(f"Missing required column(s): {', '.join(missing_required)}")
        else:
            st.success("All required columns mapped — head to the Checker tab to continue.")

# ==========================================================================
# COMMERCIAL AGREEMENT TAB
# ==========================================================================
with tab_agreement:
    ag = st.session_state.agreement

    st.header(f"Active agreement: {ag.courier_name}")
    st.caption(
        f"Source: {ag.source}. This is exactly what every AWB's price "
        "calculation in the Checker tab uses right now."
    )

    with st.expander("Change the active agreement", expanded=False):
        st.write(
            "Upload a different agreement workbook to override the built-in "
            "one for this session. Don't have one yet? Grab the template "
            "below — it comes pre-filled with the built-in Safexpress rate "
            "card, ready to overwrite for another courier."
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

        if ag.source.startswith("Uploaded"):
            if st.button("↺ Reset to built-in Safexpress agreement", use_container_width=True):
                st.session_state.agreement = core.default_agreement()
                st.rerun()

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
        "matrix is not symmetric. Every AWB's calculated price is built "
        "from the rate looked up here, using the zone derived from its "
        "Pickup/Drop State, never from any zone column the working sheet "
        "itself claims."
    )
