# Shipping Bill Checker

An ops tool that calculates the **expected price** for each AWB in a courier
working sheet from your commercial agreement, lets you approve or correct
that price per shipment, and compares it against what was actually billed —
so disputes can be raised before the bill is paid.

**No database. No data is stored anywhere** — everything happens in memory
for the current browser session and disappears when the tab is closed or
refreshed.

The **Safexpress commercial agreement** is bundled directly in the code
(`checker_core.py` → `default_agreement()`) and is active from the moment
the app loads — no upload needed to get started. The **Commercial
Agreement** tab always shows exactly which agreement (rate matrix, zone
mapping, charges) is currently active, whether that's the built-in default
or one you've uploaded — and is also where you upload a different one,
download the blank template, or reset back to Safexpress.

## Tabs

- **🔍 Checker** — the main workflow: upload the working sheet, review AWBs
  one at a time, run the check, see consolidated results.
- **🗂️ Column Mapping** — confirm/override which of the working sheet's
  columns map to AWB, Invoice Value, the billed charge components, etc.
  Auto-suggested, so most sheets need no changes here at all; changes apply
  immediately back in the Checker tab.
- **📄 Commercial Agreement** — view the currently active rate card, and
  upload/download/reset it.

## What it does

1. The app opens with the **Safexpress agreement already active** —
   nothing to upload to get going. Check what's active any time in the
   **Commercial Agreement** tab.
2. Upload the courier working sheet (xlsx / xls / csv) in the **Checker**
   tab, using the column mapping (auto-suggested, editable in the **Column
   Mapping** tab).
3. Need a different courier's agreement? Upload your own agreement workbook
   (zone-to-zone rate matrix) in the **Commercial Agreement** tab to
   override the built-in one for the session.
4. Optionally filter the AWB list by **Consignee Name** and/or **Drop
   State** to focus on a subset — only matching AWBs are listed. Each AWB
   is its own card:
   - Enter the **ideal (actual) weight** and pick **Appointment /
     Non-Appointment** delivery.
   - Click **Calculate** — the app works out the expected price from the
     active agreement (see *Price formula* below) and shows it, with a
     **Show the maths** breakdown of every line item and whether it
     already matches what was billed.
   - **Approve** that price, or **Reject** it and type in the price you
     believe is correct instead.
   - Either way, that price (system-calculated or your correction) is then
     compared against the billed subtotal: matches → **Approved**;
     doesn't → **Disputed**, with the ₹ gap shown.
   - Edits are remembered per-AWB, so switching or clearing the filters, or
     scrolling past a card, never loses what you've already entered.
5. Click **Run Check** — this always covers every AWB for the selected
   courier, not just whatever the filters are currently showing. Any AWB
   you haven't gone through the Calculate → Approve/Reject flow for yet is
   marked **Not Checked** — never silently approved.
6. Group the results by Customer, Delivery Location, Drop State, or
   Consignee Name. **Disputed rows are highlighted** (on-screen and in the
   downloaded Excel) with the mismatch reason; Approved rows are left as-is.
   Download as CSV or Excel (with separate Approved / Disputed tabs).

## Price formula

```
Expected Weight   = max(Ideal Weight, Min Chargeable Weight)
Agreement Rate/KG = ZoneMatrix[pickup zone][drop zone]      (direction matters)
Freight           = max(Expected Weight × Agreement Rate/KG, Min Chargeable Freight)
FSC               = FSC% × Freight
FOV               = max(FOV% × Invoice Value, FOV Minimum)
Docket            = flat Docket Charge
Metro             = Metro Congestion Charge, if the drop city is a metro location, else 0
Appointment       = Appointment Charge Amount, if Appointment based Delivery is selected, else 0

Calculated Total  = Freight + FSC + FOV + Docket + Metro + Appointment
```

This is compared against the **same six components read from the working
sheet's billed columns** (Freight, FSC, FOV, Docket, State Charges — the
billed counterpart of Metro Congestion — and Appointment), within ₹1.
Charges the agreement doesn't govern at all (DHP, War Surcharge, ODA,
Handling, Green Tax, Demurrage, Other Charges) are deliberately **excluded**
from this comparison — on real data they're billed on top of the
agreement-covered charges regardless of whether those are correct, so
including them would make almost every AWB look disputed. They're still
shown in each card's "Show the maths" breakdown for reference.

Two figures in the built-in Safexpress agreement aren't in the source PDF —
they're derived from the real working sheet's own billed data, which was
extremely consistent:
- **Appointment Charge Amount = ₹750 flat**, seen on every appointment-billed row.
- **Metro locations** include "Bangalore" as well as "Bengaluru" (same city,
  common alternate spelling in courier sheets) — metro detection also uses
  a contains-match, so "MUMBAI CITY" or "BANGALORE DELIVERY" style Drop City
  values are still recognized.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI / page flow (3 tabs: Checker, Column Mapping, Commercial Agreement) |
| `checker_core.py` | Parsing, zone lookup, price calculation, and Approved/Disputed decision logic (no Streamlit dependency — easy to test on its own) |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme + upload size limit |

## Commercial agreement format

The app doesn't parse courier-agreement PDFs directly (every courier
formats these differently and it's not reliable to auto-extract). Instead,
the Safexpress rate card is embedded as plain Python data in
`checker_core.py` (`_MATRIX_VALUES` / `_ZONE_LOCATIONS` / `_DEFAULT_CHARGES`)
and loaded as the active agreement automatically via `default_agreement()`.
For any other courier, download the **agreement template** from the
**Commercial Agreement** tab — it's a 3-tab Excel workbook:

- **ZoneMatrix** — per-kg rate from the zone in each row to the zone in
  each column.
- **ZoneMapping** — which states/locations fall in each zone.
- **Charges** — flat/percentage charge parameters (docket, FOV, FSC,
  minimum chargeable weight, minimum chargeable freight, metro congestion
  charge, appointment charge amount) used by the price calculation above.

The template ships pre-filled with the Safexpress rate card provided, so
it works out of the box for that courier — just overwrite the numbers for
any other courier/agreement, keeping the three tab names unchanged.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually http://localhost:8501).

## Deploy to Streamlit Community Cloud

This repo is already on GitHub. To deploy:

1. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, pick this repo/branch, and set the main file
   to `app.py`.
2. Deploy. No secrets or environment variables are required — the app has
   no external connections at all.

## Notes / things to sanity-check before rolling out to the team

- The zone lookup is driven entirely by **Pickup State / Drop State** text
  matching the `ZoneMapping` tab. If a working sheet uses a state name or
  abbreviation not listed there, that AWB's zone (and therefore its
  calculated price) can't be worked out — Calculate will show an error
  rather than a wrong number.
- An AWB can only reach **Approved** by actually going through Calculate →
  Approve/Reject for it. Nothing is ever inferred as a pass just because
  another check happened to succeed.
- The required column mapping is larger than it used to be — the price
  calculation and its comparison both need real ₹ figures (Invoice Value,
  and each billed charge component), not just a rate/kg and a total. If a
  working sheet is missing one of these columns entirely, map what you can
  in the **Column Mapping** tab; AWBs will show a "missing input" error on
  Calculate rather than a silently wrong price.
- If a working sheet has more than one courier, pick which courier the
  active agreement applies to — AWBs from other couriers are listed but
  marked "Not Checked" rather than compared against the wrong rate card.
