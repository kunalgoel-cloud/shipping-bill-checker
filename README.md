# Shipping Bill Checker

An ops tool that audits a courier working sheet (billed AWBs) against your
commercial rate agreement, so disputes can be raised before the bill is paid.

**No database. No data is stored anywhere** — everything happens in memory
for the current browser session and disappears when the tab is closed or
refreshed.

The **Safexpress commercial agreement** is bundled directly in the code
(`checker_core.py` → `default_agreement()`) and is active from the moment
the app loads — no upload needed to get started. A **Commercial Agreement**
tab in the app always shows exactly which agreement (rate matrix, zone
mapping, charges) is currently active, whether that's the built-in default
or one you've uploaded.

## What it does

1. The app opens with the **Safexpress agreement already active** —
   nothing to upload to get going. Check what's active any time in the
   **Commercial Agreement** tab.
2. Upload the courier working sheet (xlsx / xls / csv).
3. The app reads AWB, Chargeable Weight and Rate/KG from it.
4. Need a different courier's agreement? Upload your own agreement workbook
   (zone-to-zone rate matrix) in the sidebar to override the built-in one
   for the session. The app auto-plots the expected per-kg rate for each
   AWB based on pickup and drop state.
5. For each AWB, add the **ideal (actual) weight** and pick
   **Appointment / Non-Appointment** delivery from a dropdown.
6. Click **Run Check** — AWBs where the billed rate, billed weight and
   billed appointment charge all match your inputs are marked **Approved**;
   anything that doesn't match is marked **Disputed**, with the reason
   spelled out; anything that couldn't be verified yet (e.g. no ideal
   weight entered) is marked **Not Checked** — never silently approved.
7. Group the results by Customer or Delivery Location, and download the
   consolidated sheet as CSV or Excel (with separate Approved / Disputed
   tabs).

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI / page flow |
| `checker_core.py` | Parsing, zone lookup, and the approve/dispute logic (no Streamlit dependency — easy to test on its own) |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme + upload size limit |

## Commercial agreement format

The app doesn't parse courier-agreement PDFs directly (every courier
formats these differently and it's not reliable to auto-extract). Instead,
the Safexpress rate card is embedded as plain Python data in
`checker_core.py` (`_MATRIX_VALUES` / `_ZONE_LOCATIONS` / `_DEFAULT_CHARGES`)
and loaded as the active agreement automatically via `default_agreement()`.
For any other courier, download the **agreement template** from the
sidebar — it's a 3-tab Excel workbook:

- **ZoneMatrix** — per-kg rate from the zone in each row to the zone in
  each column.
- **ZoneMapping** — which states/locations fall in each zone.
- **Charges** — flat/percentage charge parameters (docket, FOV, FSC,
  minimum chargeable weight, appointment charge amount, etc.) used by the
  weight-floor and appointment checks.

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

Streamlit Cloud deploys directly from a GitHub repo that **you** own, so
this needs to be pushed from your own GitHub account (Claude does not have
a connector to your GitHub, so it can't create the repo or push code for
you — see the note in the chat for details).

1. Create a new **empty** repository on GitHub (e.g. `shipping-bill-checker`).
2. From this folder, run:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: shipping bill checker"
   git branch -M main
   git remote add origin https://github.com/<your-username>/shipping-bill-checker.git
   git push -u origin main
   ```
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, pick this repo/branch, and set the main file
   to `app.py`.
4. Deploy. No secrets or environment variables are required — the app has
   no external connections at all.

## Notes / things to sanity-check before rolling out to the team

- The zone lookup is driven entirely by **Pickup State / Drop State** text
  matching the `ZoneMapping` tab. If a working sheet uses a state name or
  abbreviation not listed there, that AWB's zone (and therefore rate
  check) will come back blank — it's shown as "Not Checked" rather than
  silently marked approved.
- The weight check compares the billed **Chargeable Weight** against
  `max(your ideal weight, Min Chargeable Weight from the agreement)` — so
  it correctly accounts for the 15 kg floor rather than disputing every
  small shipment.
- The appointment check needs the working sheet's **Appointment Charges
  (billed amount)** column mapped; if it isn't, that check is skipped
  (shown as "Not Checked") rather than guessed.
- If a working sheet has more than one courier, pick which courier the
  uploaded agreement applies to — AWBs from other couriers are listed but
  marked "Not Checked" rather than compared against the wrong rate card.
