# PCGH Annual QMR Workbook v2.5 — Balance Sheet Quarterly Logic & PA_TB_Recon Redesign

**Workbook:** `Draft_PCGH_Annual_QMR_Workbook_2026_Redesigned_v2_5.xlsx`
**Basis:** v2.4 (`Draft_PCGH_Annual_QMR_Workbook_2026_Redesigned_v2_4.xlsx`)
**Date:** 2026-08-13

This note documents the review findings and the implemented solution for the two
requirements raised on the v2.4 draft:

1. Balance Sheet quarterly opening / closing balance logic.
2. PA_TB_Recon reconciliation source (Projects Forecast → Project Activity).

---

## 1. Review findings — data architecture

The underlying architecture already supports both requirements without restructuring:

- **`Trial_Balance` already carries the mandatory Period identifier** (column B,
  `2026-Q1` … `2026-Q4` format), alongside Entity, Account, Opening_Balance,
  Debit, Credit, Closing_Balance and the QMR_Line mapping. The consolidated
  single-table approach (option 2 in the requirement) is therefore already in
  place — each quarter's TB snapshot is appended as new rows tagged with its
  period key. No separate Q1–Q4 datasets are needed.
- **`Project_Activity` carries the same Period key** plus Project_ID, Nom_Code and
  Amount_Local (ledger sign convention: income negative, costs positive), so it can
  be aggregated by nominal code × project × period directly.
- The Controls tab derives the selected period key (`Controls!B30`, e.g. `2026-Q2`)
  from the single reporting-period-end input (`Controls!B4`); every report reads it.
- The Balance Sheet's Q1–Q4 columns (B:E) were already populated strictly from each
  quarter's **own** TB rows, so loaded quarters are naturally fixed and are never
  overwritten when a later quarter is selected. What was missing was an explicit
  **Opening vs Closing view for the selected quarter**.
- PA_TB_Recon compared **Project_Forecast** flow-group totals (planning data) and
  the nominal ledger against the TB. It did not use `Project_Activity` at all.

One pre-existing defect was found and fixed (see §2.4).

## 2. Balance Sheet — implemented solution

### 2.1 Selected Quarter Opening/Closing block (new columns H:K)

`Balance_Sheet` keeps the historical Q1–Q4 grid in B:E unchanged (all downstream
INDEX references from QMR_Report, Slide_Pack, QMR_Dashboard, Entity_QMR,
Landing_Page and Checks are preserved). A new **Selected Quarter** block was added
in columns H:K for every statement line:

| Column | Content | Rule |
|---|---|---|
| H | **Opening_Balance** | Prior-quarter closing balance, resolved automatically from the quarter selected in Controls |
| I | **Closing_Balance** | The selected quarter's own Trial Balance closing position |
| J | Movement | Closing − Opening |
| K | Movement_% | Movement / abs(Opening) |

Opening resolution order (per line, fully automatic):

1. **Selected quarter is Q2–Q4 and the prior quarter's TB is loaded** → opening =
   prior-quarter closing (Q4 opening = Q3 closing, Q3 opening = Q2 closing, …).
2. **Prior-quarter TB not loaded** (e.g. the 2026 pilot starts at the Q2 snapshot)
   → opening = the selected snapshot's own `Opening_Balance` column, which by
   definition holds the prior-quarter-end position.
3. **Q1** → opening = the Q1 TB opening balances (start of year).
4. Nothing loaded → the block shows `-`.

Cell `H2` displays which basis is in use (e.g. *"Opening basis: Q2 closing trial
balance"*), so the source of the opening figures is always visible — no formula
ever needs editing between quarters.

### 2.2 Historical preservation

Unchanged by design: each Q1–Q4 column only reads TB rows tagged with its own
period key, so completed quarters stay fixed when the reporting period moves on.
The Q3 switch simulation (§4) confirms the Q2 column is bit-identical after
selecting Q3.

### 2.3 Downstream wiring

- `QMR_Report` Section 2's *Movement vs Prior* column (C66:C73) now reads the new
  J-column movements. Previously it read the B:E grid difference, which showed `-`
  whenever the prior quarter's TB was absent; it now reports the true quarterly
  movement in that case as well.
- Checks additions:
  - **Row 42 — Balance sheet opening balances balance**: opening assets − liabilities
    must equal opening capital and reserves (tolerance 1 GHS).
  - **Row 43 — Opening ↔ prior closing continuity**: the loaded quarters' TB opening
    must equal the prior quarter's TB closing (reads Balance_Sheet row 39).
  - The formula-error sweep (row 20) extended over the new columns.

### 2.4 Defect fixed — continuity check (Balance_Sheet row 39)

The inherited row-39 formula summed **every** TB opening balance for the quarter —
which nets to ~0 on any balanced TB — and compared it against prior-quarter Total
Assets, so it could never pass once two consecutive quarters were loaded (it was
masked in v2.4 because Q1 was never loaded). It now aggregates the TB opening
balances of the seven **asset** lines only and compares them with the prior
quarter's closing Total Assets. Verified: with a synthetic Q3 whose openings equal
Q2 closings, the check returns 0 / OK.

## 3. PA_TB_Recon — implemented solution

The tab was rebuilt to reconcile **Project Activity ↔ Trial Balance** (actual
project accounting activity vs accounting balances). Projects Forecast is no
longer part of this reconciliation.

### 3.1 Nominal-code reconciliation (rows 5–27)

Exactly the requested columns, in order: **Nom_Code · Nominal_Description ·
Project_Activity_YTD_GHS · Trial_Balance_YTD_GHS · Variance_PA_less_TB ·
Recon_Status** (plus Flow_Group and a selected-quarter-only PA column for
reference).

- **Project Activity total** — SUMIFS over `Project_Activity` for the nominal code,
  entity, and all periods `FY-Q1` … `FY-Qn` (year-to-date), because the TB's P&L
  closing balances are YTD movements. Like-for-like is guaranteed by the
  *comparison quarter* `n` = the earlier of the selected quarter and the latest
  loaded TB quarter (shown in row 3–4, with a lag warning when the TB trails the
  selection).
- **Trial Balance total** — SUMIFS of `Closing_Balance` for the same code, entity
  and the comparison quarter's period key.
- **Variance = Project Activity − Trial Balance**, with an adjustable rounding
  tolerance input (cell F4, default 1.00 GHS).
- **Status taxonomy:**
  - `RECONCILED` — |variance| within tolerance (suffix "— to TB Qn" when the TB
    lags the selected quarter).
  - `VARIANCE — investigate` — a genuine gap on a project-flow code.
  - `INFO — entity-level income / overhead code / balance-sheet code` — codes that
    are structurally not (fully) project-booked; their variance is displayed for
    completeness but is not a reconciliation failure.
  - `TB PENDING` — no TB loaded yet.
- **Completeness row (27)** — any Project_Activity value on a nominal code missing
  from the driving map is surfaced; the row-driver remains the MM_Income nominal
  map (Z5:AC24), so newly mapped codes appear automatically.
- `Checks!B39` gates on the count of `VARIANCE`/`CHECK` rows.

On the current Q2 data the reconciliation confirms codes 5000, 5250 (and all
zero-activity codes) reconcile exactly, and flags four genuine variances for
investigation: 4000 (+55,891.83), 4050 (+69,359.61), 4200 (−185.75) and
5200 (−51,547.45) — ledger entries not booked to projects, or project bookings
absent from the ledger export (the known 5200 gap from the v2.4 notes persists
here).

### 3.2 Project-level drill-down (rows 29–94)

Validation at **Project / Project Activity + Nominal Code + Reporting Period**
level: enter any nominal code in the yellow input (C30) and every project (driven
from the MM_Income project master, 60 slots) shows its selected-quarter amount,
YTD amount and share of the code total. A residual row catches activity booked to
projects not yet in the master, and a status cell confirms the drill-down ties to
the code total from the reconciliation table above.

## 4. Verification performed

- Full LibreOffice recalculation: **44,594 formulas, zero errors**.
- Reconciliation figures independently recomputed from the raw data (pandas) and
  matched cell-for-cell.
- Balance checks: opening and closing statements both balance (assets −
  liabilities = capital) on the live Q2 data.
- **Quarter-switch simulation**: appended a synthetic, balanced 2026-Q3 TB
  (openings = Q2 closings; +100k on a bank account, −100k on income code 4000) and
  moved the reporting period to Q3. Confirmed automatically and without any
  formula edits:
  - Opening basis label switched to *"Q2 closing trial balance"*; opening figures
    equal the Q2 closing column exactly.
  - Closing figures read the Q3 TB (cash +100k, Current P&L −100k).
  - Historical Q2 column unchanged; Q1 still shows `-`.
  - Continuity and balance checks return OK; the recon comparison quarter moved to
    Q3 and picked up the synthetic TB movement.
  (The simulation is a test artefact only — it is not part of the delivered
  workbook.)

## 5. Quarterly operating workflow (unchanged, now fully automatic)

Each quarter the user only needs to:

1. Append the quarter's source packs (`Project_Activity`, `Nominal_Activity`,
   `Trial_Balance` rows tagged with the new period key).
2. Set `Controls!B4` to the new reporting period end.
3. Maintain the usual yellow management inputs (MM_Income, Project_Forecast,
   Controls quarterly block).

The Balance Sheet opening/closing positions, the PA↔TB reconciliation, the QMR
report movement column and all related checks then update automatically; earlier
quarters remain exactly as published.
