# Phase 5 — HTML Dashboard Guide

**File:** `dashboard/PV_Dashboard.html` — one self-contained page (no internet, no server, no CDNs). Double-click to open in any modern browser; it also works from a SharePoint-synced folder. Data never leaves the browser.

## What it shows

| Section | Content | Scope |
|---|---|---|
| KPI tiles | PV count, gross spend (with 12-month sparkline and vs-prior-month delta when a single month is selected), net paid, taxes & levies (WHT called out) | filter slice |
| Gross spend by month | Column chart; when a month filter is active that month is highlighted and the rest grey out for context | selected FY |
| Approval status | Chips + share-of-gross stacked bar (Approved / Pending / Rejected) | filter slice |
| Spend by classification | Horizontal bars; blank/`0` classifications bucket as "Unclassified" | filter slice |
| Top payees | Largest 8 by gross | filter slice |
| GRA filing summary | WHT by type, VAT, NHIL, GETFund, total WHT payable | filter slice |
| Budget vs actuals | Annual budgets by classification (embedded from the CoA), actuals = Approved gross, committed = Pending, % used meter + RAG chip | **always full FY** (month filter deliberately ignored — stated on the card) |
| PV register | Filtered rows, newest first, with a 📁 link wherever column AF / `supporting_document` holds a URL — this is where the folder automation becomes visible to management | filter slice |

Filters (one row, scoping everything): financial year, month, status, and free-text search over PV number / payee / description. Every chart has a **⊞ table** toggle, and all values are reachable without hovering (labels, tables) — tooltips only enhance.

## Refreshing the data (monthly, ~2 minutes)

1. Open the PV workbook → select the **📤 Dashboard Extract** sheet.
2. File → Save As / Export → **CSV UTF-8** → e.g. `dashboard_extract.csv` (Excel exports the active sheet only).
3. Open the dashboard → **Load data (CSV)** (or drag the file onto the page).

The page ships with an embedded snapshot (37 PVs, generated 2026-08-17) so it renders even before the first refresh. Loading a CSV replaces the PV data; **the budget column stays embedded** (it comes from the Chart of Accounts, which the extract doesn't carry) — regenerate the dashboard when annual budgets change.

Date handling: ISO (`2026-01-14`), day-first (`14/01/2026`), month-first when unambiguous, and Excel serial numbers are all accepted; ambiguous dates are read **day-first** (Ghana locale). Financial year = calendar year (⚙️ Settings: FY starts January).

## Design notes

- Proforest visual identity: turquoise `#318E93` chrome, Aptos/Segoe UI/Calibri, wordmark + signature curve, subtle watermark, white cards.
- Chart marks use a **validated colour set** (`#00A6AB` primary; `#D38348`/`#2961A3` reserved for any future second/third series): checked with the dataviz palette validator for lightness band, chroma floor, colour-vision-deficiency separation and normal-vision separation. The two sub-3:1 contrast warnings are relieved by direct labels and the table views, per the validator's contract.
- Status colours (`#007361` / `#DABA38` / `#A3352F`) are semantic-only and always paired with an icon + label — never colour alone.
- Print-friendly: the browser's Print (Ctrl+P) hides controls and prints the cards — a quick way to a PDF pack for a management meeting.

## Reconciliation checks (rerun after any refresh)

- KPI gross for "FY, all months, all statuses" = SUM of register column N for that year.
- GRA card's "Total WHT payable" = SUM of column R.
- Budget vs actuals "Actuals" total = SUM of Approved gross; "Committed" = Pending gross.

Verified against the embedded snapshot: FY2026 gross GHS 1,756,132.82 across 36 PVs (+ GHS 2,405 in Dec FY2025), total WHT GHS 2,349.68 — all matching the workbook's cached values to the pesewa.
