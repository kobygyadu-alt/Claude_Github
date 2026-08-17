# Phase 1 — Workbook Audit Report

**Workbook:** `PFAFR_Payment_Voucher_Template_v1.0_Final_1_1.xlsx` (internally v1.5 per 📜 Version Log)
**Audited:** 2026-08-17 · **PV rows:** 37 (Dec 2025 – Mar 2026) · **Chart of Accounts:** 200 accounts (45 with FY2026 budgets)

## 1. Register schema (📝 PV Register, columns A:AF)

| Cols | Group | Fields | Behaviour |
|---|---|---|---|
| A–D | PV identification | PV Number, Date, Month, Year | A validated `PV-` prefix + length ≥ 10; C/D auto from B |
| E–F | Payee & description | Payee/Vendor, Payment Description | free text |
| G–J | Account coding | Nominal Code ▾, Account Name, Account Type, Classification | G dropdown from CoA; H/I/J VLOOKUP auto-fill |
| K–M | Currency & amount | Currency ▾, FX Rate, Gross Amount (FCY) | currencies list from ⚙️ Settings |
| N–Y | GRA tax computation | GHS Gross, VAT Apply?, WHT Type ▾/Rate/Amount, NHIL, GETFund, COVID, VAT, Total Taxes, Total Invoice, Net Payment | all formulas reference ⚙️ Settings rates and the CoA WHT table |
| Z–AA | Payment | Payment Method ▾, Cheque/ETO No. | method list in validation |
| AB–AF | Approval & workflow | Status ▾ (Approved/Pending/Rejected), Prepared By, Approved By, Remarks, **Supporting Document (URL)** | AF is the automation write-back target |

**📤 Dashboard Extract** mirrors the register into flat snake_case columns (`pv_number` … `supporting_document`), one formula row per register row — the intended machine-readable feed.

## 2. Findings

### F1 — Stale banner and instructions text (cosmetic, but misleading)
The register banner (A2) and 📋 Instructions still say **"WHT Services 15%"** and **"COVID-19 Levy 1%"**, while ⚙️ Settings (which the formulas actually use) has **7.5%** and **0% (levy scrapped)**. The numbers computed are correct; the text is wrong. *Fix: edit the two text cells; no formula change needed.*

### F2 — Instructions reference an obsolete column layout
The "GRA Tax Computation Model" section in 📋 Instructions cites columns K/O/P/Q/R/S/U/V for Gross/WHT/NHIL/GETFund/COVID/VAT/Total/Net, but the live register uses **M/R/S/T/U/V/X/Y**. Anyone following the guide letter-by-letter lands on the wrong columns. *Fix: update the letters (or drop letters and use column names).*

### F3 — PV number format inconsistency
Instructions say `PV-YYYY-NNN` (e.g. PV-2025-021); validation only enforces `PV-` prefix and length ≥ 10; actual data uses `PV-YYYY-NNNN`. *Fix: standardise on `PV-YYYY-NNNN` and tighten the custom validation. The automation assumes `PV-YYYY-NNNN`.*

### F4 — Column AF ("Supporting Document (URL)") is empty on all 37 rows
No PV currently links to its evidence — the core problem the folder automation solves.

### F5 — PV numbering year vs payment date year
`PV-2026-0001` is dated **2025-12-01** (register Month/Year show Dec 2025). Decide whether the number's year segment should match the payment date's fiscal year (Settings: FY starts January). The automation defaults to **date-based** foldering (`FY2025/12-December` for that row) and this is configurable; the dashboard also periods by date.

### F6 — Dashboard Extract silently drops three register columns
The extract has no `month`/`year` (derivable from `date` — fine), no **COVID levy** column (harmless while the rate is 0, but the extract's `total_taxes` would stop reconciling to its visible tax columns if the levy ever returns), and no **Cheque/ETO No.** *Fix when convenient: add `covid_levy` and `cheque_eto_no` columns.*

### F7 — Duplicate PV numbers are not prevented
Validation checks format only. Two rows can share a PV number, which would also collide in folder naming. *Fix: extend the custom validation to `COUNTIF($A$5:$A$510,A5)=1` AND the format check; the automation additionally logs duplicates and skips them.*

### F8 — Sparse classification coverage, and blank classifications display as `0`
Only 45 of 200 accounts carry a FY2026 budget, and **33 of the 37 PVs show a literal `0` in the Classification column**: the register's formula `=IF(G5="","",IFERROR(VLOOKUP(...),"—"))` catches a *missing code* but not a *blank classification cell* — VLOOKUP returns 0 for blanks. *Fix: wrap the lookup, e.g. `=IF(G5="","",IF(VLOOKUP(G5,'🗂️ Chart of Accounts'!$A:$D,4,0)=0,"—",VLOOKUP(...)))`.* The dashboard buckets `0`/blank/`—` as "Unclassified". Management reporting by classification is only as good as this coding discipline.

### F9 — Workflow fields are single-sourced
All 37 rows: status "Approved", mostly one preparer, `Approved By` largely blank. Approval is recorded, not enforced. Real segregation of duties needs the SharePoint/Power Automate layer (Phase 6), not Excel.

### F10 — File name vs version drift
The file is named `v1.0_Final_1_1` while the Version Log says v1.5. Trust the Version Log; rename the file at the next release.

## 3. What is solid (keep as-is)

- All tax formulas reference ⚙️ Settings / CoA rate tables — no hardcoded rates anywhere sampled.
- Data validation on code, classification, currency, WHT type, payment method, and status.
- The Dashboard Extract pattern (formula-mirrored flat table) is exactly the right feed for an HTML dashboard.
- Budget vs Actuals SUMIFS linkage and the Version Log discipline.

**Cached values were verified present for all 37 rows** (the computed GHS/tax columns export correctly to CSV), so the dashboard can be fed by a plain CSV export of 📤 Dashboard Extract with no recalculation step.
