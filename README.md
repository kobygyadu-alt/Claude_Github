# PFAFR PV Automation — SharePoint Folders + Management Dashboard

Deliverables for modernising the Proforest Accra (PFAFR) Payment Voucher process, built from
`PFAFR_Payment_Voucher_Template` (v1.5). Three problems solved: supporting documents get an
automatically created, linked SharePoint folder per PV; the register stays the single source of
truth; management gets a self-contained HTML dashboard.

## What's here

| Path | What it is |
|---|---|
| `PV_Automation_Prompt_Playbook.md` | The engineered prompt / iterative loop this build followed |
| `docs/01_Audit_Report.md` | Phase 1 — workbook schema + 10 findings (F1–F10) with fixes |
| `docs/02_Architecture_Decision.md` | Phase 2 — options compared; Office Scripts + Power Automate chosen, VBA fallback |
| `automation/office-scripts/PVFolders_GetPending.ts` | Office Script: find PV rows needing folders |
| `automation/office-scripts/PVFolders_WriteLinks.ts` | Office Script: write folder URLs to column AF + log sheet |
| `automation/power-automate/Flow_Build_Guide.md` | Phase 3 — click-by-click flow build, test plan, failure modes |
| `automation/vba-fallback/PVFolders.bas` | Same automation via the OneDrive-synced path (no flow needed) |
| `dashboard/PV_Dashboard.html` | Phase 5 — offline Proforest-branded dashboard (embedded snapshot + CSV refresh) |
| `docs/03_Dashboard_Guide.md` | Dashboard contents, refresh steps, reconciliation checks |

## Quick start

1. Read `docs/01_Audit_Report.md` and apply the two-minute text fixes (F1–F3) in the workbook.
2. Install the two Office Scripts, then build the flow with `automation/power-automate/Flow_Build_Guide.md`
   (dry-run first — the guide includes the test procedure).
3. Open `dashboard/PV_Dashboard.html`; refresh it monthly from a CSV export of 📤 Dashboard Extract.

## Not yet built (recommended next phases)

- **Phase 4 — month-end close**: period lock + values-only snapshot into the period's SharePoint folder.
- **Phase 6 — hardening**: library permissions, sheet protection review, pilot month, rollback plan.

The `docs/02_Architecture_Decision.md` file ends with the four configuration decisions needed
(site/library, FY-from-date vs FY-from-PV-number, subfolder set, trigger cadence).
