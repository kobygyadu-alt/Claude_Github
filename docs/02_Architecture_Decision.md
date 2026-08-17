# Phase 2 — Architecture Decision: SharePoint Folder Automation

**Decision: Option B (Office Scripts + Power Automate) as primary, Option A (VBA on the OneDrive-synced path) as fallback.** Both are built and included in this repo.

## Options compared

Scored 1 (poor) – 5 (good).

| Criterion | A. VBA + synced path | B. Office Script + Power Automate | C. PnP PowerShell / Graph | D. Standalone Windows app |
|---|---|---|---|---|
| Build effort | 4 | 4 | 3 | 1 |
| Licence cost | 5 (none) | 5 (standard connectors only) | 4 (needs PnP module + app consent) | 2 |
| Works with browser co-authoring | 1 (macros dead in Excel Online) | 5 | 3 (external to Excel) | 3 |
| Fragility | 2 (per-user sync paths, macro security) | 4 | 3 (token/module updates) | 2 |
| IT dependency | 5 (none) | 4 (Power Automate access) | 2 (admin consent likely) | 1 |
| Finance-officer maintainable | 3 | 4 | 2 | 1 |
| Audit trail | 3 (log sheet) | 5 (flow run history + log sheet) | 3 | 2 |
| **Total** | **23** | **31** | **20** | **12** |

## Why B wins

- **Survives co-authoring.** The workbook lives on SharePoint; the moment two people edit in the browser, VBA cannot run there. Office Scripts run in Excel Online natively and the flow touches the file server-side.
- **No premium licensing.** SharePoint connector, Excel Online (Business) "Run script", and Office Scripts are all standard M365 connectors.
- **Idempotent by construction.** Folder creation uses SharePoint's `POST /_api/web/folders` (ensure-semantics: returns the existing folder rather than erroring), and the script only selects rows where AF is empty — two independent guards against duplicates.
- **Two audit trails.** Every run is recorded in Power Automate run history *and* in a `🔁 Automation Log` sheet the script maintains inside the workbook.

## Why A is the fallback (and still shipped)

If Power Automate access is blocked or the flow is down, any finance PC with the library synced via OneDrive can run the VBA module: it creates the same folder structure through the synced path (no web authentication needed — OneDrive does the uploading) and writes the same URLs into AF. Costs: the workbook must be saved `.xlsm`, macros must be allowed, it must be run from desktop Excel while nobody is editing in the browser, and the synced local root differs per PC (a config constant).

## Why not C or D

**C** does the job but needs PowerShell skills and usually a tenant-admin app registration/consent — the wrong maintenance profile for a finance team. **D** (a Windows app) means authentication code, MSAL app registration, packaging, updates, and a bus-factor of one; nothing in the requirements needs it.

## Folder design (both implementations)

```
<Site>/Shared Documents/Payment Vouchers/
  FY2026/
    01-January/
      PV-2026-0003_Complete-Solutions-Ltd/
        01_Invoice/
        02_Approvals/
        03_Payment_Evidence/
```

- **FY and month derive from the PV date** (configurable; see audit finding F5).
- **Folder name** = `PV number + "_" + sanitised payee`: illegal characters `" * : < > ? / \ | # %` and leading/trailing dots/spaces removed, spaces→`-`, capped at 80 characters.
- **Write-back**: the folder's URL goes into register column **AF** as a clickable hyperlink; the Dashboard Extract's `supporting_document` column then carries it to the dashboard, which renders a 📁 link per PV.
- **Idempotency contract**: a row with non-empty AF is never touched; re-running creates nothing new; a duplicate PV number is logged and skipped.

## Decisions needed from you

1. Confirm the SharePoint site + library that should hold `Payment Vouchers` (the guides use placeholders).
2. FY/month from payment **date** (default) or from the PV number's year segment?
3. Subfolder set — keep `01_Invoice / 02_Approvals / 03_Payment_Evidence` or adjust?
4. Flow trigger — scheduled (hourly, weekdays) or a manual "Create folders now" button, or both (default in the guide: both).
