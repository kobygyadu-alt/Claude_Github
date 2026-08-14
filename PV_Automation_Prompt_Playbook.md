# PV Automation — Prompt Playbook

A rewritten, expanded version of the original request:

> *"Study this excel workbook and see if we can turn this into a windows application or VBA coded workbook that can create folders on sharepoint for each PV and its relevant supporting documents… have a master sheet that keeps PVs for the period to use for a dashboard (html) for management reporting."*

The original prompt bundles four decisions into one sentence, pre-commits to a solution ("Windows app or VBA") before the architecture question is answered, gives the AI no schema, no environment facts, and no definition of "better." The versions below fix that. Use **Option A** (one master prompt) with a strong agent that can hold the whole task; use **Option B** (the loop) when working iteratively across sessions.

---

## Option A — The Master Prompt (single-shot)

Copy everything in the block below, attach the workbook, and send.

```text
ROLE
You are a Microsoft 365 solutions architect and finance-systems developer with deep,
current knowledge of: Excel (formulas, tables, protection), VBA, Office Scripts,
Power Automate, SharePoint Online (document libraries, metadata, permissions),
Microsoft Graph / PnP PowerShell, and lightweight self-contained HTML/JS dashboards.
You are helping the Proforest Africa (PFAFR) finance team in Accra, Ghana modernise
their Payment Voucher (PV) process. Be opinionated: recommend, don't list options
without a verdict.

CONTEXT — WHAT EXISTS TODAY
Attached: PFAFR_Payment_Voucher_Template (v1.5 per its Version Log). Structure:

1. "📝 PV Register" — the core sheet. One row per PV, columns A:AF (32 columns):
   PV Number (validated format PV-YYYY-NNNN), Date, Month/Year (auto), Payee,
   Description, Nominal Code (dropdown from Chart of Accounts + VLOOKUP auto-fills
   for Account Name / Type / Classification), Currency + FX rate to GHS, Gross
   Amount, full GRA tax computation (VAT, NHIL, GETFund, COVID levy, WHT by type),
   Total Invoice, Net Payment, Payment Method, Cheque/ETO No., Approval Status
   (Approved / Pending / Rejected), Prepared By, Approved By, Remarks, and
   column AF = "Supporting Document (URL)" — currently EMPTY on every row.
2. "📤 Dashboard Extract" — a flat, snake_case, machine-readable mirror of the
   register (pv_number, date, payee, …, supporting_document), one formula row per
   register row, built specifically to feed HTML dashboards.
3. "🗂️ Chart of Accounts" — 200 accounts: code, name, account type,
   classification, category, 2026 budget (GHS); plus the WHT type/rate table.
4. "💰 Budget vs Actuals" — live SUMIFS of the register against budget per
   account, with committed/pending exposure and RAG status.
5. "⚙️ Settings" — organisation parameters and GRA tax rates (VAT 15%,
   NHIL 2.5%, GETFund 2.5%, COVID-19 levy 0% — scrapped, WHT: services-resident
   7.5%, goods 3%, rent 8%, dividends/interest 8%, exempt 0%).
6. "📋 Instructions" and "📜 Version Log".

Known data-quality issues to verify and fix in your audit:
- The register's banner text and Instructions sheet still say "WHT Services 15%"
  and "COVID-19 Levy 1%", but Settings and the Chart of Accounts use 7.5% and 0%.
  Formulas reference Settings, so the banner/instructions are stale text.
- Instructions say the PV format is PV-YYYY-NNN while validation and sample data
  use PV-YYYY-NNNN.

ENVIRONMENT
- Microsoft 365 tenant (proforest.net) with SharePoint Online and OneDrive sync
  on Windows PCs. The workbook currently lives on SharePoint / synced locally.
- Assume standard Power Automate connectors are available. If any part of your
  recommendation needs premium connectors, Power Apps licences, Azure resources,
  or tenant-admin consent (e.g. Graph app registration), SAY SO EXPLICITLY and
  provide a no-premium fallback.
- Users are finance officers, not developers. IT support is limited.

THE PROBLEM
1. Supporting documents (invoices, quotes, approvals, receipts) live in email and
   personal folders. Nothing links a PV row to its evidence; column AF is unused.
2. The register is one ever-growing sheet with no monthly-period discipline —
   no clean "period close", no locked snapshot for reporting or audit.
3. Management reporting is manual. We want a refreshable HTML dashboard.

GOALS (ranked — trade lower goals away before higher ones)
1. FOLDER AUTOMATION: for every new PV row, automatically create a SharePoint
   folder named from the row (e.g. PV-2026-0001_PayeeName) inside a structured
   library (e.g. /Finance/Payment Vouchers/FY2026/01-January/), with standard
   subfolders (01_Invoice, 02_Approvals, 03_Payment_Evidence), and write the
   folder URL back into column AF of that row. Idempotent: re-running must never
   duplicate folders or overwrite a non-empty AF with a different URL.
2. MASTER REGISTER + PERIODS: keep ONE master register as the single source of
   truth, add a Period concept (month + FY), a month-end close routine that
   locks/snapshots the closed period, and carry the period into the Dashboard
   Extract so reports can filter by period.
3. HTML DASHBOARD: a single self-contained HTML file for management, fed from
   the Dashboard Extract (via exported CSV/JSON or direct read — you choose and
   justify). KPIs: PV count and GHS value for the period, by approval status, by
   classification, top payees, budget vs actuals RAG summary, and a WHT/VAT/levy
   summary usable for GRA filing. Must work offline in a browser with no server.
4. ERROR REDUCTION: tighten validation, protect formula columns, and make the
   PV number self-assigning if feasible.

DELIVERABLES — work in phases. STOP at the end of each phase, present results,
and wait for my confirmation before continuing.

PHASE 1 — AUDIT. Read the workbook. Produce: (a) a schema table of the register,
(b) every formula/validation/consistency issue found (including the known ones
above), (c) questions you need answered. Do not propose solutions yet.

PHASE 2 — ARCHITECTURE DECISION. Compare at least these four approaches for the
folder automation, in a table scored on: build effort, licence cost, fragility,
IT dependency, offline behaviour, and who can maintain it:
  A. VBA in a macro-enabled workbook writing to the locally synced SharePoint
     path via OneDrive sync (no web auth needed).
  B. Office Script + Power Automate flow (button- or schedule-triggered) creating
     folders via the SharePoint connector and writing AF back.
  C. PnP PowerShell / Microsoft Graph script run by finance on a schedule.
  D. A small standalone Windows app (only if A–C genuinely fail a requirement).
Recommend ONE primary and one fallback. Ask me anything that changes the choice
(e.g. is the workbook edited by multiple people simultaneously in the browser?
If yes, VBA is largely dead on arrival — say so).

PHASE 3 — BUILD FOLDER AUTOMATION. Deliver complete, runnable code/flow
definitions for the chosen approach — not fragments. Include: the naming
convention (sanitise payee names for illegal path characters, cap length),
duplicate/rename handling, a dry-run mode, a log sheet of actions taken, and
step-by-step deployment instructions a finance officer can follow, with
screenshots described in words at each click.

PHASE 4 — MASTER REGISTER & PERIOD CLOSE. Deliver the workbook changes: Period
column(s), a Close-Month routine (what gets locked, where the snapshot goes —
e.g. a values-only copy saved to the period's SharePoint folder), reopened-period
rules, and how the Dashboard Extract picks up the period. Preserve all existing
formulas and the existing 37 rows of data.

PHASE 5 — DASHBOARD. Deliver one self-contained HTML file (inline CSS/JS, no
CDNs) that ingests the Dashboard Extract (state the exact export steps), renders
the KPIs above with period and status filters, and degrades gracefully when
fields are blank. Include a 3-step monthly refresh procedure.

PHASE 6 — HARDENING & ROLLOUT. Permissions model (who can see payment data vs
the folder library), sheet protection plan, a pilot plan (one month, parallel
run), a rollback plan, and Version Log entries for every change you made.

CONSTRAINTS
- Do not break existing formulas, validations, or the Budget vs Actuals linkage.
- Do not rename existing sheets or columns without flagging it as a breaking
  change and updating every dependent formula.
- All tax logic must keep referencing the Settings sheet — never hardcode rates.
- Everything must be maintainable by a finance officer with Excel skills.

ACCEPTANCE CRITERIA (I will test against these)
- Entering a new PV row and running the automation yields a correctly named
  folder with subfolders, and AF contains a working link — within one minute.
- Running the automation twice creates nothing new the second time.
- Closing a month locks it and the dashboard filtered to that month matches the
  register's SUMIFS totals to the pesewa.
- The dashboard HTML opens from a double-click with no internet and no errors.

OUTPUT FORMAT
- Code in complete fenced blocks with file/module names.
- Every manual step numbered, written for a non-developer.
- End every phase with: "DECISIONS NEEDED FROM YOU:" followed by a numbered list
  (or 'none').
```

---

## Option B — The Loop (iterative, one phase per session)

Run these in order. Each step has a **repeat-until** condition — do not advance until it's met. Start every new session by pasting the ROLE + CONTEXT + ENVIRONMENT blocks from Option A plus the decisions log you've accumulated.

**Loop 0 — standing header (paste at the top of every session):**

```text
[ROLE + CONTEXT + ENVIRONMENT blocks from the master prompt]

DECISIONS LOG (append to this as we go):
- <date>: <decision made and why>

Current phase: <N>. Do only this phase. Stop when done.
```

**Loop 1 — Audit.** "Audit the attached workbook. Output: register schema table, every formula/validation/text inconsistency (check the WHT 15%-vs-7.5% and COVID 1%-vs-0% banner discrepancies and the PV number format mismatch), and your open questions. No solutions yet."
*Repeat until:* it has found the known issues (proof it actually read the file) and its questions are ones only you can answer. Answer them, add answers to the decisions log.

**Loop 2 — Architecture.** "Using the audit and decisions log, produce the Phase 2 comparison table and recommend one primary + one fallback approach for SharePoint folder automation. Key fact: [the workbook IS / IS NOT edited concurrently in the browser by multiple users]."
*Repeat until:* the recommendation names concrete licence requirements and you agree with it. Record the choice in the decisions log.

**Loop 3 — Build.** "Build Phase 3 for the chosen approach: complete code/flow, naming convention, idempotency, dry-run mode, action log, deployment steps for a non-developer."
*Repeat until it survives a real test.* Run it. Whatever happens, feed it back verbatim:

```text
I ran it. Result:
<paste the exact error message / wrong folder name / screenshot description>
Fix the cause and reissue the COMPLETE corrected code, not a diff.
```

This error-feedback sub-loop is where most of the quality comes from — never paraphrase errors, always paste them.

**Loop 4 — Periods.** "Build Phase 4: period columns, month-end close routine, snapshot location, Dashboard Extract changes. Existing 37 rows and all formulas must survive."
*Repeat until:* you can close a test month and reopen the file with nothing broken.

**Loop 5 — Dashboard.** "Build Phase 5: one self-contained HTML dashboard fed from the Dashboard Extract with period/status filters and the KPI set from the goals. No CDNs."
*Repeat until:* totals match the register exactly for a closed month and it opens offline. Feed rendering problems back with the same error template.

**Loop 6 — Hardening.** "Build Phase 6: permissions, protection, pilot plan, rollback, Version Log entries for everything changed."
*Exit the loop when:* the acceptance criteria from Option A all pass in a one-month parallel run.

---

## Why this rewrite gets better results

1. **Context is injected, not assumed** — the AI gets the real schema, sheet names, and known defects, so it spends its effort building instead of guessing (and the planted defects let you verify it actually read the file).
2. **The solution is no longer pre-decided** — "Windows app or VBA" became an explicit architecture phase; VBA cannot easily authenticate to SharePoint Online, and the right answer usually turns out to be Power Automate or the OneDrive-synced-path trick, which the original phrasing would never have surfaced.
3. **Phases with stop points** prevent the one-shot wall of half-working code.
4. **Ranked goals + acceptance criteria** define "better" so you can test the output instead of eyeballing it.
5. **The error-feedback sub-loop** (paste exact errors, demand complete reissues) converts each failure into a correction instead of drift.
6. **A decisions log carried between sessions** keeps long-running iterative work consistent when context is lost.
