# Power Automate Flow — Build Guide (primary solution)

Creates a SharePoint folder (with subfolders) for every PV row whose
"Supporting Document (URL)" cell is empty, then writes the folder link back
into the workbook. Uses **standard connectors only** (SharePoint, Excel Online
Business). Allow ~45 minutes the first time.

## Before you start — fill in your values

| Placeholder | Meaning | Example |
|---|---|---|
| `SITE_URL` | SharePoint site holding the library | `https://proforest.sharepoint.com/sites/PFAFRFinance` |
| `SITE_PATH` | Server-relative site path | `/sites/PFAFRFinance` |
| `LIBRARY` | Document library (server-relative folder name) | `Shared Documents` |
| `ROOT` | Root folder for vouchers | `Payment Vouchers` |
| Workbook | The PV workbook stored in SharePoint | `…/Shared Documents/PV Register/PFAFR_PV_Register.xlsx` |

**Prerequisite:** install the two Office Scripts first (open the workbook →
**Automate** tab → **New Script** → paste each file from
`automation/office-scripts/` → rename to exactly **"PV Folders – Get Pending"**
and **"PV Folders – Write Links"** → Save script).

---

## Build the flow

Go to [make.powerautomate.com](https://make.powerautomate.com) → **Create** →
**Scheduled cloud flow** → name it `PV – Create SharePoint folders` → repeat
**every 1 hour** → Create. (You can also add a manual button later — step 9.)

### 1. Initialize the links array

**+ New step** → search *Initialize variable* →
Name: `links` · Type: `Array` · Value: *(leave empty)*

### 2. Initialize the dry-run switch

**+ New step** → *Initialize variable* →
Name: `dryRun` · Type: `Boolean` · Value: `false`
*(Set to `true` while testing: folders are still created in SharePoint, but no
links are written to the workbook — the log sheet records what would happen.
For a fully consequence-free first test, point step 4 at a scratch library.)*

### 3. Run script — Get Pending

**+ New step** → connector **Excel Online (Business)** → action **Run script** →
- Location / Document Library / File: browse to the PV workbook
- Script: **PV Folders – Get Pending**

### 4. Parse the script result

**+ New step** → *Parse JSON* →
- Content: `Result` (dynamic content from step 3)
- Schema → **Generate from sample** → paste:

```json
{"pending":[{"row":5,"pvNumber":"PV-2026-0001","payee":"Name","dateIso":"2026-01-14","fy":"FY2026","monthFolder":"01-January","folderName":"PV-2026-0001_Name","subPath":"FY2026/01-January/PV-2026-0001_Name"}],"skipped":[{"row":6,"pvNumber":"PV-2026-0002","reason":"text"}],"scannedRows":506}
```

### 5. For each pending PV — create the folder tree

**+ New step** → *Apply to each* → output: `pending` (from Parse JSON).
Open **Settings** of this Apply-to-each (⋯ menu) → **Concurrency control: On,
degree of parallelism 1** (keeps the workbook write-back and log orderly).

Inside the loop add **five** *"Send an HTTP request to SharePoint"* actions
(connector: **SharePoint** — this action is standard, not premium). All five are
identical except the folder path at the end. This endpoint is **idempotent**:
if the folder already exists it returns it without erroring — safe to re-run.

Common settings for all five:
- Site Address: `SITE_URL`
- Method: `POST`
- Uri: `_api/web/folders`
- Headers — key `Accept`, value `application/json;odata=nometadata` · key `Content-Type`, value `application/json`

Body of each (replace `SITE_PATH`, `LIBRARY`, `ROOT` with your literal values;
the `@{…}` parts are dynamic content from Parse JSON):

1. **Ensure FY folder**
```json
{ "ServerRelativeUrl": "SITE_PATH/LIBRARY/ROOT/@{items('Apply_to_each')?['fy']}" }
```
2. **Ensure month folder**
```json
{ "ServerRelativeUrl": "SITE_PATH/LIBRARY/ROOT/@{items('Apply_to_each')?['fy']}/@{items('Apply_to_each')?['monthFolder']}" }
```
3. **Create PV folder**
```json
{ "ServerRelativeUrl": "SITE_PATH/LIBRARY/ROOT/@{items('Apply_to_each')?['subPath']}" }
```
4. **Subfolder 01_Invoice**
```json
{ "ServerRelativeUrl": "SITE_PATH/LIBRARY/ROOT/@{items('Apply_to_each')?['subPath']}/01_Invoice" }
```
5. **Subfolder 02_Approvals** — same with `/02_Approvals`, then duplicate once
more for **03_Payment_Evidence**.

*(Tip: build the first HTTP action, then use ⋯ → Copy to my clipboard → paste
and edit the path for the rest.)*

### 6. Append the link for write-back

Still inside the loop: **+ New step** → *Append to array variable* →
Name: `links` · Value:

```json
{
  "row": @{items('Apply_to_each')?['row']},
  "pvNumber": "@{items('Apply_to_each')?['pvNumber']}",
  "url": "SITE_URL/LIBRARY_ENCODED/ROOT_ENCODED/@{uriComponent(items('Apply_to_each')?['fy'])}/@{uriComponent(items('Apply_to_each')?['monthFolder'])}/@{uriComponent(items('Apply_to_each')?['folderName'])}"
}
```

Where `LIBRARY_ENCODED`/`ROOT_ENCODED` are your literal values with spaces as
`%20` — e.g. `Shared%20Documents/Payment%20Vouchers`.

### 7. Run script — Write Links (after the loop)

Outside the Apply-to-each: **+ New step** → **Excel Online (Business)** →
**Run script** →
- File: the same workbook
- Script: **PV Folders – Write Links**
- linksJson (expression):

```
concat('{"links":', string(variables('links')), ',"runId":"', workflow()?['run']?['name'], '","dryRun":', if(variables('dryRun'), 'true', 'false'), '}')
```

### 8. (Optional) Notify when folders were created

**+ New step** → *Condition* → `length(variables('links'))` **is greater than** `0` →
If yes: **Office 365 Outlook → Send an email (V2)** to the finance team:
subject `PV folders created`, body: the `links` variable. (Skip this if you
prefer checking the 🔁 Automation Log sheet.)

### 9. (Optional) Manual "run now" button

Save a copy of the flow (⋯ → Save As), open the copy and replace the Recurrence
trigger with **Manually trigger a flow**. You then have both a scheduled flow
and an on-demand one. (Keep the Office Scripts unchanged — they are shared.)

---

## Test procedure

1. Set `dryRun` = `true`. Run the flow manually (Test → Manually).
2. Check: folders appear in the library; workbook's **🔁 Automation Log** sheet
   shows `DRY RUN (would link)` rows; column AF is untouched.
3. Set `dryRun` = `false`. Run again. AF now holds working links; the log shows
   `FOLDER LINKED`. Click two links to verify.
4. Run a third time: the run history shows `pending` = empty, nothing changes —
   the idempotency check.
5. Add a new test PV row, wait for the schedule (or run manually) and verify the
   one new folder + link.

## Failure modes & fixes

| Symptom | Cause | Fix |
|---|---|---|
| HTTP action fails `403` | Flow owner lacks edit rights on the library | Grant the owner (or a service account that owns the flow) contribute rights |
| `Run script` fails "script not found" | Script renamed or saved to another workbook context | Scripts are stored per-user by default — share them: Automate tab → script → ⋯ → *Share* into the workbook |
| Link in AF gives 404 | `LIBRARY_ENCODED`/`ROOT_ENCODED` mismatch with actual site path | Open a created folder in the browser, copy its real URL prefix into step 6 |
| Folder named `PV-2026-0001_Unknown-Payee` | Payee cell was blank when the flow ran | Fill the payee, delete the AF cell + rename the folder, re-run |
| Same PV twice in register | Duplicate PV number | See `skipped` in run history / log sheet; fix the register (audit finding F7) |
