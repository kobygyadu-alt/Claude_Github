/**
 * PV Folders – Get Pending
 * ------------------------
 * Office Script for the PFAFR PV workbook.
 * Scans 📝 PV Register rows 5–510 and returns (as a JSON string) every row that
 * has a PV Number but an empty "Supporting Document (URL)" cell (column AF).
 * Power Automate runs this first, then creates the folders, then calls
 * PVFolders_WriteLinks to write the URLs back.
 *
 * Returns JSON: { pending: [...], skipped: [...], scannedRows: n }
 * Each pending item: { row, pvNumber, payee, dateIso, fy, monthFolder,
 *                      folderName, subPath }
 *
 * To install: Excel (web or desktop) → Automate tab → New Script → paste this
 * whole file → rename the script to "PV Folders – Get Pending" → Save.
 */

const SHEET_REGISTER = "📝 PV Register";
const FIRST_DATA_ROW = 5;   // 1-based, first PV row
const LAST_DATA_ROW = 510;  // 1-based, last validated row
const COL_PV = 0;           // A  (0-based offsets within A:AF)
const COL_DATE = 1;         // B
const COL_PAYEE = 4;        // E
const COL_DOCURL = 31;      // AF
const MAX_FOLDER_NAME = 80;

const MONTH_FOLDERS = [
  "01-January", "02-February", "03-March", "04-April", "05-May", "06-June",
  "07-July", "08-August", "09-September", "10-October", "11-November", "12-December"
];

function main(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet(SHEET_REGISTER);
  if (!sheet) {
    return JSON.stringify({ error: `Sheet "${SHEET_REGISTER}" not found`, pending: [], skipped: [] });
  }

  const rowCount = LAST_DATA_ROW - FIRST_DATA_ROW + 1;
  const range = sheet.getRangeByIndexes(FIRST_DATA_ROW - 1, 0, rowCount, 32); // A5:AF510
  const values = range.getValues();

  const pending: object[] = [];
  const skipped: object[] = [];
  const seenPv: { [pv: string]: number } = {};

  for (let i = 0; i < values.length; i++) {
    const rowNum = FIRST_DATA_ROW + i;
    const pv = String(values[i][COL_PV] ?? "").trim();
    if (pv === "") continue;

    // Duplicate PV numbers would collide in folder naming — log and skip.
    if (seenPv[pv] !== undefined) {
      skipped.push({ row: rowNum, pvNumber: pv, reason: `Duplicate of row ${seenPv[pv]}` });
      continue;
    }
    seenPv[pv] = rowNum;

    const docUrl = String(values[i][COL_DOCURL] ?? "").trim();
    if (docUrl !== "") continue; // already linked — idempotency guard #1

    const payee = String(values[i][COL_PAYEE] ?? "").trim();
    const dateVal = values[i][COL_DATE];
    const date = excelDateToJs(dateVal);
    if (!date) {
      skipped.push({ row: rowNum, pvNumber: pv, reason: "Missing or invalid date" });
      continue;
    }

    const folderName = sanitizeFolderName(`${pv}_${payee || "Unknown-Payee"}`);
    pending.push({
      row: rowNum,
      pvNumber: pv,
      payee: payee,
      dateIso: date.toISOString().substring(0, 10),
      fy: `FY${date.getUTCFullYear()}`,
      monthFolder: MONTH_FOLDERS[date.getUTCMonth()],
      folderName: folderName,
      subPath: `FY${date.getUTCFullYear()}/${MONTH_FOLDERS[date.getUTCMonth()]}/${folderName}`
    });
  }

  return JSON.stringify({ pending: pending, skipped: skipped, scannedRows: rowCount });
}

/** Excel stores dates as serial numbers (days since 1899-12-30). */
function excelDateToJs(v: string | number | boolean): Date | null {
  if (typeof v === "number" && v > 0) {
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (typeof v === "string" && v.trim() !== "") {
    const parsed = new Date(v);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

/** Strip characters SharePoint forbids in folder names; keep it readable. */
function sanitizeFolderName(name: string): string {
  let out = name
    .replace(/["*:<>?/\\|#%&{}~]/g, " ") // illegal / troublesome characters
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "")              // no trailing dots or spaces
    .replace(/ /g, "-");
  if (out.length > MAX_FOLDER_NAME) {
    out = out.substring(0, MAX_FOLDER_NAME).replace(/[-_.]+$/g, "");
  }
  return out;
}
