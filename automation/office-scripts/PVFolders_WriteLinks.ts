/**
 * PV Folders – Write Links
 * ------------------------
 * Office Script for the PFAFR PV workbook.
 * Power Automate calls this AFTER creating the SharePoint folders, passing the
 * list of {row, pvNumber, url}. For each item the script:
 *   1. re-checks that column AF is still empty (idempotency guard #2 — never
 *      overwrites an existing link),
 *   2. writes the folder URL into AF as a clickable hyperlink,
 *   3. appends an entry to the "🔁 Automation Log" sheet (created on first run).
 *
 * Parameters:
 *   linksJson — JSON string: {"links":[{"row":5,"pvNumber":"PV-2026-0001",
 *                "url":"https://…"}], "runId":"<flow run id>", "dryRun":false}
 *
 * Returns JSON: { written: n, alreadyLinked: n, errors: [...] }
 *
 * To install: Automate tab → New Script → paste → rename to
 * "PV Folders – Write Links" → Save.
 */

const SHEET_REGISTER = "📝 PV Register";
const SHEET_LOG = "🔁 Automation Log";
const COL_AF_INDEX = 31; // 0-based column index of AF
const LOG_HEADERS = ["Timestamp (UTC)", "PV Number", "Register Row", "Action", "Folder URL", "Flow Run"];

interface LinkItem {
  row: number;
  pvNumber: string;
  url: string;
}

function main(workbook: ExcelScript.Workbook, linksJson: string): string {
  let payload: { links?: LinkItem[]; runId?: string; dryRun?: boolean };
  try {
    payload = JSON.parse(linksJson) as { links?: LinkItem[]; runId?: string; dryRun?: boolean };
  } catch (e) {
    return JSON.stringify({ written: 0, alreadyLinked: 0, errors: ["linksJson is not valid JSON"] });
  }
  const links = payload.links ?? [];
  const runId = payload.runId ?? "manual";
  const dryRun = payload.dryRun === true;

  const sheet = workbook.getWorksheet(SHEET_REGISTER);
  if (!sheet) {
    return JSON.stringify({ written: 0, alreadyLinked: 0, errors: [`Sheet "${SHEET_REGISTER}" not found`] });
  }
  const log = getOrCreateLogSheet(workbook);
  const now = new Date().toISOString();

  let written = 0;
  let alreadyLinked = 0;
  const errors: string[] = [];

  for (const item of links) {
    if (!item || typeof item.row !== "number" || !item.url) {
      errors.push(`Malformed item: ${JSON.stringify(item)}`);
      continue;
    }
    const cell = sheet.getRangeByIndexes(item.row - 1, COL_AF_INDEX, 1, 1);
    const current = String(cell.getValues()[0][0] ?? "").trim();
    if (current !== "") {
      alreadyLinked++;
      appendLog(log, [now, item.pvNumber, item.row, "SKIPPED (already linked)", current, runId]);
      continue;
    }
    if (dryRun) {
      appendLog(log, [now, item.pvNumber, item.row, "DRY RUN (would link)", item.url, runId]);
      continue;
    }
    cell.setHyperlink({ address: item.url, textToDisplay: item.url });
    written++;
    appendLog(log, [now, item.pvNumber, item.row, "FOLDER LINKED", item.url, runId]);
  }

  return JSON.stringify({ written: written, alreadyLinked: alreadyLinked, errors: errors });
}

function getOrCreateLogSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  let log = workbook.getWorksheet(SHEET_LOG);
  if (!log) {
    log = workbook.addWorksheet(SHEET_LOG);
    const header = log.getRangeByIndexes(0, 0, 1, LOG_HEADERS.length);
    header.setValues([LOG_HEADERS]);
    header.getFormat().getFont().setBold(true);
    log.getRange("A:A").getFormat().setColumnWidth(160);
    log.getRange("E:E").getFormat().setColumnWidth(320);
  }
  return log;
}

function appendLog(log: ExcelScript.Worksheet, rowValues: (string | number)[]): void {
  const used = log.getUsedRange();
  const nextRow = used ? used.getRowCount() : 1;
  log.getRangeByIndexes(nextRow, 0, 1, rowValues.length).setValues([rowValues]);
}
