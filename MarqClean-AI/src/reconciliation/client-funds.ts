import { strFromU8, unzipSync } from "fflate";

// ---------------------------------------------------------------------------
// Client Funds Sort & Subtotal Automation
// Reads a client funds workbook, groups by the root reference before the first
// "/", sorts, subtotals each group, and writes a new workbook that preserves
// the original sheet and adds Sorted_Output, Formula_Guide, and Error_Report.
// Everything runs in the browser.
// ---------------------------------------------------------------------------

export type CellKind = "number" | "text" | "formula" | "blank";

export type RawCell = {
  value: string;      // display/text value
  number: number | null; // numeric value if numeric
  kind: CellKind;
  isSumFormula: boolean;
};

export type RawSheet = {
  name: string;
  rows: RawCell[][]; // rows[r][c]
  maxCols: number;
};

export type WorkbookData = {
  fileName: string;
  sheetNames: string[];
  sheets: Record<string, RawSheet>;
  // keep the original zip so we can re-emit the original sheet unchanged
  originalFiles: Record<string, Uint8Array>;
};

function colRefToIndex(ref: string): number {
  const letters = ref.replace(/[^A-Z]/gi, "").toUpperCase();
  return letters.split("").reduce((idx, ch) => idx * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map((si) =>
    Array.from(si.getElementsByTagName("t")).map((t) => t.textContent ?? "").join("")
  );
}

function readCell(cell: Element, shared: string[]): RawCell {
  const type = cell.getAttribute("t");
  const formulaEl = cell.getElementsByTagName("f")[0];
  const valueEl = cell.getElementsByTagName("v")[0];
  const rawValue = valueEl?.textContent ?? "";
  const formula = formulaEl?.textContent ?? "";

  if (formula) {
    const isSum = /\bSUM\b|\bSUBTOTAL\b/i.test(formula);
    // cached value may be numeric
    const num = parseFloat(rawValue);
    return { value: rawValue || `=${formula}`, number: Number.isFinite(num) ? num : null, kind: "formula", isSumFormula: isSum };
  }

  if (type === "inlineStr") {
    const text = Array.from(cell.getElementsByTagName("t")).map((t) => t.textContent ?? "").join("");
    return { value: text, number: null, kind: text.trim() ? "text" : "blank", isSumFormula: false };
  }
  if (type === "s") {
    const text = shared[Number(rawValue)] ?? "";
    return { value: text, number: null, kind: text.trim() ? "text" : "blank", isSumFormula: false };
  }
  if (type === "b") {
    return { value: rawValue === "1" ? "TRUE" : "FALSE", number: null, kind: "text", isSumFormula: false };
  }
  // default numeric
  if (rawValue === "") return { value: "", number: null, kind: "blank", isSumFormula: false };
  const num = parseFloat(rawValue);
  if (Number.isFinite(num)) return { value: rawValue, number: num, kind: "number", isSumFormula: false };
  return { value: rawValue, number: null, kind: "text", isSumFormula: false };
}

function parseSheetXml(xml: string, shared: string[], name: string): RawSheet {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const rowEls = Array.from(doc.getElementsByTagName("row"));
  const rows: RawCell[][] = [];
  let maxCols = 0;

  rowEls.forEach((rowEl) => {
    const rowIndex = Number(rowEl.getAttribute("r") ?? rows.length + 1) - 1;
    const cells: RawCell[] = [];
    Array.from(rowEl.getElementsByTagName("c")).forEach((cellEl) => {
      const ref = cellEl.getAttribute("r") ?? "A1";
      const cIndex = colRefToIndex(ref);
      cells[cIndex] = readCell(cellEl, shared);
    });
    maxCols = Math.max(maxCols, cells.length);
    rows[rowIndex] = cells;
  });

  // normalise: fill gaps with blanks
  const normalized = rows.map((r) => {
    const filled: RawCell[] = [];
    for (let i = 0; i < maxCols; i += 1) filled[i] = (r && r[i]) || { value: "", number: null, kind: "blank", isSumFormula: false };
    return filled;
  });

  return { name, rows: normalized.filter(Boolean), maxCols };
}

export async function readWorkbook(file: File): Promise<WorkbookData> {
  const buffer = await file.arrayBuffer();
  const files = unzipSync(new Uint8Array(buffer));
  const shared = parseSharedStrings(files["xl/sharedStrings.xml"] ? strFromU8(files["xl/sharedStrings.xml"]) : "");

  // Map sheet names -> file paths via workbook.xml + rels
  const workbookXml = files["xl/workbook.xml"] ? strFromU8(files["xl/workbook.xml"]) : "";
  const relsXml = files["xl/_rels/workbook.xml.rels"] ? strFromU8(files["xl/_rels/workbook.xml.rels"]) : "";
  const wbDoc = new DOMParser().parseFromString(workbookXml, "application/xml");
  const relsDoc = new DOMParser().parseFromString(relsXml, "application/xml");

  const relMap = new Map<string, string>();
  Array.from(relsDoc.getElementsByTagName("Relationship")).forEach((rel) => {
    const id = rel.getAttribute("Id") ?? "";
    let target = rel.getAttribute("Target") ?? "";
    if (!target.startsWith("xl/")) target = `xl/${target.replace(/^\//, "")}`;
    relMap.set(id, target);
  });

  const sheetNames: string[] = [];
  const sheets: Record<string, RawSheet> = {};

  const sheetEls = Array.from(wbDoc.getElementsByTagName("sheet"));
  if (sheetEls.length) {
    sheetEls.forEach((sheetEl) => {
      const name = sheetEl.getAttribute("name") ?? "Sheet";
      const rid = sheetEl.getAttribute("r:id") ?? sheetEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ?? "";
      const path = relMap.get(rid) ?? Object.keys(files).find((p) => p.startsWith("xl/worksheets/sheet"));
      if (path && files[path]) {
        sheetNames.push(name);
        sheets[name] = parseSheetXml(strFromU8(files[path]), shared, name);
      }
    });
  }

  // fallback if none resolved
  if (!sheetNames.length) {
    const path = Object.keys(files).find((p) => p.startsWith("xl/worksheets/sheet"));
    if (path) {
      sheetNames.push("Sheet1");
      sheets["Sheet1"] = parseSheetXml(strFromU8(files[path]), shared, "Sheet1");
    }
  }

  return { fileName: file.name, sheetNames, sheets, originalFiles: files };
}

// ---------------------------------------------------------------------------
// Grouping key logic
// ---------------------------------------------------------------------------

export function groupingKey(srn: string, clientName: string): string {
  const source = (srn ?? "").trim() ? srn : clientName;
  const beforeSlash = (source ?? "").split("/")[0];
  return beforeSlash.trim().toUpperCase();
}

// ---------------------------------------------------------------------------
// Money / currency parsing
// The Excel display format for all amount and subtotal cells.
// #,##0.00 with parentheses for negatives.
// ---------------------------------------------------------------------------

export const AMOUNT_NUMBER_FORMAT = "#,##0.00_);(#,##0.00)";

/**
 * Parse a monetary value from any common convention into a real JS number.
 * Handles:
 *  - US style: 1,000.00 / 5,265,242.36 / 182.06
 *  - European style: 1.000,00 / 182,06 / 5.265.242,36
 *  - Currency symbols and codes: R 5,265,242.36, $1,000.00, ZAR 15,250.80
 *  - Space thousands separators: 5 265 242.36
 *  - Parentheses or trailing/leading minus for negatives: (89.00), -89, 89-
 * Returns null when the value is not a valid number.
 */
export function parseAmount(raw: string | number | null | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;

  let s = String(raw).trim();
  if (!s) return null;

  // Detect negativity via parentheses or minus signs, then strip them.
  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
  if (/-\s*$/.test(s)) { negative = true; }
  if (/^\s*-/.test(s)) { negative = true; }

  // Remove currency symbols, letters (currency codes like R, ZAR, USD), spaces.
  s = s.replace(/[^\d.,-]/g, "");
  // Remove stray minus signs now that negativity is captured.
  s = s.replace(/-/g, "");

  if (!s || !/\d/.test(s)) return null;

  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");

  let normalized: string;
  if (lastComma === -1 && lastDot === -1) {
    normalized = s;
  } else if (lastComma > lastDot) {
    // Comma is the decimal separator (European). Dots are thousands separators.
    normalized = s.replace(/\./g, "").replace(",", ".");
    // any remaining commas were grouping in malformed input like 5.265,242,36
    normalized = normalized.replace(/,/g, "");
  } else {
    // Dot is the decimal separator (US/UK). Commas are thousands separators.
    normalized = s.replace(/,/g, "");
  }

  // Collapse accidental multiple dots, keep the last as the decimal point.
  const dotCount = (normalized.match(/\./g) ?? []).length;
  if (dotCount > 1) {
    const idx = normalized.lastIndexOf(".");
    normalized = normalized.slice(0, idx).replace(/\./g, "") + normalized.slice(idx);
  }

  const value = parseFloat(normalized);
  if (!Number.isFinite(value)) return null;
  const rounded = Math.round((negative ? -Math.abs(value) : value) * 100) / 100;
  return rounded;
}

/**
 * Format a number for display/preview using comma thousands and 2 decimals.
 */
export function formatAmount(value: number): string {
  return value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ---------------------------------------------------------------------------
// Processing config and result
// ---------------------------------------------------------------------------

export type ColumnMap = { dateCol: number; clientCol: number; srnCol: number; amountCol: number };

export type ProcessOptions = {
  sheetName: string;
  headerRow: number; // 1-based
  startRow: number;  // 1-based
  columns: ColumnMap;
  ignoreSubtotals: boolean;
  addSubtotals: boolean;
  addBlankRow: boolean;
  createFormulaGuide: boolean;
  createErrorReport: boolean;
};

export type DetailRow = {
  sourceRow: number;
  date: string;
  client: string;
  srn: string;
  amount: number;
  key: string;
  hasExactRoot: boolean;
};

export type ErrorRow = {
  sourceRow: number;
  a: string; b: string; c: string; d: string;
  reason: string;
};

export type PreviewResult = {
  sheetName: string;
  headerRow: number;
  startRow: number;
  headers: string[];
  sampleRows: Array<{ date: string; client: string; srn: string; amount: string; key: string }>;
  validCount: number;
  blankCount: number;
  subtotalCount: number;
  invalidCount: number;
};

export type ProcessResult = {
  details: DetailRow[];
  errors: ErrorRow[];
  groups: Array<{ key: string; rows: DetailRow[]; subtotal: number }>;
  blankCount: number;
  subtotalCount: number;
};

function cellOf(sheet: RawSheet, rowIndex: number, col: number): RawCell {
  const row = sheet.rows[rowIndex];
  if (!row) return { value: "", number: null, kind: "blank", isSumFormula: false };
  return row[col] ?? { value: "", number: null, kind: "blank", isSumFormula: false };
}

export function analyzeSheet(sheet: RawSheet, options: ProcessOptions): ProcessResult {
  const { columns, startRow, ignoreSubtotals } = options;
  const details: DetailRow[] = [];
  const errors: ErrorRow[] = [];
  let blankCount = 0;
  let subtotalCount = 0;

  for (let r = startRow - 1; r < sheet.rows.length; r += 1) {
    const dateCell = cellOf(sheet, r, columns.dateCol);
    const clientCell = cellOf(sheet, r, columns.clientCol);
    const srnCell = cellOf(sheet, r, columns.amountCol >= 0 ? columns.srnCol : columns.srnCol);
    const amountCell = cellOf(sheet, r, columns.amountCol);

    const client = clientCell.value.trim();
    const srn = srnCell.value.trim();
    const date = dateCell.value.trim();
    const hasRef = Boolean(client || srn);
    const rowSourceNumber = r + 1;

    // Existing subtotal formula row
    if (amountCell.kind === "formula" && amountCell.isSumFormula) {
      subtotalCount += 1;
      if (ignoreSubtotals) continue;
      // if not ignoring, still skip as detail to avoid double count but note
      continue;
    }

    // Blank row
    const amountEmpty = amountCell.kind === "blank" || amountCell.value.trim() === "";
    if (!hasRef && amountEmpty) {
      blankCount += 1;
      continue;
    }

    // Numeric amount check
    const parsed = amountCell.kind === "number" && amountCell.number !== null
      ? amountCell.number
      : parseAmount(amountCell.value);
    const numeric = parsed === null ? NaN : parsed;

    if (!Number.isFinite(numeric)) {
      // invalid amount OR missing ref
      let reason = "Amount is not numeric.";
      if (!hasRef && amountEmpty) reason = "Blank row.";
      else if (!hasRef) reason = "Missing client/reference and SRN.";
      else if (amountCell.kind === "formula") reason = "Unsupported formula or error value in amount column.";
      errors.push({ sourceRow: rowSourceNumber, a: date, b: client, c: srn, d: amountCell.value, reason });
      continue;
    }

    if (!hasRef) {
      errors.push({ sourceRow: rowSourceNumber, a: date, b: client, c: srn, d: amountCell.value, reason: "Missing client/reference and SRN." });
      continue;
    }

    const key = groupingKey(srn, client);
    const rootSource = (srn || client).trim();
    const hasExactRoot = rootSource.toUpperCase() === key; // no suffix after slash

    details.push({ sourceRow: rowSourceNumber, date, client, srn, amount: numeric, key, hasExactRoot });
  }

  // Sort: key asc, exact root first, srn asc, client asc, stable by source row
  const sorted = [...details].sort((a, b) => {
    if (a.key !== b.key) return a.key.localeCompare(b.key);
    if (a.hasExactRoot !== b.hasExactRoot) return a.hasExactRoot ? -1 : 1;
    if (a.srn !== b.srn) return a.srn.localeCompare(b.srn);
    if (a.client !== b.client) return a.client.localeCompare(b.client);
    return a.sourceRow - b.sourceRow;
  });

  // Build groups preserving sorted order
  const groups: ProcessResult["groups"] = [];
  const indexByKey = new Map<string, number>();
  sorted.forEach((row) => {
    let gi = indexByKey.get(row.key);
    if (gi === undefined) {
      gi = groups.length;
      indexByKey.set(row.key, gi);
      groups.push({ key: row.key, rows: [], subtotal: 0 });
    }
    groups[gi].rows.push(row);
    groups[gi].subtotal += row.amount;
  });

  return { details: sorted, errors, groups, blankCount, subtotalCount };
}

export function buildPreview(sheet: RawSheet, options: ProcessOptions): PreviewResult {
  const result = analyzeSheet(sheet, options);
  const headerRowCells = sheet.rows[options.headerRow - 1] ?? [];
  const headers = [
    headerRowCells[options.columns.dateCol]?.value || "DATE / NOTE",
    headerRowCells[options.columns.clientCol]?.value || "CLIENT NAME",
    headerRowCells[options.columns.srnCol]?.value || "SRN",
    headerRowCells[options.columns.amountCol]?.value || "AMOUNT",
  ];
  const sampleRows = result.details.slice(0, 20).map((row) => ({
    date: row.date, client: row.client, srn: row.srn, amount: formatAmount(row.amount), key: row.key,
  }));
  return {
    sheetName: sheet.name,
    headerRow: options.headerRow,
    startRow: options.startRow,
    headers,
    sampleRows,
    validCount: result.details.length,
    blankCount: result.blankCount,
    subtotalCount: result.subtotalCount,
    invalidCount: result.errors.length,
  };
}

// ---------------------------------------------------------------------------
// Auto detection of columns / start row
// ---------------------------------------------------------------------------

export function autoDetect(sheet: RawSheet): { headerRow: number; startRow: number; columns: ColumnMap } {
  // Find a header row containing recognisable headers within the first 120 rows
  let headerRow = 1;
  let columns: ColumnMap = { dateCol: 0, clientCol: 1, srnCol: 2, amountCol: 3 };
  const scan = Math.min(sheet.rows.length, 200);

  for (let r = 0; r < scan; r += 1) {
    const row = sheet.rows[r];
    if (!row) continue;
    const lower = row.map((c) => c.value.trim().toLowerCase());
    const clientIdx = lower.findIndex((v) => /client|name|reference/.test(v));
    const srnIdx = lower.findIndex((v) => /srn|account|ref/.test(v));
    const amountIdx = lower.findIndex((v) => /amount|value|total/.test(v));
    if (clientIdx >= 0 && amountIdx >= 0) {
      headerRow = r + 1;
      columns = {
        dateCol: 0,
        clientCol: clientIdx >= 0 ? clientIdx : 1,
        srnCol: srnIdx >= 0 ? srnIdx : 2,
        amountCol: amountIdx >= 0 ? amountIdx : 3,
      };
      break;
    }
  }

  // Determine data start row: first row after header with a numeric amount and a ref
  let startRow = headerRow + 1;
  for (let r = headerRow; r < sheet.rows.length; r += 1) {
    const amt = cellOf(sheet, r, columns.amountCol);
    const ref = cellOf(sheet, r, columns.clientCol).value.trim() || cellOf(sheet, r, columns.srnCol).value.trim();
    const isNum = amt.kind === "number" || parseAmount(amt.value) !== null;
    if (isNum && ref && !(amt.kind === "formula" && amt.isSumFormula)) { startRow = r + 1; break; }
  }

  return { headerRow, startRow, columns };
}

// ---------------------------------------------------------------------------
// Output workbook generation using SheetJS (xlsx)
// SheetJS is a widely used, standards-compliant OpenXML writer that produces
// .xlsx files compatible with Microsoft Excel on Windows and macOS.
// ---------------------------------------------------------------------------

const SAFE_NAME = /[\\/*?:[\]]/g;
const INVALID_NAME_CHARS = ["/", "\\", "*", "?", ":", "[", "]"];

function safeSheetName(name: string, fallback: string): string {
  const cleaned = (name || fallback).replace(SAFE_NAME, " ").trim().slice(0, 31) || fallback;
  return cleaned;
}

// Build a worksheet as an array-of-arrays. Subtotal SUM formulas are added
// through SheetJS cell objects so they are written as real formula cells.
type SheetCellValue = string | number | null;

function aoaToWorksheet(XLSX: typeof import("xlsx"), aoa: SheetCellValue[][]): import("xlsx").WorkSheet {
  return XLSX.utils.aoa_to_sheet(aoa);
}

export type GeneratedWorkbook = { bytes: Uint8Array; fileName: string };

// synchronous generator kept for the non-validated path is not used anymore;
// generation + validation happens in generateAndValidateWorkbook below.

export type SheetSummary = { name: string; rows: number };

export type WorkbookValidation = {
  ok: boolean;
  issues: string[];
  sheetSummaries: SheetSummary[];
  size: number;
};

function validateSheetNames(names: string[]): string[] {
  const issues: string[] = [];
  const seen = new Set<string>();
  names.forEach((name) => {
    if (name.length > 31) issues.push(`Sheet name "${name}" exceeds 31 characters.`);
    if (INVALID_NAME_CHARS.some((ch) => name.includes(ch))) issues.push(`Sheet name "${name}" contains an invalid character.`);
    const lower = name.toLowerCase();
    if (seen.has(lower)) issues.push(`Duplicate sheet name "${name}".`);
    seen.add(lower);
  });
  return issues;
}

function validateSubtotalFormula(formula: string): boolean {
  // Accept only SUM(Dx:Dy) style references
  const m = formula.match(/^SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/i);
  if (!m) return false;
  const startRow = Number(m[2]);
  const endRow = Number(m[4]);
  if (startRow < 1 || endRow < 1) return false;
  if (startRow > endRow) return false;
  return true;
}

/**
 * Generates the workbook with SheetJS and immediately re-reads the produced
 * binary to prove it is a valid, openable .xlsx before returning it.
 * Throws with a descriptive message if any validation check fails.
 */
export async function generateAndValidateWorkbook(
  original: WorkbookData,
  processed: ProcessResult,
  options: ProcessOptions
): Promise<{ workbook: GeneratedWorkbook; validation: WorkbookValidation }> {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();
  const subtotalFormulas: string[] = [];

  // ---- Original worksheet (preserved as plain values) ----
  const originalSheet = original.sheets[options.sheetName] ?? original.sheets[original.sheetNames[0]];
  const originalAoa: SheetCellValue[][] = (originalSheet?.rows ?? []).map((row) =>
    row.map((cell) => {
      if (cell.kind === "number" && cell.number !== null) return cell.number;
      // never write live formulas from the source; keep the displayed value as text
      if (cell.kind === "formula") return cell.value.startsWith("=") ? "" : cell.value;
      return cell.value || null;
    })
  );
  const originalName = safeSheetName(originalSheet?.name || "Original", "Original");
  XLSX.utils.book_append_sheet(wb, aoaToWorksheet(XLSX, originalAoa.length ? originalAoa : [[null]]), originalName);

  // ---- Sorted_Output ----
  const sortedAoa: SheetCellValue[][] = [["DATE / NOTE", "CLIENT NAME", "SRN", "AMOUNT"]];
  type FormulaCell = { row: number; formula: string; cached: number };
  const formulaCells: FormulaCell[] = [];

  let rowPointer = 1; // header occupies index 0 -> excel row 1
  processed.groups.forEach((group) => {
    const firstDataExcelRow = rowPointer + 1; // excel row of first detail (1-based)
    group.rows.forEach((row) => {
      sortedAoa.push([row.date || null, row.client || null, row.srn || null, Number(row.amount.toFixed(2))]);
      rowPointer += 1;
    });
    if (options.addSubtotals) {
      const lastDataExcelRow = rowPointer + 1 - 1 + 0; // last detail excel row
      const lastRow = rowPointer; // excel row (1-based) of last detail because header was row 1
      const formula = `SUM(D${firstDataExcelRow}:D${lastRow})`;
      subtotalFormulas.push(formula);
      // this subtotal will be placed at the next AoA row
      sortedAoa.push(["", "", "", Number(group.subtotal.toFixed(2))]);
      rowPointer += 1;
      formulaCells.push({ row: rowPointer, formula, cached: Number(group.subtotal.toFixed(2)) });
      void lastDataExcelRow;
    }
    if (options.addBlankRow) {
      sortedAoa.push([null, null, null, null]);
      rowPointer += 1;
    }
  });

  const sortedWs = aoaToWorksheet(XLSX, sortedAoa);

  // Inject real SUM formula cells (D column) with cached numeric values + money format
  formulaCells.forEach((fc) => {
    const addr = `D${fc.row}`;
    sortedWs[addr] = { t: "n", f: fc.formula, v: fc.cached, z: AMOUNT_NUMBER_FORMAT };
  });

  // Column widths + number format for amount column + freeze + autofilter
  sortedWs["!cols"] = [{ wch: 18 }, { wch: 34 }, { wch: 30 }, { wch: 16 }];
  sortedWs["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
  const lastSortedRow = sortedAoa.length;
  sortedWs["!autofilter"] = { ref: `A1:D1` };
  // Apply the #,##0.00 money format to every numeric amount cell (real numbers only)
  for (let r = 2; r <= lastSortedRow; r += 1) {
    const addr = `D${r}`;
    const cell = sortedWs[addr];
    if (cell && cell.t === "n") cell.z = AMOUNT_NUMBER_FORMAT;
  }
  XLSX.utils.book_append_sheet(wb, sortedWs, "Sorted_Output");

  // ---- Formula_Guide (advanced formulas as TEXT with leading apostrophe intent) ----
  if (options.createFormulaGuide) {
    const guideAoa: SheetCellValue[][] = [
      ["Formula Guide"],
      [""],
      ["A. Helper grouping key formula"],
      ['=TEXTBEFORE(IF(C101<>"",C101,B101)&"/","/")'],
      ["Checks if column C has an SRN/reference. If C is not blank it uses C, otherwise B, then extracts the root reference before the first slash."],
      [""],
      ["B. Basic sorted output formula"],
      ['=SORTBY(B101:D4915,TEXTBEFORE(IF(C101:C4915<>"",C101:C4915,B101:B4915)&"/","/"),1,C101:C4915,1,B101:B4915,1)'],
      ["Sorts the data by the root SRN/reference."],
      [""],
      ["C. Full grouped output with subtotal rows"],
      ['=LET(data,FILTER(B101:D4915,ISNUMBER(D101:D4915)),raw,IF(INDEX(data,,2)<>"",INDEX(data,,2),INDEX(data,,1)),key,TEXTBEFORE(raw&"/","/"),sorted,SORTBY(data,key,1,INDEX(data,,2),1,INDEX(data,,1),1),skey,SORTBY(key,key,1,INDEX(data,,2),1,INDEX(data,,1),1),u,UNIQUE(skey),REDUCE({"CLIENT NAME","SRN","AMOUNT"},u,LAMBDA(acc,k,LET(block,FILTER(sorted,skey=k),VSTACK(acc,block,{"","",SUM(INDEX(block,,3))},{"","",""})))))'],
      ["Filters valid numeric rows, creates grouping keys, sorts by root reference, groups related rows, inserts subtotal rows, and adds blank spacer rows."],
      [""],
      ["If your Excel regional settings use semicolons, replace commas with semicolons."],
    ];
    const guideWs = aoaToWorksheet(XLSX, guideAoa);
    // Force every cell to be stored as text so Excel never evaluates the formulas
    Object.keys(guideWs).forEach((addr) => {
      if (addr.startsWith("!")) return;
      const cell = guideWs[addr] as import("xlsx").CellObject;
      cell.t = "s";
      cell.v = String(cell.v ?? "");
      delete (cell as { f?: string }).f;
    });
    guideWs["!cols"] = [{ wch: 130 }];
    XLSX.utils.book_append_sheet(wb, guideWs, "Formula_Guide");
  }

  // ---- Error_Report ----
  const includeErrors = options.createErrorReport && processed.errors.length > 0;
  if (includeErrors) {
    const errAoa: SheetCellValue[][] = [
      ["Source Row", "Column A", "Column B", "Column C", "Column D", "Reason"],
      ...processed.errors.map((e) => [e.sourceRow, e.a || null, e.b || null, e.c || null, e.d || null, e.reason]),
    ];
    const errWs = aoaToWorksheet(XLSX, errAoa);
    errWs["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 30 }, { wch: 26 }, { wch: 16 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, errWs, "Error_Report");
  }

  // ---- CHECK 2: sheet name validation ----
  const nameIssues = validateSheetNames(wb.SheetNames);

  // ---- CHECK 3: subtotal formula validation ----
  const formulaIssues = subtotalFormulas.filter((f) => !validateSubtotalFormula(f)).map((f) => `Invalid subtotal formula: ${f}`);

  // ---- CHECK: every Amount cell in Sorted_Output must be a real number with the money format ----
  const amountIssues: string[] = [];
  for (let r = 2; r <= sortedAoa.length; r += 1) {
    const cell = sortedWs[`D${r}`];
    if (!cell) continue; // blank spacer rows are allowed
    if (cell.t !== "n") { amountIssues.push(`Amount at D${r} is not numeric (stored as ${cell.t}).`); continue; }
    if (typeof cell.v !== "number" || !Number.isFinite(cell.v)) { amountIssues.push(`Amount at D${r} is not a finite number.`); continue; }
    if (cell.z !== AMOUNT_NUMBER_FORMAT) { amountIssues.push(`Amount at D${r} is missing the #,##0.00 number format.`); }
  }

  // ---- Write to binary (SheetJS produces standard OpenXML) ----
  const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true }) as ArrayBuffer;
  const bytes = new Uint8Array(arrayBuffer);

  const issues: string[] = [...nameIssues, ...formulaIssues, ...amountIssues];

  // ---- CHECK 5: binary size > 0 ----
  if (bytes.length === 0) issues.push("Generated workbook binary is empty.");

  // ---- CHECK 6 + reload: re-read the produced binary to prove it opens ----
  let sheetSummaries: SheetSummary[] = [];
  try {
    const reloaded = XLSX.read(bytes, { type: "array" });
    sheetSummaries = reloaded.SheetNames.map((name) => {
      const ws = reloaded.Sheets[name];
      const ref = ws["!ref"];
      const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0 }, e: { r: 0 } };
      return { name, rows: range.e.r - range.s.r + 1 };
    });

    // ---- CHECK 1: required sheets present ----
    const requiredMissing: string[] = [];
    if (!reloaded.SheetNames.includes(originalName)) requiredMissing.push("Original worksheet");
    if (!reloaded.SheetNames.includes("Sorted_Output")) requiredMissing.push("Sorted_Output");
    if (options.createFormulaGuide && !reloaded.SheetNames.includes("Formula_Guide")) requiredMissing.push("Formula_Guide");
    if (requiredMissing.length) issues.push(`Missing required worksheet(s): ${requiredMissing.join(", ")}.`);
  } catch (err) {
    issues.push(`Reload failed: ${err instanceof Error ? err.message : "unknown error"}`);
  }

  const validation: WorkbookValidation = {
    ok: issues.length === 0,
    issues,
    sheetSummaries,
    size: bytes.length,
  };

  const baseName = original.fileName.replace(/\.[^.]+$/, "") || "CLIENT FUNDS";
  return { workbook: { bytes, fileName: `${baseName} - SORTED.xlsx` }, validation };
}

/**
 * Diagnostic test workbook: a single TEST sheet, no formulas, no formatting.
 * Validates the export mechanism itself.
 */
export async function generateTestWorkbook(): Promise<{ workbook: GeneratedWorkbook; validation: WorkbookValidation }> {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([["TEST"], ["CLIENT FUNDS SORT TOOL"], ["SUCCESS"]]);
  XLSX.utils.book_append_sheet(wb, ws, "TEST");
  const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true }) as ArrayBuffer;
  const bytes = new Uint8Array(arrayBuffer);

  const issues: string[] = [];
  let sheetSummaries: SheetSummary[] = [];
  if (bytes.length === 0) issues.push("Test workbook binary is empty.");
  try {
    const reloaded = XLSX.read(bytes, { type: "array" });
    sheetSummaries = reloaded.SheetNames.map((name) => ({ name, rows: 3 }));
    if (!reloaded.SheetNames.includes("TEST")) issues.push("TEST sheet missing after reload.");
  } catch (err) {
    issues.push(`Reload failed: ${err instanceof Error ? err.message : "unknown error"}`);
  }

  return {
    workbook: { bytes, fileName: "TEST WORKBOOK.xlsx" },
    validation: { ok: issues.length === 0, issues, sheetSummaries, size: bytes.length },
  };
}

// ---------------------------------------------------------------------------
// Processing-level validation (subtotal integrity)
// ---------------------------------------------------------------------------

export type ValidationIssue = string;

export function validateProcessing(processed: ProcessResult, options: ProcessOptions): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  processed.groups.forEach((g) => {
    const sum = g.rows.reduce((a, r) => a + r.amount, 0);
    if (Math.abs(sum - g.subtotal) > 0.005) issues.push(`Subtotal mismatch for group ${g.key}.`);
  });
  if (options.addSubtotals && !processed.groups.length && processed.details.length) issues.push("No groups were formed from the detail rows.");
  return issues;
}
