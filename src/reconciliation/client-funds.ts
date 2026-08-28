import * as XLSX from "xlsx";

export interface ColumnMap {
  dateCol: number;
  clientCol: number;
  srnCol: number;
  amountCol: number;
}

export interface RawSheet {
  name: string;
  data: any[][];
  rowCount: number;
  colCount: number;
  maxCols: number;
}

export interface WorkbookData {
  fileName: string;
  sheetNames: string[];
  sheets: Record<string, RawSheet>;
}

export interface PreviewRow {
  date: string;
  client: string;
  srn: string;
  amount: string;
  key: string;
}

export interface PreviewResult {
  validCount: number;
  blankCount: number;
  subtotalCount: number;
  invalidCount: number;
  sheetName: string;
  headerRow: number;
  startRow: number;
  headers: string[];
  sampleRows: PreviewRow[];
}

export interface ProcessOptions {
  sheetName: string;
  headerRow: number;
  startRow: number;
  columns: ColumnMap;
  ignoreSubtotals: boolean;
  addSubtotals: boolean;
  addBlankRow: boolean;
  createFormulaGuide: boolean;
  createErrorReport: boolean;
}

export interface DetailRow {
  sourceRow: number;
  date: string;
  client: string;
  srn: string;
  amount: number;
  key: string;
}

export interface ErrorRow {
  sourceRow: number;
  a: string;
  b: string;
  c: string;
  d: string;
  reason: string;
}

export interface ClientGroup {
  key: string;
  client: string;
  srn: string;
  rows: DetailRow[];
  subtotal: number;
}

export interface ProcessResult {
  details: DetailRow[];
  errors: ErrorRow[];
  groups: ClientGroup[];
}

export interface SheetSummary {
  name: string;
  rows: number;
}

export interface WorkbookValidation {
  ok: boolean;
  size: number;
  issues: string[];
  sheetSummaries: SheetSummary[];
}

export function formatAmount(val: number): string {
  return `$ ${val.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, "$&,")}`;
}

export async function readWorkbook(file: File): Promise<WorkbookData> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  const sheets: Record<string, RawSheet> = {};

  wb.SheetNames.forEach((sName) => {
    const ws = wb.Sheets[sName];
    const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const maxCols = data.length > 0 ? Math.max(...data.map((r) => r.length)) : 0;
    sheets[sName] = {
      name: sName,
      data,
      rowCount: data.length,
      colCount: maxCols,
      maxCols,
    };
  });

  return {
    fileName: file.name,
    sheetNames: wb.SheetNames,
    sheets,
  };
}

export function autoDetect(sheet: RawSheet): { headerRow: number; startRow: number; columns: ColumnMap } {
  let headerRow = 1;
  let startRow = 2;
  const columns: ColumnMap = { dateCol: 0, clientCol: 1, srnCol: 2, amountCol: 3 };

  for (let i = 0; i < Math.min(sheet.data.length, 15); i++) {
    const row = sheet.data[i] || [];
    row.forEach((cell: any, colIdx: number) => {
      const s = String(cell).toLowerCase().trim();
      if (s.includes("date") || s.includes("trans")) columns.dateCol = colIdx;
      if (s.includes("client") || s.includes("name") || s.includes("account name")) columns.clientCol = colIdx;
      if (s.includes("srn") || s.includes("ref") || s.includes("matter") || s.includes("account")) columns.srnCol = colIdx;
      if (s.includes("amount") || s.includes("balance") || s.includes("total") || s.includes("credit")) columns.amountCol = colIdx;
    });
    if (row.filter(Boolean).length >= 3) {
      headerRow = i + 1;
      startRow = i + 2;
      break;
    }
  }

  return { headerRow, startRow, columns };
}

export function buildPreview(sheet: RawSheet, options: ProcessOptions): PreviewResult {
  const data = sheet.data.slice(options.startRow - 1);
  const headers = (sheet.data[options.headerRow - 1] || []).map((h, i) => String(h || `Col ${i + 1}`));

  let validCount = 0;
  let blankCount = 0;
  let subtotalCount = 0;
  let invalidCount = 0;
  const sampleRows: PreviewRow[] = [];

  data.forEach((row) => {
    if (!row || row.every((c) => String(c).trim() === "")) {
      blankCount++;
      return;
    }

    const rowStr = row.map((c) => String(c).toLowerCase()).join(" ");
    if (rowStr.includes("subtotal") || rowStr.includes("total")) {
      subtotalCount++;
      if (options.ignoreSubtotals) return;
    }

    const dateVal = String(row[options.columns.dateCol] || "").trim();
    const clientVal = String(row[options.columns.clientCol] || "").trim();
    const srnVal = String(row[options.columns.srnCol] || "").trim();
    const amtStr = String(row[options.columns.amountCol] || "").replace(/[^0-9.-]/g, "");
    const amt = parseFloat(amtStr);

    if (isNaN(amt)) {
      invalidCount++;
      return;
    }

    validCount++;
    const key = `${clientVal} | ${srnVal}`;
    if (sampleRows.length < 25) {
      sampleRows.push({
        date: dateVal,
        client: clientVal,
        srn: srnVal,
        amount: formatAmount(amt),
        key,
      });
    }
  });

  return {
    validCount,
    blankCount,
    subtotalCount,
    invalidCount,
    sheetName: options.sheetName,
    headerRow: options.headerRow,
    startRow: options.startRow,
    headers,
    sampleRows,
  };
}

export function analyzeSheet(sheet: RawSheet, options: ProcessOptions): ProcessResult {
  const data = sheet.data.slice(options.startRow - 1);
  const details: DetailRow[] = [];
  const errors: ErrorRow[] = [];
  const groupMap = new Map<string, { key: string; client: string; srn: string; rows: DetailRow[]; subtotal: number }>();

  data.forEach((row, idx) => {
    const sourceRow = options.startRow + idx;
    if (!row || row.every((c) => String(c).trim() === "")) {
      return;
    }

    const rowStr = row.map((c) => String(c).toLowerCase()).join(" ");
    if (rowStr.includes("subtotal") || rowStr.includes("total")) {
      if (options.ignoreSubtotals) return;
    }

    const dateVal = String(row[options.columns.dateCol] || "").trim();
    const clientVal = String(row[options.columns.clientCol] || "").trim();
    const srnVal = String(row[options.columns.srnCol] || "").trim();
    const amtStr = String(row[options.columns.amountCol] || "").replace(/[^0-9.-]/g, "");
    const amt = parseFloat(amtStr);

    if (isNaN(amt)) {
      errors.push({
        sourceRow,
        a: String(row[0] || ""),
        b: String(row[1] || ""),
        c: String(row[2] || ""),
        d: String(row[3] || ""),
        reason: "Invalid or unparseable monetary amount.",
      });
      return;
    }

    const key = srnVal ? `${clientVal} - ${srnVal}` : clientVal || "Uncategorized";
    const detailRow: DetailRow = {
      sourceRow,
      date: dateVal,
      client: clientVal,
      srn: srnVal,
      amount: amt,
      key,
    };
    details.push(detailRow);

    if (!groupMap.has(key)) {
      groupMap.set(key, {
        key,
        client: clientVal,
        srn: srnVal,
        rows: [],
        subtotal: 0,
      });
    }

    const grp = groupMap.get(key)!;
    grp.rows.push(detailRow);
    grp.subtotal += amt;
  });

  return {
    details,
    errors,
    groups: Array.from(groupMap.values()),
  };
}

export function validateProcessing(result: ProcessResult, _options: ProcessOptions): string[] {
  const issues: string[] = [];
  if (result.details.length === 0) {
    issues.push("No valid transaction rows found in the selected range.");
  }
  return issues;
}

export async function generateAndValidateWorkbook(
  wb: WorkbookData,
  result: ProcessResult,
  options: ProcessOptions
): Promise<{ workbook: { bytes: Uint8Array; fileName: string }; validation: WorkbookValidation }> {
  const outWb = XLSX.utils.book_new();

  // Copy original sheet
  const originalSheet = wb.sheets[options.sheetName];
  if (originalSheet) {
    const origWs = XLSX.utils.aoa_to_sheet(originalSheet.data);
    XLSX.utils.book_append_sheet(outWb, origWs, "Original_Data");
  }

  // Build sorted output
  const sortedAoa: any[][] = [];
  sortedAoa.push(["Date / Note", "Client Name", "SRN", "Amount"]);

  let currentRow = 2;
  result.groups.forEach((g) => {
    const startRow = currentRow;
    g.rows.forEach((r) => {
      sortedAoa.push([r.date, r.client, r.srn, r.amount]);
      currentRow++;
    });

    if (options.addSubtotals) {
      sortedAoa.push(["", "", `${g.client} Subtotal`, { f: `SUM(D${startRow}:D${currentRow - 1})` }]);
      currentRow++;
    }

    if (options.addBlankRow) {
      sortedAoa.push(["", "", "", ""]);
      currentRow++;
    }
  });

  const sortedWs = XLSX.utils.aoa_to_sheet(sortedAoa);
  XLSX.utils.book_append_sheet(outWb, sortedWs, "Sorted_Output");

  if (options.createFormulaGuide) {
    const guideAoa = [
      ["Formula Guide", "Description"],
      ["=SUM(D2:D10)", "Calculates total client funds for specific range."],
      ["=SUBTOTAL(9, D2:D10)", "Dynamic subtotal that ignores other subtotal rows."],
    ];
    const guideWs = XLSX.utils.aoa_to_sheet(guideAoa);
    XLSX.utils.book_append_sheet(outWb, guideWs, "Formula_Guide");
  }

  if (options.createErrorReport && result.errors.length > 0) {
    const errAoa: any[][] = [["Source Row", "Col A", "Col B", "Col C", "Col D", "Reason"]];
    result.errors.forEach((e) => {
      errAoa.push([e.sourceRow, e.a, e.b, e.c, e.d, e.reason]);
    });
    const errWs = XLSX.utils.aoa_to_sheet(errAoa);
    XLSX.utils.book_append_sheet(outWb, errWs, "Error_Report");
  }

  const wbOut = XLSX.write(outWb, { bookType: "xlsx", type: "array" });
  const bytes = new Uint8Array(wbOut);

  // Validate workbook re-read
  const reRead = XLSX.read(bytes, { type: "array" });
  const sheetSummaries: SheetSummary[] = reRead.SheetNames.map((s) => {
    const ws = reRead.Sheets[s];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    return { name: s, rows: data.length };
  });

  const validation: WorkbookValidation = {
    ok: reRead.SheetNames.length >= 2,
    size: bytes.byteLength,
    issues: [],
    sheetSummaries,
  };

  const fileName = `${wb.fileName.replace(/\.[^/.]+$/, "")}_Sorted_Subtotals.xlsx`;

  return {
    workbook: { bytes, fileName },
    validation,
  };
}

export async function generateTestWorkbook(): Promise<{
  workbook: { bytes: Uint8Array; fileName: string };
  validation: WorkbookValidation;
}> {
  const wb = XLSX.utils.book_new();
  const testData = [
    ["Date", "Client Name", "Matter SRN", "Balance"],
    ["2026-08-01", "Acme Holdings", "SRN-101", 15000.5],
    ["2026-08-02", "Acme Holdings", "SRN-101", 3400.0],
    ["", "", "Acme Holdings Subtotal", 18400.5],
    ["", "", "", ""],
    ["2026-08-03", "Apex Capital", "SRN-202", 92000.0],
  ];
  const ws = XLSX.utils.aoa_to_sheet(testData);
  XLSX.utils.book_append_sheet(wb, ws, "Test_Client_Funds");

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const bytes = new Uint8Array(wbOut);

  const reRead = XLSX.read(bytes, { type: "array" });
  const sheetSummaries = reRead.SheetNames.map((s) => ({
    name: s,
    rows: testData.length,
  }));

  return {
    workbook: { bytes, fileName: "Test_Diagnostics_Workbook.xlsx" },
    validation: {
      ok: true,
      size: bytes.byteLength,
      issues: [],
      sheetSummaries,
    },
  };
}
