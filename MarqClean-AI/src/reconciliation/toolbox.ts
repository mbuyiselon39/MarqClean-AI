import Papa from "papaparse";
import {
  buildStyledWorkbookFromMatrix,
  similarity,
  standardizeName,
  type DataTable,
} from "./toolbox-helpers";

export type { DataTable } from "./toolbox-helpers";

// ---------------------------------------------------------------------------
// Merge & Combine (Merzing / Merge Spreadsheets)
// ---------------------------------------------------------------------------

export type MergeMode = "append" | "join";

function normHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Append: stack rows aligning by matching (fuzzy) column names.
export function mergeAppend(tables: DataTable[]): DataTable {
  const headers: string[] = [];
  const headerKeyToIndex = new Map<string, number>();

  tables.forEach((table) => {
    table.headers.forEach((header) => {
      const key = normHeader(header);
      if (!headerKeyToIndex.has(key)) {
        headerKeyToIndex.set(key, headers.length);
        headers.push(header);
      }
    });
  });

  const rows: string[][] = [];
  tables.forEach((table) => {
    const columnMap = table.headers.map((header) => headerKeyToIndex.get(normHeader(header)) ?? -1);
    table.rows.forEach((row) => {
      const newRow = new Array(headers.length).fill("");
      row.forEach((value, index) => {
        const target = columnMap[index];
        if (target >= 0) newRow[target] = value;
      });
      rows.push(newRow);
    });
  });

  return { headers, rows, sourceName: "merged" };
}

// Join: left join two tables on a key column from each.
export function mergeJoin(left: DataTable, right: DataTable, leftKeyCol: number, rightKeyCol: number): DataTable {
  const rightIndex = new Map<string, string[]>();
  right.rows.forEach((row) => {
    const key = (row[rightKeyCol] ?? "").trim().toLowerCase();
    if (key && !rightIndex.has(key)) rightIndex.set(key, row);
  });

  const rightHeaders = right.headers.filter((_, i) => i !== rightKeyCol);
  const headers = [...left.headers, ...rightHeaders.map((h) => `${h} (matched)`)];

  const rows = left.rows.map((row) => {
    const key = (row[leftKeyCol] ?? "").trim().toLowerCase();
    const match = rightIndex.get(key);
    const rightValues = right.headers.map((_, i) => (match ? match[i] ?? "" : "")).filter((_, i) => i !== rightKeyCol);
    return [...row, ...rightValues];
  });

  return { headers, rows, sourceName: "joined" };
}

// ---------------------------------------------------------------------------
// Remove Duplicates (Remove Duplicates in Excel / Ablebits style)
// ---------------------------------------------------------------------------

export type DuplicateMode = "remove" | "keep-first" | "flag";

export function removeDuplicatesByColumns(table: DataTable, columns: number[]): { table: DataTable; removed: number; duplicates: number } {
  if (!columns.length) columns = table.headers.map((_, i) => i);
  const seen = new Set<string>();
  const rows: string[][] = [];
  let duplicates = 0;

  table.rows.forEach((row) => {
    const key = columns.map((c) => (row[c] ?? "").trim().toLowerCase()).join("\u0001");
    if (seen.has(key)) {
      duplicates += 1;
    } else {
      seen.add(key);
      rows.push(row);
    }
  });

  return { table: { ...table, rows }, removed: duplicates, duplicates };
}

export function findFuzzyDuplicates(table: DataTable, column: number, threshold: number): Array<{ a: number; b: number; score: number; valueA: string; valueB: string }> {
  const pairs: Array<{ a: number; b: number; score: number; valueA: string; valueB: string }> = [];
  const values = table.rows.map((row) => standardizeName(row[column] ?? ""));
  const limit = Math.min(table.rows.length, 1500); // keep it responsive

  for (let i = 0; i < limit; i += 1) {
    if (!values[i]) continue;
    for (let j = i + 1; j < limit; j += 1) {
      if (!values[j]) continue;
      const score = similarity(values[i], values[j]);
      if (score >= threshold && score < 1) {
        pairs.push({ a: i, b: j, score, valueA: table.rows[i][column] ?? "", valueB: table.rows[j][column] ?? "" });
      }
    }
  }
  return pairs.sort((a, b) => b.score - a.score).slice(0, 200);
}

// ---------------------------------------------------------------------------
// Fuzzy Match / Beyond String Metrics (name & address similarity search)
// ---------------------------------------------------------------------------

export function fuzzySearch(table: DataTable, column: number, query: string, threshold: number): Array<{ rowIndex: number; value: string; score: number }> {
  const q = standardizeName(query);
  if (!q) return [];
  return table.rows
    .map((row, rowIndex) => {
      const value = row[column] ?? "";
      const score = similarity(standardizeName(value), q);
      return { rowIndex, value, score };
    })
    .filter((r) => r.score >= threshold)
    .sort((a, b) => b.score - a.score)
    .slice(0, 100);
}

// ---------------------------------------------------------------------------
// CleanSmart Labs: Data Health / Clarity Score
// ---------------------------------------------------------------------------

export type ClarityReport = {
  score: number;
  totalRows: number;
  totalCells: number;
  filledCells: number;
  duplicateRows: number;
  issues: Array<{ column: string; issue: string; count: number }>;
};

function isEmail(v: string) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.trim()); }

export function computeClarityScore(table: DataTable): ClarityReport {
  const totalRows = table.rows.length;
  const totalCells = totalRows * table.headers.length || 1;
  let filledCells = 0;
  const issues: Array<{ column: string; issue: string; count: number }> = [];

  // Duplicates (whole row)
  const seen = new Set<string>();
  let duplicateRows = 0;
  table.rows.forEach((row) => {
    const key = row.map((v) => (v ?? "").trim().toLowerCase()).join("\u0001");
    if (seen.has(key)) duplicateRows += 1;
    else seen.add(key);
  });

  table.headers.forEach((header, col) => {
    let empty = 0;
    let invalidEmail = 0;
    let inconsistentCase = 0;
    const looksEmail = /email|e-mail|mail/i.test(header);
    table.rows.forEach((row) => {
      const value = (row[col] ?? "").trim();
      if (value) filledCells += 1;
      else empty += 1;
      if (value && looksEmail && !isEmail(value)) invalidEmail += 1;
      if (value && /[a-z]/.test(value) && /[A-Z]/.test(value) && value !== value.toUpperCase() && value === value.toLowerCase()) inconsistentCase += 1;
    });
    if (empty) issues.push({ column: header, issue: "Missing values", count: empty });
    if (invalidEmail) issues.push({ column: header, issue: "Invalid email format", count: invalidEmail });
  });

  const completeness = filledCells / totalCells;
  const dupPenalty = totalRows ? duplicateRows / totalRows : 0;
  const score = Math.max(0, Math.round((completeness * 0.8 + (1 - dupPenalty) * 0.2) * 100));

  return { score, totalRows, totalCells, filledCells, duplicateRows, issues: issues.slice(0, 50) };
}

// SmartFill: fill blank cells using the most common value in the column (mode)
export function smartFill(table: DataTable, column: number): { table: DataTable; filled: number } {
  const counts = new Map<string, number>();
  table.rows.forEach((row) => {
    const v = (row[column] ?? "").trim();
    if (v) counts.set(v, (counts.get(v) ?? 0) + 1);
  });
  let mode = "";
  let best = 0;
  counts.forEach((count, value) => { if (count > best) { best = count; mode = value; } });
  if (!mode) return { table, filled: 0 };

  let filled = 0;
  const rows = table.rows.map((row) => {
    if (!(row[column] ?? "").trim()) {
      filled += 1;
      const copy = [...row];
      copy[column] = mode;
      return copy;
    }
    return row;
  });
  return { table: { ...table, rows }, filled };
}

// AutoFormat: standardize casing / trims for a column
export type FormatMode = "title" | "upper" | "lower" | "trim";

export function autoFormatColumn(table: DataTable, column: number, mode: FormatMode): DataTable {
  const rows = table.rows.map((row) => {
    const copy = [...row];
    const value = (row[column] ?? "").trim().replace(/\s+/g, " ");
    if (mode === "upper") copy[column] = value.toUpperCase();
    else if (mode === "lower") copy[column] = value.toLowerCase();
    else if (mode === "trim") copy[column] = value;
    else copy[column] = value.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
    return copy;
  });
  return { ...table, rows };
}

// ---------------------------------------------------------------------------
// Coco Alemana: Data Explorer (filter, distribution, simple query)
// ---------------------------------------------------------------------------

export function columnDistribution(table: DataTable, column: number, topN = 12): Array<{ label: string; count: number }> {
  const counts = new Map<string, number>();
  table.rows.forEach((row) => {
    const value = (row[column] ?? "").trim() || "(blank)";
    counts.set(value, (counts.get(value) ?? 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, topN);
}

export type FilterOp = "contains" | "equals" | "gt" | "lt" | "nonempty" | "empty";

export function filterTable(table: DataTable, column: number, op: FilterOp, value: string): DataTable {
  const needle = value.trim().toLowerCase();
  const num = parseFloat(value);
  const rows = table.rows.filter((row) => {
    const cell = (row[column] ?? "").trim();
    const lower = cell.toLowerCase();
    switch (op) {
      case "contains": return lower.includes(needle);
      case "equals": return lower === needle;
      case "gt": return parseFloat(cell.replace(/[^0-9.-]/g, "")) > num;
      case "lt": return parseFloat(cell.replace(/[^0-9.-]/g, "")) < num;
      case "nonempty": return cell.length > 0;
      case "empty": return cell.length === 0;
      default: return true;
    }
  });
  return { ...table, rows };
}

// ---------------------------------------------------------------------------
// Melon: Smart Predict (one-click prediction)
// ---------------------------------------------------------------------------

// Predict a numeric target from a single numeric feature using linear regression.
export function linearPredict(table: DataTable, featureCol: number, targetCol: number, inputValue: number): { prediction: number; slope: number; intercept: number; r2: number; points: number } | null {
  const xs: number[] = [];
  const ys: number[] = [];
  table.rows.forEach((row) => {
    const x = parseFloat((row[featureCol] ?? "").replace(/[^0-9.-]/g, ""));
    const y = parseFloat((row[targetCol] ?? "").replace(/[^0-9.-]/g, ""));
    if (Number.isFinite(x) && Number.isFinite(y)) { xs.push(x); ys.push(y); }
  });
  const n = xs.length;
  if (n < 2) return null;

  const meanX = xs.reduce((a, b) => a + b, 0) / n;
  const meanY = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i += 1) { num += (xs[i] - meanX) * (ys[i] - meanY); den += (xs[i] - meanX) ** 2; }
  if (den === 0) return null;
  const slope = num / den;
  const intercept = meanY - slope * meanX;

  let ssRes = 0;
  let ssTot = 0;
  for (let i = 0; i < n; i += 1) {
    const pred = slope * xs[i] + intercept;
    ssRes += (ys[i] - pred) ** 2;
    ssTot += (ys[i] - meanY) ** 2;
  }
  const r2 = ssTot === 0 ? 1 : 1 - ssRes / ssTot;

  return { prediction: slope * inputValue + intercept, slope, intercept, r2, points: n };
}

// Predict a category from a numeric feature using nearest-neighbour on 1 dimension.
export function classifyPredict(table: DataTable, featureCol: number, labelCol: number, inputValue: number): { label: string; neighbours: number } | null {
  const points: Array<{ x: number; label: string }> = [];
  table.rows.forEach((row) => {
    const x = parseFloat((row[featureCol] ?? "").replace(/[^0-9.-]/g, ""));
    const label = (row[labelCol] ?? "").trim();
    if (Number.isFinite(x) && label) points.push({ x, label });
  });
  if (points.length < 3) return null;

  points.sort((a, b) => Math.abs(a.x - inputValue) - Math.abs(b.x - inputValue));
  const k = Math.min(5, points.length);
  const votes = new Map<string, number>();
  for (let i = 0; i < k; i += 1) votes.set(points[i].label, (votes.get(points[i].label) ?? 0) + 1);
  let label = "";
  let best = 0;
  votes.forEach((count, value) => { if (count > best) { best = count; label = value; } });
  return { label, neighbours: k };
}

// ---------------------------------------------------------------------------
// Listly: HTML table extractor (paste HTML -> structured table)
// ---------------------------------------------------------------------------

export function extractHtmlTable(html: string): DataTable | null {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (table) {
    const rowEls = Array.from(table.querySelectorAll("tr"));
    const matrix = rowEls.map((tr) => Array.from(tr.querySelectorAll("th,td")).map((cell) => (cell.textContent ?? "").trim()));
    const usable = matrix.filter((row) => row.some(Boolean));
    if (!usable.length) return null;
    const headers = usable[0];
    return { headers, rows: usable.slice(1), sourceName: "extracted-table" };
  }
  // Fallback: list items
  const items = Array.from(doc.querySelectorAll("li,p")).map((el) => (el.textContent ?? "").trim()).filter(Boolean);
  if (items.length) return { headers: ["Value"], rows: items.map((i) => [i]), sourceName: "extracted-list" };
  return null;
}

// ---------------------------------------------------------------------------
// Quidlo: Timesheet builder & report
// ---------------------------------------------------------------------------

export type TimesheetEntry = { date: string; project: string; task: string; hours: number; billable: boolean };

export function timesheetSummary(entries: TimesheetEntry[]): {
  totalHours: number; billableHours: number; nonBillableHours: number;
  byProject: Array<{ project: string; hours: number }>;
} {
  let total = 0;
  let billable = 0;
  const projects = new Map<string, number>();
  entries.forEach((entry) => {
    total += entry.hours;
    if (entry.billable) billable += entry.hours;
    projects.set(entry.project || "Unassigned", (projects.get(entry.project || "Unassigned") ?? 0) + entry.hours);
  });
  return {
    totalHours: total,
    billableHours: billable,
    nonBillableHours: total - billable,
    byProject: Array.from(projects.entries()).map(([project, hours]) => ({ project, hours })).sort((a, b) => b.hours - a.hours),
  };
}

export function timesheetToMatrix(entries: TimesheetEntry[]): string[][] {
  return [
    ["Date", "Project", "Task", "Hours", "Billable"],
    ...entries.map((e) => [e.date, e.project, e.task, String(e.hours), e.billable ? "Yes" : "No"]),
  ];
}

// ---------------------------------------------------------------------------
// Shared export helpers
// ---------------------------------------------------------------------------

export function tableToMatrix(table: DataTable): string[][] {
  return [table.headers, ...table.rows];
}

export function downloadTableCsv(table: DataTable, fileName: string): void {
  const csv = Papa.unparse(tableToMatrix(table));
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  triggerDownload(blob, fileName);
}

export function downloadTableExcel(table: DataTable, fileName: string, sheetName: string): void {
  const workbook = buildStyledWorkbookFromMatrix(tableToMatrix(table), sheetName);
  const blob = new Blob([workbook as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  triggerDownload(blob, fileName);
}

export function downloadMatrixExcel(matrix: string[][], fileName: string, sheetName: string): void {
  const workbook = buildStyledWorkbookFromMatrix(matrix, sheetName);
  const blob = new Blob([workbook as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  triggerDownload(blob, fileName);
}

function triggerDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 250);
}
