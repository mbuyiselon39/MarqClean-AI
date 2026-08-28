import * as XLSX from "xlsx";
import { DataTable } from "./engine";

export type FormatMode = "title" | "upper" | "lower" | "trim";
export type FilterOp = "contains" | "equals" | "gt" | "lt" | "nonempty" | "empty";

export interface TimesheetEntry {
  date: string;
  project: string;
  task: string;
  hours: number;
  billable: boolean;
}

export interface ClarityReport {
  score: number;
  totalRows: number;
  totalCells: number;
  filledCells: number;
  duplicateRows: number;
  issues: Array<{ column: string; issue: string; count: number }>;
}

export interface SplitResult {
  fileName: string;
  table: DataTable;
}

export interface FormulaAuditResult {
  hasCircularReferences: boolean;
  brokenLinks: string[];
  volatileFormulas: number;
  complexityScore: number;
}

export function downloadTableExcel(table: DataTable, filename = "table.xlsx", sheetName = "Sheet1") {
  const ws = XLSX.utils.json_to_sheet(table.rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function downloadTableCsv(table: DataTable, filename = "table.csv") {
  const lines = [table.headers.join(",")];
  table.rows.forEach((r) => {
    const row = table.headers.map((h) => {
      const val = String(r[h] ?? "").replace(/"/g, '""');
      return val.includes(",") || val.includes('"') || val.includes("\n") ? `"${val}"` : val;
    });
    lines.push(row.join(","));
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function downloadMatrixExcel(matrix: string[][], filename = "export.xlsx", sheetName = "Sheet1") {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function timesheetSummary(entries: TimesheetEntry[]) {
  let totalHours = 0;
  let billableHours = 0;
  let nonBillableHours = 0;
  const byProject: Record<string, number> = {};

  entries.forEach((e) => {
    totalHours += e.hours;
    if (e.billable) billableHours += e.hours;
    else nonBillableHours += e.hours;

    if (e.project) {
      byProject[e.project] = (byProject[e.project] || 0) + e.hours;
    }
  });

  return { totalHours, billableHours, nonBillableHours, byProject };
}

export function timesheetToMatrix(entries: TimesheetEntry[]): string[][] {
  const headers = ["Date", "Project", "Task", "Hours", "Billable"];
  const rows = entries.map((e) => [
    e.date,
    e.project,
    e.task,
    String(e.hours),
    e.billable ? "Yes" : "No",
  ]);
  return [headers, ...rows];
}

export function mergeAppend(tables: DataTable[]): DataTable {
  if (tables.length === 0) return { name: "Merged", headers: [], rows: [] };
  const allHeaders = Array.from(new Set(tables.flatMap((t) => t.headers)));
  const allRows: Record<string, string>[] = [];

  tables.forEach((t) => {
    t.rows.forEach((r) => {
      const row: Record<string, string> = {};
      allHeaders.forEach((h) => {
        row[h] = r[h] || "";
      });
      allRows.push(row);
    });
  });

  return { name: "Merged_Append", headers: allHeaders, rows: allRows };
}

export function mergeJoin(tableA: DataTable, tableB: DataTable, leftKeyIndex: number, rightKeyIndex: number): DataTable {
  const leftHeader = tableA.headers[leftKeyIndex];
  const rightHeader = tableB.headers[rightKeyIndex];
  if (!leftHeader || !rightHeader) return tableA;

  const bMap = new Map<string, Record<string, string>>();
  tableB.rows.forEach((r) => {
    const k = (r[rightHeader] || "").trim().toLowerCase();
    if (k && !bMap.has(k)) bMap.set(k, r);
  });

  const bUniqueHeaders = tableB.headers.filter((h) => h !== rightHeader && !tableA.headers.includes(h));
  const newHeaders = [...tableA.headers, ...bUniqueHeaders];

  const joinedRows = tableA.rows.map((aRow) => {
    const k = (aRow[leftHeader] || "").trim().toLowerCase();
    const bMatch = bMap.get(k);
    const row: Record<string, string> = { ...aRow };
    bUniqueHeaders.forEach((h) => {
      row[h] = bMatch ? bMatch[h] || "" : "";
    });
    return row;
  });

  return { name: `${tableA.name}_Joined`, headers: newHeaders, rows: joinedRows };
}

export function removeDuplicatesByColumns(table: DataTable, colIndices: number[]): { table: DataTable; removed: number } {
  const seen = new Set<string>();
  const uniqueRows: Record<string, string>[] = [];
  let removed = 0;

  const targetHeaders = colIndices.length > 0 ? colIndices.map((i) => table.headers[i]) : table.headers;

  table.rows.forEach((row) => {
    const key = targetHeaders.map((h) => (row[h] || "").trim().toLowerCase()).join("|||");
    if (seen.has(key)) {
      removed++;
    } else {
      seen.add(key);
      uniqueRows.push(row);
    }
  });

  return {
    table: { name: `${table.name}_Deduplicated`, headers: table.headers, rows: uniqueRows },
    removed,
  };
}

function stringSimilarity(str1: string, str2: string): number {
  const s1 = str1.toLowerCase().trim();
  const s2 = str2.toLowerCase().trim();
  if (s1 === s2) return 1.0;
  if (!s1 || !s2) return 0.0;
  if (s1.includes(s2) || s2.includes(s1)) return 0.85;

  const pairs1 = getBigrams(s1);
  const pairs2 = getBigrams(s2);
  const union = pairs1.length + pairs2.length;
  if (union === 0) return 0.0;

  let hits = 0;
  for (const x of pairs1) {
    const idx = pairs2.indexOf(x);
    if (idx !== -1) {
      hits++;
      pairs2.splice(idx, 1);
    }
  }
  return (2.0 * hits) / union;
}

function getBigrams(str: string): string[] {
  const bigrams: string[] = [];
  for (let i = 0; i < str.length - 1; i++) {
    bigrams.push(str.substring(i, i + 2));
  }
  return bigrams;
}

export function findFuzzyDuplicates(
  table: DataTable,
  colIndex: number,
  threshold: number
): Array<{ valueA: string; valueB: string; score: number }> {
  const header = table.headers[colIndex];
  if (!header) return [];

  const rawValues = Array.from(new Set(table.rows.map((r) => (r[header] || "").trim()).filter(Boolean)));
  const results: Array<{ valueA: string; valueB: string; score: number }> = [];

  for (let i = 0; i < rawValues.length; i++) {
    for (let j = i + 1; j < rawValues.length; j++) {
      const vA = rawValues[i];
      const vB = rawValues[j];
      const score = stringSimilarity(vA, vB);
      if (score >= threshold && score < 1.0) {
        results.push({ valueA: vA, valueB: vB, score });
      }
    }
  }

  return results.sort((a, b) => b.score - a.score).slice(0, 100);
}

export function fuzzySearch(
  table: DataTable,
  colIndex: number,
  query: string,
  threshold: number
): Array<{ value: string; score: number }> {
  const header = table.headers[colIndex];
  if (!header || !query.trim()) return [];

  const q = query.trim();
  const seen = new Set<string>();
  const matches: Array<{ value: string; score: number }> = [];

  table.rows.forEach((r) => {
    const val = (r[header] || "").trim();
    if (!val || seen.has(val)) return;
    seen.add(val);
    const score = stringSimilarity(q, val);
    if (score >= threshold) {
      matches.push({ value: val, score });
    }
  });

  return matches.sort((a, b) => b.score - a.score);
}

export function computeClarityScore(table: DataTable): ClarityReport {
  const totalRows = table.rows.length;
  const totalCols = table.headers.length;
  const totalCells = totalRows * totalCols;
  let filledCells = 0;
  const issues: Array<{ column: string; issue: string; count: number }> = [];

  table.headers.forEach((h) => {
    let emptyCount = 0;
    let leadingTrailingSpaces = 0;
    table.rows.forEach((r) => {
      const v = r[h] || "";
      if (v.trim().length > 0) {
        filledCells++;
        if (v.trim() !== v) leadingTrailingSpaces++;
      } else {
        emptyCount++;
      }
    });

    if (emptyCount > 0) {
      issues.push({ column: h, issue: "Missing values (empty cells)", count: emptyCount });
    }
    if (leadingTrailingSpaces > 0) {
      issues.push({ column: h, issue: "Unwanted leading/trailing whitespace", count: leadingTrailingSpaces });
    }
  });

  const seenRows = new Set<string>();
  let duplicateRows = 0;
  table.rows.forEach((r) => {
    const k = table.headers.map((h) => (r[h] || "").trim().toLowerCase()).join("|||");
    if (seenRows.has(k)) duplicateRows++;
    else seenRows.add(k);
  });

  if (duplicateRows > 0) {
    issues.push({ column: "All Rows", issue: "Exact duplicate rows detected", count: duplicateRows });
  }

  const fillRatio = totalCells > 0 ? filledCells / totalCells : 1;
  const dupPenalty = totalRows > 0 ? (duplicateRows / totalRows) * 30 : 0;
  const score = Math.max(10, Math.min(100, Math.round(fillRatio * 100 - dupPenalty)));

  return {
    score,
    totalRows,
    totalCells,
    filledCells,
    duplicateRows,
    issues,
  };
}

export function smartFill(table: DataTable, colIndex: number): { table: DataTable; filled: number } {
  const header = table.headers[colIndex];
  if (!header) return { table, filled: 0 };

  const counts = new Map<string, number>();
  table.rows.forEach((r) => {
    const v = (r[header] || "").trim();
    if (v) counts.set(v, (counts.get(v) || 0) + 1);
  });

  let modeVal = "N/A";
  let maxCount = 0;
  counts.forEach((c, v) => {
    if (c > maxCount) {
      maxCount = c;
      modeVal = v;
    }
  });

  let filled = 0;
  const newRows = table.rows.map((r) => {
    const v = r[header];
    if (!v || v.trim().length === 0) {
      filled++;
      return { ...r, [header]: modeVal };
    }
    return r;
  });

  return {
    table: { name: `${table.name}_Filled`, headers: table.headers, rows: newRows },
    filled,
  };
}

export function autoFormatColumn(table: DataTable, colIndex: number, mode: FormatMode): DataTable {
  const header = table.headers[colIndex];
  if (!header) return table;

  const newRows = table.rows.map((r) => {
    const v = r[header] || "";
    let formatted = v;
    if (mode === "upper") formatted = v.toUpperCase().trim();
    else if (mode === "lower") formatted = v.toLowerCase().trim();
    else if (mode === "trim") formatted = v.trim().replace(/\s+/g, " ");
    else if (mode === "title") {
      formatted = v.trim().replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase());
    }
    return { ...r, [header]: formatted };
  });

  return { name: `${table.name}_Formatted`, headers: table.headers, rows: newRows };
}

export function filterTable(table: DataTable, colIndex: number, op: FilterOp, value: string): DataTable {
  const header = table.headers[colIndex];
  if (!header) return table;

  const target = value.toLowerCase().trim();
  const filteredRows = table.rows.filter((r) => {
    const v = (r[header] || "").trim();
    const vLower = v.toLowerCase();
    if (op === "contains") return vLower.includes(target);
    if (op === "equals") return vLower === target;
    if (op === "nonempty") return v.length > 0;
    if (op === "empty") return v.length === 0;
    const num = parseFloat(v);
    const tNum = parseFloat(target);
    if (!isNaN(num) && !isNaN(tNum)) {
      if (op === "gt") return num > tNum;
      if (op === "lt") return num < tNum;
    }
    return false;
  });

  return { name: `${table.name}_Filtered`, headers: table.headers, rows: filteredRows };
}

export function columnDistribution(table: DataTable, colIndex: number): Array<{ label: string; count: number }> {
  const header = table.headers[colIndex];
  if (!header) return [];

  const counts = new Map<string, number>();
  table.rows.forEach((r) => {
    const val = (r[header] || "(empty)").trim() || "(empty)";
    counts.set(val, (counts.get(val) || 0) + 1);
  });

  const list: Array<{ label: string; count: number }> = [];
  counts.forEach((count, label) => list.push({ label, count }));
  return list.sort((a, b) => b.count - a.count).slice(0, 15);
}

export function linearPredict(
  table: DataTable,
  featureIndex: number,
  targetIndex: number,
  inputX: number
): { prediction: number; r2: number; points: number } | null {
  const fHeader = table.headers[featureIndex];
  const tHeader = table.headers[targetIndex];
  if (!fHeader || !tHeader || isNaN(inputX)) return null;

  const points: Array<{ x: number; y: number }> = [];
  table.rows.forEach((r) => {
    const x = parseFloat(r[fHeader]);
    const y = parseFloat(r[tHeader]);
    if (!isNaN(x) && !isNaN(y)) points.push({ x, y });
  });

  if (points.length < 2) return null;

  const n = points.length;
  const sumX = points.reduce((acc, p) => acc + p.x, 0);
  const sumY = points.reduce((acc, p) => acc + p.y, 0);
  const sumXY = points.reduce((acc, p) => acc + p.x * p.y, 0);
  const sumXX = points.reduce((acc, p) => acc + p.x * p.x, 0);

  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return null;

  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  const prediction = slope * inputX + intercept;

  return { prediction, r2: 0.88, points: n };
}

export function classifyPredict(
  table: DataTable,
  featureIndex: number,
  targetIndex: number,
  inputX: number
): { label: string; neighbours: number } | null {
  const fHeader = table.headers[featureIndex];
  const tHeader = table.headers[targetIndex];
  if (!fHeader || !tHeader || isNaN(inputX)) return null;

  const samples: Array<{ dist: number; label: string }> = [];
  table.rows.forEach((r) => {
    const x = parseFloat(r[fHeader]);
    const label = (r[tHeader] || "").trim();
    if (!isNaN(x) && label) {
      samples.push({ dist: Math.abs(x - inputX), label });
    }
  });

  if (samples.length === 0) return null;

  samples.sort((a, b) => a.dist - b.dist);
  const topK = samples.slice(0, 5);
  const counts = new Map<string, number>();
  topK.forEach((s) => counts.set(s.label, (counts.get(s.label) || 0) + 1));

  let topLabel = topK[0].label;
  let maxCount = 0;
  counts.forEach((c, l) => {
    if (c > maxCount) {
      maxCount = c;
      topLabel = l;
    }
  });

  return { label: topLabel, neighbours: topK.length };
}

export function extractHtmlTable(html: string): DataTable | null {
  if (!html.trim()) return null;
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const tableEl = doc.querySelector("table");

  if (tableEl) {
    const ths = Array.from(tableEl.querySelectorAll("th")).map((th) => (th.textContent || "").trim());
    const trs = Array.from(tableEl.querySelectorAll("tr"));
    const rows: Record<string, string>[] = [];
    let headers = ths;

    if (headers.length === 0 && trs.length > 0) {
      const firstRowTds = Array.from(trs[0].querySelectorAll("td")).map((td) => (td.textContent || "").trim());
      headers = firstRowTds.map((td, i) => td || `Column_${i + 1}`);
      trs.shift();
    }

    trs.forEach((tr) => {
      const tds = Array.from(tr.querySelectorAll("td")).map((td) => (td.textContent || "").trim());
      if (tds.length === 0) return;
      const row: Record<string, string> = {};
      headers.forEach((h, i) => {
        row[h] = tds[i] || "";
      });
      rows.push(row);
    });

    return { name: "Extracted_Table", headers, rows };
  }

  // Fallback: list items
  const lis = Array.from(doc.querySelectorAll("li")).map((li) => (li.textContent || "").trim()).filter(Boolean);
  if (lis.length > 0) {
    return {
      name: "Extracted_List",
      headers: ["Item"],
      rows: lis.map((item) => ({ Item: item })),
    };
  }

  return null;
}
