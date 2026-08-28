import * as XLSX from "xlsx";

export interface DataTable {
  sourceName?: string;
  name: string;
  headers: string[];
  rows: Record<string, string>[];
}

export interface CompareFieldConfig {
  fieldKey: string;
  sourceColumn: number;
  targetColumn: number;
}

export interface MacroDefinition {
  id: string;
  name: string;
  category: string;
  description: string;
  fieldKeys: string[];
  keyField: string;
}

export interface ReconRecordField {
  fieldLabel: string;
  sourceValue: string;
  targetValue: string;
  status: string;
  comment: string;
}

export interface ReconRecord {
  key: string;
  fields: ReconRecordField[];
}

export interface ReconQualityIssue {
  rowRef: string;
  field: string;
  issue: string;
}

export interface ReconResult {
  summary: {
    totalRecords: number;
    matchPercentage: number;
    exceptions: number;
    duplicates: number;
    matched: number;
    partial: number;
    mismatched: number;
    missing: number;
    unmatched: number;
  };
  records: ReconRecord[];
  quality: ReconQualityIssue[];
}

export const PLATFORM_FIELDS = [
  { key: "accountNumber", label: "Account / Ref #", group: "Identifier" },
  { key: "clientName", label: "Customer / Entity Name", group: "Identity" },
  { key: "transactionDate", label: "Posting Date", group: "Temporal" },
  { key: "grossAmount", label: "Gross Value", group: "Financial" },
  { key: "netAmount", label: "Net Value", group: "Financial" },
  { key: "currency", label: "Currency ISO", group: "Financial" },
  { key: "referenceCode", label: "Audit Trace SRN", group: "Identifier" },
  { key: "taxId", label: "Tax Identification", group: "Identity" },
  { key: "status", label: "Settlement Status", group: "Operational" },
];

export const MACRO_LIBRARY: MacroDefinition[] = [
  {
    id: "bank-ledger-standard",
    name: "Standard Bank Ledger Reconciliation",
    category: "Banking",
    description: "Reconcile account numbers, posting dates, and gross balances across statement and ledger.",
    fieldKeys: ["accountNumber", "transactionDate", "grossAmount"],
    keyField: "accountNumber",
  },
  {
    id: "client-kyc-verify",
    name: "Client Account & KYC Validation",
    category: "Compliance",
    description: "Validate client legal name, tax identifiers, and currency denomination.",
    fieldKeys: ["clientName", "taxId", "currency"],
    keyField: "clientName",
  },
];

export function detectFileKind(fileName: string): "excel" | "csv" | "pdf" | "word" | "text" | "unknown" {
  const ext = fileName.split(".").pop()?.toLowerCase();
  if (ext === "xlsx" || ext === "xls" || ext === "xlsm") return "excel";
  if (ext === "csv" || ext === "tsv") return "csv";
  if (ext === "pdf") return "pdf";
  if (ext === "docx" || ext === "doc") return "word";
  if (ext === "txt") return "text";
  return "unknown";
}

export function autoDetectMapping(table: DataTable): Record<number, string> {
  const mapping: Record<number, string> = {};
  const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

  table.headers.forEach((h, idx) => {
    const norm = normalize(h);
    if (norm.includes("acc") || norm.includes("ref") || norm.includes("id")) mapping[idx] = "accountNumber";
    else if (norm.includes("name") || norm.includes("client") || norm.includes("customer")) mapping[idx] = "clientName";
    else if (norm.includes("date") || norm.includes("time")) mapping[idx] = "transactionDate";
    else if (norm.includes("gross") || norm.includes("amount") || norm.includes("bal")) mapping[idx] = "grossAmount";
    else if (norm.includes("net")) mapping[idx] = "netAmount";
    else if (norm.includes("curr") || norm.includes("iso")) mapping[idx] = "currency";
    else if (norm.includes("srn") || norm.includes("trace")) mapping[idx] = "referenceCode";
    else if (norm.includes("tax")) mapping[idx] = "taxId";
    else if (norm.includes("status")) mapping[idx] = "status";
  });

  return mapping;
}

export async function extractToTable(file: File): Promise<DataTable> {
  const kind = detectFileKind(file.name);
  if (kind === "excel") {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const firstSheet = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheet];
    const rawRows = XLSX.utils.sheet_to_json<Record<string, any>>(ws);
    const headers = rawRows.length > 0 ? Object.keys(rawRows[0]) : [];
    const rows = rawRows.map((r) => {
      const row: Record<string, string> = {};
      headers.forEach((h) => {
        row[h] = String(r[h] ?? "");
      });
      return row;
    });
    return { name: file.name, sourceName: file.name, headers, rows };
  } else {
    const text = await file.text();
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length === 0) {
      return { name: file.name, sourceName: file.name, headers: [], rows: [] };
    }
    const headers = lines[0].split(",").map((h) => h.trim().replace(/^["']|["']$/g, ""));
    const rows = lines.slice(1).map((line) => {
      const cells = line.split(",").map((c) => c.trim().replace(/^["']|["']$/g, ""));
      const row: Record<string, string> = {};
      headers.forEach((h, i) => {
        row[h] = cells[i] ?? "";
      });
      return row;
    });
    return { name: file.name, sourceName: file.name, headers, rows };
  }
}

export function runReconciliation(
  sourceTable: DataTable,
  targetTable: DataTable,
  configs: CompareFieldConfig[],
  effectiveKey: string
): ReconResult {
  const keyConfig = configs.find((c) => c.fieldKey === effectiveKey) || configs[0];
  const targetKeyHeader = targetTable.headers[keyConfig.targetColumn];
  const sourceKeyHeader = sourceTable.headers[keyConfig.sourceColumn];

  const targetMap = new Map<string, Record<string, string>>();
  targetTable.rows.forEach((r) => {
    const val = (r[targetKeyHeader] || "").trim().toLowerCase();
    if (val) targetMap.set(val, r);
  });

  let matched = 0;
  let partial = 0;
  let mismatched = 0;
  let unmatched = 0;

  const records: ReconRecord[] = [];
  const quality: ReconQualityIssue[] = [];

  sourceTable.rows.forEach((sRow, idx) => {
    const sKeyVal = (sRow[sourceKeyHeader] || "").trim();
    const match = targetMap.get(sKeyVal.toLowerCase());

    if (!match) {
      unmatched++;
      records.push({
        key: sKeyVal || `Row_${idx + 1}`,
        fields: configs.map((c) => ({
          fieldLabel: c.fieldKey,
          sourceValue: sRow[sourceTable.headers[c.sourceColumn]] || "",
          targetValue: "-",
          status: "Unmatched Source",
          comment: "Missing in target file",
        })),
      });
      return;
    }

    let isExact = true;
    const fields: ReconRecordField[] = [];

    configs.forEach((c) => {
      const sVal = (sRow[sourceTable.headers[c.sourceColumn]] || "").trim();
      const tVal = (match[targetTable.headers[c.targetColumn]] || "").trim();
      const isFieldMatch = sVal.toLowerCase() === tVal.toLowerCase();

      if (!isFieldMatch) isExact = false;

      fields.push({
        fieldLabel: c.fieldKey,
        sourceValue: sVal,
        targetValue: tVal,
        status: isFieldMatch ? "Exact Match" : "Mismatch",
        comment: isFieldMatch ? "Verified" : `Discrepancy: ${sVal} != ${tVal}`,
      });
    });

    if (isExact) {
      matched++;
    } else {
      partial++;
    }

    records.push({
      key: sKeyVal,
      fields,
    });
  });

  const total = sourceTable.rows.length;
  const matchPercentage = total > 0 ? Math.round((matched / total) * 100) : 0;

  return {
    summary: {
      totalRecords: total,
      matchPercentage,
      exceptions: partial + mismatched + unmatched,
      duplicates: 0,
      matched,
      partial,
      mismatched,
      missing: 0,
      unmatched,
    },
    records,
    quality,
  };
}

export function buildReconciliationWorkbook(
  recon: ReconResult,
  sourceName = "Source_Data",
  targetName = "Target_Data"
): Blob {
  const wb = XLSX.utils.book_new();

  const summarySheet = XLSX.utils.json_to_sheet([
    { Metric: "Source File", Value: sourceName },
    { Metric: "Target File", Value: targetName },
    { Metric: "Total Records", Value: recon.summary.totalRecords },
    { Metric: "Match Percentage", Value: `${recon.summary.matchPercentage}%` },
    { Metric: "Exact Matches", Value: recon.summary.matched },
    { Metric: "Partial Discrepancies", Value: recon.summary.partial },
    { Metric: "Exceptions", Value: recon.summary.exceptions },
  ]);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  const rows: any[] = [];
  recon.records.forEach((r) => {
    r.fields.forEach((f) => {
      rows.push({
        Record: r.key,
        Field: f.fieldLabel,
        Source: f.sourceValue,
        Target: f.targetValue,
        Status: f.status,
        Comment: f.comment,
      });
    });
  });

  const detailSheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, detailSheet, "Recon Details");

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([wbOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export function downloadBlob(blob: Blob | Uint8Array | string, fileName: string, mimeType?: string) {
  let b: Blob;
  if (typeof blob === "string") {
    b = new Blob([blob], { type: mimeType || "text/plain;charset=utf-8" });
  } else if (blob instanceof Blob) {
    b = blob;
  } else {
    b = new Blob([blob.buffer as ArrayBuffer], { type: mimeType || "application/octet-stream" });
  }
  const url = URL.createObjectURL(b);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
