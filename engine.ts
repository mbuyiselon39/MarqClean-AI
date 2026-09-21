import Papa from "papaparse";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

// ---------------------------------------------------------------------------
// Shared low level helpers (self contained so the engine has no App.tsx deps)
// ---------------------------------------------------------------------------

export type DataTable = {
  headers: string[];
  rows: string[][];
  sourceName: string;
};

export function sanitize(value: unknown): string {
  return String(value ?? "")
    .replace(/\uFEFF/g, "")
    .replace(/[\t\n\r]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function columnIndexToReference(index: number): string {
  let dividend = index + 1;
  let reference = "";
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    reference = String.fromCharCode(65 + modulo) + reference;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return reference;
}

function columnReferenceToIndex(reference: string): number {
  const letters = reference.replace(/[^A-Z]/gi, "").toUpperCase();
  return letters.split("").reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function downloadBlob(content: string | Uint8Array, fileName: string, type: string): void {
  const blob = new Blob([content as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 250);
}

// ---------------------------------------------------------------------------
// File extraction into a common structured dataset
// ---------------------------------------------------------------------------

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map((item) =>
    Array.from(item.getElementsByTagName("t"))
      .map((node) => node.textContent ?? "")
      .join("")
  );
}

function readCellValue(cell: Element, shared: string[]): string {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") {
    return Array.from(cell.getElementsByTagName("t")).map((n) => n.textContent ?? "").join("");
  }
  const raw = cell.getElementsByTagName("v")[0]?.textContent ?? "";
  if (type === "s") return shared[Number(raw)] ?? "";
  if (type === "b") return raw === "1" ? "TRUE" : "FALSE";
  return raw;
}

export async function readWorkbookMatrix(file: File): Promise<string[][]> {
  const buffer = await file.arrayBuffer();
  const files = unzipSync(new Uint8Array(buffer));
  const sheetPath = files["xl/worksheets/sheet1.xml"]
    ? "xl/worksheets/sheet1.xml"
    : Object.keys(files).find((path) => path.startsWith("xl/worksheets/sheet"));

  if (!sheetPath) throw new Error("The Excel workbook does not contain a readable worksheet.");

  const shared = parseSharedStrings(files["xl/sharedStrings.xml"] ? strFromU8(files["xl/sharedStrings.xml"]) : "");
  const doc = new DOMParser().parseFromString(strFromU8(files[sheetPath]), "application/xml");

  return Array.from(doc.getElementsByTagName("row")).map((rowEl) => {
    const values: string[] = [];
    Array.from(rowEl.getElementsByTagName("c")).forEach((cell) => {
      const ref = cell.getAttribute("r") ?? "A1";
      values[columnReferenceToIndex(ref)] = readCellValue(cell, shared);
    });
    return values;
  });
}

export { extractPdfTextInWorker as extractPdfText } from "./pdfWorkerClient";


export function matrixToTable(matrix: string[][], sourceName: string): DataTable {
  const usable = matrix.filter((row) => row.some((v) => sanitize(v)));
  if (!usable.length) return { headers: [], rows: [], sourceName };

  const columnCount = usable.reduce((max, row) => Math.max(max, row.length), 0);
  const headers = Array.from({ length: columnCount }, (_, i) => sanitize(usable[0][i]) || `Column ${i + 1}`);
  const rows = usable.slice(1).map((row) => Array.from({ length: columnCount }, (_, i) => sanitize(row[i] ?? "")));

  return { headers, rows, sourceName };
}

export type FileKind = "excel" | "csv" | "pdf" | "word" | "text" | "unknown";

export function detectFileKind(fileName: string): FileKind {
  const ext = fileName.split(".").pop()?.toLowerCase();
  if (ext === "xlsx" || ext === "xls") return "excel";
  if (ext === "csv") return "csv";
  if (ext === "pdf") return "pdf";
  if (ext === "docx" || ext === "doc") return "word";
  if (ext === "txt") return "text";
  return "unknown";
}

async function extractWordText(file: File): Promise<string> {
  // .docx is a zip containing word/document.xml
  try {
    const buffer = await file.arrayBuffer();
    const files = unzipSync(new Uint8Array(buffer));
    const docXml = files["word/document.xml"] ? strFromU8(files["word/document.xml"]) : "";
    if (!docXml) return "";
    const doc = new DOMParser().parseFromString(docXml, "application/xml");
    const paragraphs = Array.from(doc.getElementsByTagName("w:p")).map((p) =>
      Array.from(p.getElementsByTagName("w:t")).map((t) => t.textContent ?? "").join("")
    );
    return paragraphs.filter((line) => line.trim()).join("\n");
  } catch {
    return "";
  }
}

export async function extractToTable(file: File): Promise<DataTable> {
  const kind = detectFileKind(file.name);

  if (kind === "excel") {
    return matrixToTable(await readWorkbookMatrix(file), file.name);
  }
  if (kind === "csv" || kind === "text") {
    return matrixToTable(textToMatrix(await file.text()), file.name);
  }
  if (kind === "pdf") {
    return matrixToTable(textToMatrix(await extractPdfText(file)), file.name);
  }
  if (kind === "word") {
    const text = await extractWordText(file);
    if (!text) throw new Error("Could not read the Word document. Save it as .docx and try again.");
    return matrixToTable(textToMatrix(text), file.name);
  }

  throw new Error("Unsupported file type. Upload PDF, Excel, CSV, or Word.");
}

// ---------------------------------------------------------------------------
// Platform fields, auto detection and mapping
// ---------------------------------------------------------------------------

export type PlatformField = {
  key: string;
  label: string;
  group: "Identity" | "Address" | "Contact" | "Compliance" | "Banking" | "Account";
  aliases: string[];
  kind: "name" | "id" | "address" | "phone" | "email" | "postal" | "generic";
};

export const PLATFORM_FIELDS: PlatformField[] = [
  { key: "clientName", label: "Client Name", group: "Identity", kind: "name", aliases: ["client name", "name", "full name", "account holder", "holder", "investor name", "shareholder name", "member name"] },
  { key: "firstName", label: "First Name", group: "Identity", kind: "name", aliases: ["first name", "firstname", "given name", "forename"] },
  { key: "middleName", label: "Middle Name", group: "Identity", kind: "name", aliases: ["middle name", "middlename", "second name"] },
  { key: "surname", label: "Surname", group: "Identity", kind: "name", aliases: ["surname", "last name", "lastname", "family name"] },
  { key: "entityName", label: "Entity Name", group: "Identity", kind: "name", aliases: ["entity name", "entity", "legal entity"] },
  { key: "trustName", label: "Trust Name", group: "Identity", kind: "name", aliases: ["trust name", "trust"] },
  { key: "companyName", label: "Company Name", group: "Identity", kind: "name", aliases: ["company name", "company", "organisation", "organization", "business name"] },
  { key: "addressLine1", label: "Address Line 1", group: "Address", kind: "address", aliases: ["address", "address line 1", "address1", "street", "residential address", "physical address", "postal address"] },
  { key: "addressLine2", label: "Address Line 2", group: "Address", kind: "address", aliases: ["address line 2", "address2", "line 2"] },
  { key: "suburb", label: "Suburb", group: "Address", kind: "address", aliases: ["suburb", "area"] },
  { key: "city", label: "City", group: "Address", kind: "address", aliases: ["city", "town"] },
  { key: "province", label: "Province", group: "Address", kind: "address", aliases: ["province", "state", "region"] },
  { key: "postalCode", label: "Postal Code", group: "Address", kind: "postal", aliases: ["postal code", "postcode", "zip", "zip code", "post code"] },
  { key: "country", label: "Country", group: "Address", kind: "generic", aliases: ["country", "nation"] },
  { key: "mobile", label: "Mobile Number", group: "Contact", kind: "phone", aliases: ["mobile", "cell", "cellphone", "mobile number", "cell number"] },
  { key: "telephone", label: "Telephone Number", group: "Contact", kind: "phone", aliases: ["telephone", "phone", "tel", "landline", "phone number", "contact number"] },
  { key: "email", label: "Email Address", group: "Contact", kind: "email", aliases: ["email", "e-mail", "email address", "mail"] },
  { key: "saId", label: "South African ID", group: "Compliance", kind: "id", aliases: ["id number", "sa id", "south african id", "identity number", "id"] },
  { key: "passport", label: "Passport Number", group: "Compliance", kind: "id", aliases: ["passport", "passport number", "passport no"] },
  { key: "companyReg", label: "Company Registration Number", group: "Compliance", kind: "id", aliases: ["company registration", "registration number", "reg number", "company reg", "cipc"] },
  { key: "trustReg", label: "Trust Registration Number", group: "Compliance", kind: "id", aliases: ["trust registration", "trust reg", "trust number", "it number"] },
  { key: "taxNumber", label: "Tax Number", group: "Compliance", kind: "id", aliases: ["tax number", "tax", "tax reference", "income tax number", "vat number"] },
  { key: "crsCountry", label: "CRS Country", group: "Compliance", kind: "generic", aliases: ["crs country", "crs"] },
  { key: "fatcaCountry", label: "FATCA Country", group: "Compliance", kind: "generic", aliases: ["fatca country", "fatca"] },
  { key: "bankName", label: "Bank Name", group: "Banking", kind: "generic", aliases: ["bank name", "bank"] },
  { key: "branchCode", label: "Branch Code", group: "Banking", kind: "id", aliases: ["branch code", "branch", "sort code"] },
  { key: "accountNumber", label: "Account Number", group: "Banking", kind: "id", aliases: ["account number", "account no", "acc number", "bank account"] },
  { key: "accountType", label: "Account Type", group: "Banking", kind: "generic", aliases: ["account type", "acc type"] },
  { key: "accountHolder", label: "Account Holder Name", group: "Banking", kind: "name", aliases: ["account holder", "account holder name", "holder name"] },
  { key: "brokerNumber", label: "Broker Number", group: "Account", kind: "id", aliases: ["broker number", "broker no", "broker"] },
  { key: "investorNumber", label: "Investor Number", group: "Account", kind: "id", aliases: ["investor number", "investor no", "investor id"] },
  { key: "portfolioNumber", label: "Portfolio Number", group: "Account", kind: "id", aliases: ["portfolio number", "portfolio no", "portfolio"] },
  { key: "shareholderNumber", label: "Shareholder Number", group: "Account", kind: "id", aliases: ["shareholder number", "shareholder no", "shareholder id"] },
  { key: "clientReference", label: "Client Reference", group: "Account", kind: "generic", aliases: ["client reference", "client ref", "reference", "ref"] },
];

export function getPlatformField(key: string): PlatformField | undefined {
  return PLATFORM_FIELDS.find((field) => field.key === key);
}

function normalizeHeaderText(header: string): string {
  return header.toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

function scoreHeaderMatch(header: string, field: PlatformField): number {
  const cleaned = normalizeHeaderText(header);
  if (!cleaned) return 0;
  let best = 0;
  field.aliases.forEach((alias) => {
    if (cleaned === alias) best = Math.max(best, 100);
    else if (cleaned.includes(alias) || alias.includes(cleaned)) best = Math.max(best, 70);
  });
  return best;
}

// Auto detect: map each column index -> platform field key (or "")
export function autoDetectMapping(table: DataTable): Record<number, string> {
  const mapping: Record<number, string> = {};
  const used = new Set<string>();

  table.headers.forEach((header, index) => {
    let bestKey = "";
    let bestScore = 40; // threshold
    PLATFORM_FIELDS.forEach((field) => {
      if (used.has(field.key)) return;
      const score = scoreHeaderMatch(header, field);
      if (score > bestScore) {
        bestScore = score;
        bestKey = field.key;
      }
    });
    if (bestKey) {
      mapping[index] = bestKey;
      used.add(bestKey);
    } else {
      mapping[index] = "";
    }
  });

  return mapping;
}

// ---------------------------------------------------------------------------
// Standardisation and matching logic (micro-service level)
// ---------------------------------------------------------------------------

const ADDRESS_ABBREVIATIONS: Record<string, string> = {
  st: "street", str: "street", rd: "road", ave: "avenue", av: "avenue",
  dr: "drive", ln: "lane", blvd: "boulevard", apt: "apartment", ste: "suite",
  "p o box": "po box", "p.o. box": "po box", "p.o box": "po box",
};

export function standardizeName(value: string): string {
  return sanitize(value)
    .toLowerCase()
    .replace(/\b(mr|mrs|ms|miss|dr|prof|adv)\b\.?/g, "")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function standardizeAddress(value: string): string {
  let text = sanitize(value).toLowerCase().replace(/\./g, " ").replace(/,/g, " ").replace(/\s+/g, " ").trim();
  text = text.replace(/\bp\s*o\s*box\b/g, "po box");
  Object.entries(ADDRESS_ABBREVIATIONS).forEach(([abbr, full]) => {
    text = text.replace(new RegExp(`\\b${abbr.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "g"), full);
  });
  return text.replace(/\s+/g, " ").trim();
}

export function standardizePhone(value: string): string {
  const digits = sanitize(value).replace(/\D/g, "");
  return digits.replace(/^0027/, "27").replace(/^27/, "").replace(/^0/, "");
}

export function standardizeEmail(value: string): string {
  return sanitize(value).toLowerCase();
}

export function standardizePostal(value: string): string {
  return sanitize(value).replace(/\D/g, "");
}

export function standardizeId(value: string): string {
  return sanitize(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function standardizeGeneric(value: string): string {
  return sanitize(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function standardizeByKind(value: string, kind: PlatformField["kind"]): string {
  switch (kind) {
    case "name": return standardizeName(value);
    case "address": return standardizeAddress(value);
    case "phone": return standardizePhone(value);
    case "email": return standardizeEmail(value);
    case "postal": return standardizePostal(value);
    case "id": return standardizeId(value);
    default: return standardizeGeneric(value);
  }
}

// Levenshtein based similarity (0..1)
export function similarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const matrix: number[][] = Array.from({ length: a.length + 1 }, () => new Array(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j += 1) matrix[0][j] = j;
  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(matrix[i - 1][j] + 1, matrix[i][j - 1] + 1, matrix[i - 1][j - 1] + cost);
    }
  }
  const distance = matrix[a.length][b.length];
  return 1 - distance / Math.max(a.length, b.length);
}

export type FieldStatus = "Exact Match" | "Partial Match" | "Mismatch" | "Missing Value";

export function compareValues(sourceRaw: string, targetRaw: string, kind: PlatformField["kind"]): { status: FieldStatus; score: number } {
  const sourceEmpty = !sanitize(sourceRaw);
  const targetEmpty = !sanitize(targetRaw);

  if (sourceEmpty && targetEmpty) return { status: "Missing Value", score: 0 };
  if (sourceEmpty || targetEmpty) return { status: "Missing Value", score: 0 };

  const s = standardizeByKind(sourceRaw, kind);
  const t = standardizeByKind(targetRaw, kind);

  if (s === t) return { status: "Exact Match", score: 1 };

  const score = similarity(s, t);
  if (score >= 0.85) return { status: "Partial Match", score };
  return { status: "Mismatch", score };
}

// ---------------------------------------------------------------------------
// Reconciliation
// ---------------------------------------------------------------------------

export type CompareFieldConfig = {
  fieldKey: string;
  sourceColumn: number;
  targetColumn: number;
};

export type ReconRecordField = {
  fieldKey: string;
  fieldLabel: string;
  sourceValue: string;
  targetValue: string;
  status: FieldStatus;
  comment: string;
};

export type ReconRecord = {
  key: string;
  status: FieldStatus | "Duplicate" | "Unmatched Source" | "Unmatched Target";
  fields: ReconRecordField[];
};

export type ReconSummary = {
  totalRecords: number;
  matched: number;
  partial: number;
  mismatched: number;
  missing: number;
  duplicates: number;
  unmatched: number;
  exceptions: number;
  matchPercentage: number;
};

export type DataQualityIssue = {
  source: string;
  rowRef: string;
  field: string;
  issue: string;
};

export type ReconResult = {
  records: ReconRecord[];
  summary: ReconSummary;
  quality: DataQualityIssue[];
  keyFieldLabel: string;
};

function buildRowKey(row: string[], keyColumn: number, kind: PlatformField["kind"]): string {
  return standardizeByKind(row[keyColumn] ?? "", kind);
}

export function runReconciliation(
  source: DataTable,
  target: DataTable,
  fields: CompareFieldConfig[],
  keyFieldKey: string
): ReconResult {
  const keyConfig = fields.find((f) => f.fieldKey === keyFieldKey) ?? fields[0];
  const keyField = getPlatformField(keyConfig.fieldKey);
  const keyKind = keyField?.kind ?? "generic";

  // Index target rows by key
  const targetIndex = new Map<string, number[]>();
  target.rows.forEach((row, index) => {
    const key = buildRowKey(row, keyConfig.targetColumn, keyKind);
    if (!key) return;
    const bucket = targetIndex.get(key) ?? [];
    bucket.push(index);
    targetIndex.set(key, bucket);
  });

  // Track source duplicate keys
  const sourceKeySeen = new Map<string, number>();
  const records: ReconRecord[] = [];
  const quality: DataQualityIssue[] = [];
  const matchedTargetRows = new Set<number>();

  const summary: ReconSummary = {
    totalRecords: 0, matched: 0, partial: 0, mismatched: 0, missing: 0,
    duplicates: 0, unmatched: 0, exceptions: 0, matchPercentage: 0,
  };

  source.rows.forEach((sourceRow, sourceRowIndex) => {
    const key = buildRowKey(sourceRow, keyConfig.sourceColumn, keyKind);
    const displayKey = sanitize(sourceRow[keyConfig.sourceColumn] ?? "") || `Row ${sourceRowIndex + 2}`;

    // Duplicate detection in source
    if (key) {
      const seen = sourceKeySeen.get(key) ?? 0;
      sourceKeySeen.set(key, seen + 1);
      if (seen >= 1) {
        summary.duplicates += 1;
        quality.push({ source: source.sourceName, rowRef: `Row ${sourceRowIndex + 2}`, field: keyField?.label ?? "Key", issue: "Duplicate key in source file" });
      }
    }

    const targetMatches = key ? targetIndex.get(key) ?? [] : [];
    const targetRowIndex = targetMatches.find((i) => !matchedTargetRows.has(i)) ?? targetMatches[0];
    const hasTarget = targetRowIndex !== undefined;
    if (hasTarget) matchedTargetRows.add(targetRowIndex);

    const recordFields: ReconRecordField[] = [];
    let worst: FieldStatus = "Exact Match";

    fields.forEach((config) => {
      const field = getPlatformField(config.fieldKey);
      const kind = field?.kind ?? "generic";
      const sourceValue = sanitize(sourceRow[config.sourceColumn] ?? "");
      const targetValue = hasTarget ? sanitize(target.rows[targetRowIndex][config.targetColumn] ?? "") : "";

      // Data quality checks on source
      if (!sourceValue) {
        quality.push({ source: source.sourceName, rowRef: `Row ${sourceRowIndex + 2}`, field: field?.label ?? config.fieldKey, issue: "Missing value" });
      } else if (kind === "email" && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(sourceValue)) {
        quality.push({ source: source.sourceName, rowRef: `Row ${sourceRowIndex + 2}`, field: field?.label ?? config.fieldKey, issue: "Invalid email format" });
      } else if (kind === "postal" && standardizePostal(sourceValue).length < 4) {
        quality.push({ source: source.sourceName, rowRef: `Row ${sourceRowIndex + 2}`, field: field?.label ?? config.fieldKey, issue: "Invalid postal code" });
      } else if (kind === "id" && config.fieldKey === "saId" && standardizeId(sourceValue).length !== 13 && /^\d+$/.test(standardizeId(sourceValue))) {
        quality.push({ source: source.sourceName, rowRef: `Row ${sourceRowIndex + 2}`, field: field?.label ?? config.fieldKey, issue: "SA ID should be 13 digits" });
      }

      const comparison = hasTarget
        ? compareValues(sourceValue, targetValue, kind)
        : { status: "Missing Value" as FieldStatus, score: 0 };

      let comment = "";
      if (comparison.status === "Partial Match") comment = `Similarity ${(comparison.score * 100).toFixed(0)}% - review`;
      if (comparison.status === "Mismatch") comment = "Values differ - investigate";
      if (comparison.status === "Missing Value") comment = hasTarget ? "Value missing in one file" : "No matching target record";

      recordFields.push({
        fieldKey: config.fieldKey,
        fieldLabel: field?.label ?? config.fieldKey,
        sourceValue,
        targetValue,
        status: comparison.status,
        comment,
      });

      // Track worst status
      const rank: Record<FieldStatus, number> = { "Exact Match": 0, "Partial Match": 1, "Missing Value": 2, "Mismatch": 3 };
      if (rank[comparison.status] > rank[worst]) worst = comparison.status;
    });

    const recordStatus: ReconRecord["status"] = hasTarget ? worst : "Unmatched Source";

    records.push({ key: displayKey, status: recordStatus, fields: recordFields });

    summary.totalRecords += 1;
    const statusValue = recordStatus as string;
    if (statusValue === "Exact Match") summary.matched += 1;
    else if (statusValue === "Partial Match") summary.partial += 1;
    else if (statusValue === "Mismatch") summary.mismatched += 1;
    else if (statusValue === "Missing Value") summary.missing += 1;
    else summary.unmatched += 1;
  });

  // Unmatched target rows
  target.rows.forEach((targetRow, index) => {
    if (matchedTargetRows.has(index)) return;
    const key = buildRowKey(targetRow, keyConfig.targetColumn, keyKind);
    if (!key) return;
    const displayKey = sanitize(targetRow[keyConfig.targetColumn] ?? "") || `Row ${index + 2}`;
    const recordFields: ReconRecordField[] = fields.map((config) => {
      const field = getPlatformField(config.fieldKey);
      return {
        fieldKey: config.fieldKey,
        fieldLabel: field?.label ?? config.fieldKey,
        sourceValue: "",
        targetValue: sanitize(targetRow[config.targetColumn] ?? ""),
        status: "Missing Value" as FieldStatus,
        comment: "Record only exists in target file",
      };
    });
    records.push({ key: displayKey, status: "Unmatched Target", fields: recordFields });
    summary.totalRecords += 1;
    summary.unmatched += 1;
  });

  summary.exceptions = summary.mismatched + summary.missing + summary.unmatched + summary.duplicates;
  const denominator = summary.totalRecords || 1;
  summary.matchPercentage = Math.round((summary.matched / denominator) * 1000) / 10;

  return { records, summary, quality, keyFieldLabel: keyField?.label ?? "Key" };
}

// ---------------------------------------------------------------------------
// Colour coded multi sheet Excel report
// ---------------------------------------------------------------------------

const STATUS_STYLE: Record<string, number> = {
  "Exact Match": 3,      // green
  "Partial Match": 4,    // blue
  "Mismatch": 5,         // light red
  "Missing Value": 6,    // orange
  "Duplicate": 6,
  "Unmatched Source": 6,
  "Unmatched Target": 6,
};

const REPORT_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/><color rgb="FF0F172A"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font><font><b/><sz val="14"/><name val="Calibri"/><color rgb="FF0F172A"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCDEFD6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCFE3FB"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF9C9CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE2BD"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

type SheetCell = { text?: string; number?: number; style: number };

function cellXml(ref: string, cell: SheetCell): string {
  if (cell.number !== undefined) return `<c r="${ref}" s="${cell.style}"><v>${cell.number}</v></c>`;
  const text = cell.text ?? "";
  if (!text) return `<c r="${ref}" s="${cell.style}"/>`;
  return `<c r="${ref}" t="inlineStr" s="${cell.style}"><is><t>${escapeXml(text)}</t></is></c>`;
}

function sheetFromRows(rows: SheetCell[][]): string {
  const columnCount = rows.reduce((max, r) => Math.max(max, r.length), 1);
  const rowXml = rows
    .map((row, rowIndex) => {
      const cells = row.map((cell, colIndex) => cellXml(`${columnIndexToReference(colIndex)}${rowIndex + 1}`, cell)).join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");
  const cols = Array.from({ length: columnCount }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="26" customWidth="1"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${cols}</cols><sheetData>${rowXml}</sheetData></worksheet>`;
}

export function buildReconciliationWorkbook(result: ReconResult, sourceName: string, targetName: string): Uint8Array {
  const { summary, records, quality } = result;

  // Sheet 1: Summary
  const summaryRows: SheetCell[][] = [
    [{ text: "Reconciliation Summary", style: 2 }],
    [{ text: `Source: ${sourceName}`, style: 0 }],
    [{ text: `Target: ${targetName}`, style: 0 }],
    [],
    [{ text: "Metric", style: 1 }, { text: "Value", style: 1 }],
    [{ text: "Total Records", style: 0 }, { number: summary.totalRecords, style: 0 }],
    [{ text: "Exact Matches", style: 3 }, { number: summary.matched, style: 0 }],
    [{ text: "Partial Matches", style: 4 }, { number: summary.partial, style: 0 }],
    [{ text: "Mismatches", style: 5 }, { number: summary.mismatched, style: 0 }],
    [{ text: "Missing Values", style: 6 }, { number: summary.missing, style: 0 }],
    [{ text: "Unmatched Records", style: 6 }, { number: summary.unmatched, style: 0 }],
    [{ text: "Duplicate Records", style: 6 }, { number: summary.duplicates, style: 0 }],
    [{ text: "Exception Count", style: 0 }, { number: summary.exceptions, style: 0 }],
    [{ text: "Overall Match %", style: 0 }, { number: summary.matchPercentage, style: 0 }],
  ];

  // Sheet 2: Detailed Reconciliation
  const detailHeader: SheetCell[] = [
    { text: "Record Key", style: 1 }, { text: "Field", style: 1 }, { text: "Source Value", style: 1 },
    { text: "Target Value", style: 1 }, { text: "Status", style: 1 }, { text: "Comment", style: 1 },
  ];
  const detailRows: SheetCell[][] = [detailHeader];
  records.forEach((record) => {
    record.fields.forEach((field) => {
      const style = STATUS_STYLE[field.status] ?? 0;
      detailRows.push([
        { text: record.key, style: 0 },
        { text: field.fieldLabel, style: 0 },
        { text: field.sourceValue, style: 0 },
        { text: field.targetValue, style: 0 },
        { text: field.status, style },
        { text: field.comment, style: 0 },
      ]);
    });
  });

  // Sheet 3: Exceptions Only
  const exceptionRows: SheetCell[][] = [detailHeader];
  records.forEach((record) => {
    record.fields.forEach((field) => {
      if (field.status === "Exact Match") return;
      const style = STATUS_STYLE[field.status] ?? 0;
      exceptionRows.push([
        { text: record.key, style: 0 },
        { text: field.fieldLabel, style: 0 },
        { text: field.sourceValue, style: 0 },
        { text: field.targetValue, style: 0 },
        { text: field.status, style },
        { text: field.comment, style: 0 },
      ]);
    });
  });

  // Sheet 4: Data Quality Report
  const qualityRows: SheetCell[][] = [
    [{ text: "Source", style: 1 }, { text: "Location", style: 1 }, { text: "Field", style: 1 }, { text: "Issue", style: 1 }],
    ...quality.map((issue) => [
      { text: issue.source, style: 0 },
      { text: issue.rowRef, style: 0 },
      { text: issue.field, style: 0 },
      { text: issue.issue, style: 6 },
    ]),
  ];

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="Detailed Reconciliation" sheetId="2" r:id="rId2"/><sheet name="Exceptions Only" sheetId="3" r:id="rId3"/><sheet name="Data Quality Report" sheetId="4" r:id="rId4"/></sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/styles.xml": strToU8(REPORT_STYLES_XML),
    "xl/worksheets/sheet1.xml": strToU8(sheetFromRows(summaryRows)),
    "xl/worksheets/sheet2.xml": strToU8(sheetFromRows(detailRows)),
    "xl/worksheets/sheet3.xml": strToU8(sheetFromRows(exceptionRows)),
    "xl/worksheets/sheet4.xml": strToU8(sheetFromRows(qualityRows)),
  });
}

// ---------------------------------------------------------------------------
// Macro library (searchable list of common reconciliation macros/workflows)
// ---------------------------------------------------------------------------

export type MacroDefinition = {
  id: string;
  name: string;
  category: string;
  description: string;
  keyField: string;
  fieldKeys: string[];
};

export const MACRO_LIBRARY: MacroDefinition[] = [
  { id: "client-recon", name: "Client Data Reconciliation", category: "Reconciliation", description: "Compare client identity, ID, and contact details across two client registers.", keyField: "saId", fieldKeys: ["clientName", "saId", "email", "mobile", "addressLine1", "postalCode"] },
  { id: "investor-recon", name: "Investor Data Reconciliation", category: "Reconciliation", description: "Reconcile investor registers by investor number, name, and holdings.", keyField: "investorNumber", fieldKeys: ["investorNumber", "clientName", "email", "portfolioNumber"] },
  { id: "account-recon", name: "Account Reconciliation", category: "Reconciliation", description: "Match account numbers, holders, and reference numbers between systems.", keyField: "accountNumber", fieldKeys: ["accountNumber", "accountHolder", "clientReference", "accountType"] },
  { id: "banking-recon", name: "Banking Details Reconciliation", category: "Banking", description: "Validate bank name, branch code, and account numbers between two files.", keyField: "accountNumber", fieldKeys: ["accountHolder", "bankName", "branchCode", "accountNumber", "accountType"] },
  { id: "shareholder-register", name: "Shareholder Register Reconciliation", category: "Register", description: "Compare shareholder registers by shareholder number and name.", keyField: "shareholderNumber", fieldKeys: ["shareholderNumber", "clientName", "saId", "companyReg"] },
  { id: "beneficiary-recon", name: "Beneficiary List Reconciliation", category: "Register", description: "Reconcile beneficiary lists by name and ID with allocation checks.", keyField: "saId", fieldKeys: ["clientName", "saId", "addressLine1", "mobile"] },
  { id: "broker-recon", name: "Broker List Reconciliation", category: "Register", description: "Match broker numbers and details across broker registers.", keyField: "brokerNumber", fieldKeys: ["brokerNumber", "clientName", "email", "telephone"] },
  { id: "holdings-recon", name: "Holdings Register Reconciliation", category: "Register", description: "Reconcile holdings registers by portfolio and investor number.", keyField: "portfolioNumber", fieldKeys: ["portfolioNumber", "investorNumber", "clientName"] },
  { id: "kyc-fica", name: "KYC / FICA Validation", category: "Compliance", description: "Validate KYC and FICA identity and address fields for completeness.", keyField: "saId", fieldKeys: ["clientName", "saId", "passport", "addressLine1", "postalCode"] },
  { id: "crs-fatca", name: "CRS / FATCA Validation", category: "Compliance", description: "Validate CRS and FATCA country and tax number information.", keyField: "taxNumber", fieldKeys: ["clientName", "taxNumber", "crsCountry", "fatcaCountry"] },
  { id: "director-validation", name: "Director Validation", category: "Compliance", description: "Validate company directors by ID and company registration.", keyField: "saId", fieldKeys: ["clientName", "saId", "companyReg", "companyName"] },
  { id: "trust-validation", name: "Trust Information Validation", category: "Compliance", description: "Validate trust name and trust registration information.", keyField: "trustReg", fieldKeys: ["trustName", "trustReg", "clientName"] },
  { id: "address-verify", name: "Address Verification", category: "Document Validation", description: "Compare and standardize address fields for proof of address checks.", keyField: "clientName", fieldKeys: ["clientName", "addressLine1", "suburb", "city", "postalCode", "country"] },
  { id: "contact-cleanse", name: "Contact Detail Cleansing", category: "Cleansing", description: "Standardise and validate phone numbers and email addresses.", keyField: "email", fieldKeys: ["clientName", "email", "mobile", "telephone"] },
  { id: "change-detection", name: "Change of Details Detection", category: "Change Detection", description: "Detect changed personal, banking, and address details between two dates.", keyField: "saId", fieldKeys: ["clientName", "addressLine1", "mobile", "email", "accountNumber", "bankName"] },
];
