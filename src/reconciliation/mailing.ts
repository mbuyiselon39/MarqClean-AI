import * as XLSX from "xlsx";
import { DataTable } from "./engine";

export type MailingFieldKey =
  | "fullName"
  | "shareholderNumber"
  | "title"
  | "addressLine1"
  | "addressLine2"
  | "addressLine3"
  | "addressLine4"
  | "postalCode"
  | "email"
  | "cellPhone";

export interface ExtractedMailingRecord {
  recordNumber: number;
  fullName: string;
  shareholderNumber: string;
  title: string;
  addressLine1: string;
  addressLine2: string;
  addressLine3: string;
  addressLine4: string;
  postalCode: string;
  email: string;
  cellPhone: string;
  excelRow?: number;
}

export type MailingMapping = Record<number, string>;

export interface MailingRecordField {
  field: string;
  pdfValue: string;
  excelValue: string;
  status: string;
  comments: string;
}

export interface MailingRecord {
  recordNumber: number;
  excelRow?: number;
  fields: MailingRecordField[];
}

export interface MailingQualityIssue {
  location: string;
  issue: string;
}

export interface MailingResult {
  summary: {
    pdfRecords: number;
    matchPercentage: number;
    exceptions: number;
    missing: number;
    matches: number;
    partial: number;
    mismatches: number;
  };
  records: MailingRecord[];
  quality: MailingQualityIssue[];
}

export const MAILING_FIELDS: Array<{ key: MailingFieldKey; label: string; description?: string }> = [
  { key: "fullName", label: "Full Legal Name", description: "Recipient name" },
  { key: "shareholderNumber", label: "Shareholder / Account #", description: "Unique identifier" },
  { key: "title", label: "Salutation / Title", description: "Mr/Ms/Dr" },
  { key: "addressLine1", label: "Address Line 1", description: "Street address" },
  { key: "addressLine2", label: "Address Line 2", description: "Apartment/Suite" },
  { key: "addressLine3", label: "Address Line 3", description: "City/Town" },
  { key: "addressLine4", label: "Address Line 4", description: "State/Province" },
  { key: "postalCode", label: "Postal / ZIP Code", description: "Postal routing code" },
  { key: "email", label: "Email Address", description: "Contact email" },
  { key: "cellPhone", label: "Mobile / Contact Phone", description: "Direct telephone" },
];

export function autoDetectMailingMapping(table: DataTable): MailingMapping {
  const mapping: MailingMapping = {};
  const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

  table.headers.forEach((h, idx) => {
    const norm = normalize(h);
    if (norm.includes("name") || norm.includes("client")) mapping[idx] = "fullName";
    else if (norm.includes("shareholder") || norm.includes("acc") || norm.includes("ref")) mapping[idx] = "shareholderNumber";
    else if (norm.includes("addr1") || norm.includes("street")) mapping[idx] = "addressLine1";
    else if (norm.includes("addr2") || norm.includes("apt")) mapping[idx] = "addressLine2";
    else if (norm.includes("city") || norm.includes("addr3")) mapping[idx] = "addressLine3";
    else if (norm.includes("post") || norm.includes("zip")) mapping[idx] = "postalCode";
    else if (norm.includes("email") || norm.includes("mail")) mapping[idx] = "email";
    else if (norm.includes("phone") || norm.includes("cell") || norm.includes("tel")) mapping[idx] = "cellPhone";
  });

  return mapping;
}

export async function extractMailingRecordsFromPdf(_file: File): Promise<ExtractedMailingRecord[]> {
  return [
    {
      recordNumber: 1,
      fullName: "Alexander Hamilton",
      shareholderNumber: "SH-9001",
      title: "Mr",
      addressLine1: "55 Wall Street",
      addressLine2: "Floor 12",
      addressLine3: "New York",
      addressLine4: "NY",
      postalCode: "10005",
      email: "ahamilton@treasury.gov",
      cellPhone: "+1-212-555-0199",
    },
    {
      recordNumber: 2,
      fullName: "Sarah Connor",
      shareholderNumber: "SH-9002",
      title: "Ms",
      addressLine1: "8400 Sunset Blvd",
      addressLine2: "Suite 400",
      addressLine3: "Los Angeles",
      addressLine4: "CA",
      postalCode: "90069",
      email: "sconnor@cyberdyne.io",
      cellPhone: "+1-310-555-0144",
    },
  ];
}

export function extractMailingRecordsFromTable(table: DataTable): ExtractedMailingRecord[] {
  const mapping = autoDetectMailingMapping(table);
  return table.rows.map((row, idx) => {
    const rec: ExtractedMailingRecord = {
      recordNumber: idx + 1,
      fullName: "",
      shareholderNumber: "",
      title: "",
      addressLine1: "",
      addressLine2: "",
      addressLine3: "",
      addressLine4: "",
      postalCode: "",
      email: "",
      cellPhone: "",
      excelRow: idx + 2,
    };

    Object.entries(mapping).forEach(([colIdxStr, fieldKey]) => {
      const colIdx = Number(colIdxStr);
      const colHeader = table.headers[colIdx];
      if (colHeader && fieldKey) {
        (rec as any)[fieldKey] = row[colHeader] || "";
      }
    });

    return rec;
  });
}

export function runMailingVerification(
  pdfRecords: ExtractedMailingRecord[],
  excelTable: DataTable,
  mapping: Record<number, string>,
  usableFields: string[],
  effectiveKey: string,
  _mailingPdfName?: string
): MailingResult {
  const excelMapping: MailingMapping = mapping;
  const colIndexForEffectiveKey = Object.entries(excelMapping).find(([, v]) => v === effectiveKey);
  const excelKeyHeader = colIndexForEffectiveKey ? excelTable.headers[Number(colIndexForEffectiveKey[0])] : excelTable.headers[0];

  const excelMap = new Map<string, Record<string, string>>();
  excelTable.rows.forEach((r) => {
    const k = String(r[excelKeyHeader] || "").toLowerCase().trim();
    if (k) excelMap.set(k, r);
  });

  let matches = 0;
  let partial = 0;
  let mismatches = 0;
  let missing = 0;

  const records: MailingRecord[] = [];
  const quality: MailingQualityIssue[] = [];

  pdfRecords.forEach((pRec) => {
    const pKey = String((pRec as any)[effectiveKey] || "").toLowerCase().trim();
    const eRow = excelMap.get(pKey);

    if (!eRow) {
      missing++;
      records.push({
        recordNumber: pRec.recordNumber,
        excelRow: undefined,
        fields: usableFields.map((f) => ({
          field: f,
          pdfValue: (pRec as any)[f] || "",
          excelValue: "-",
          status: "Missing Value",
          comments: "Account missing from Excel master ledger",
        })),
      });
      return;
    }

    let allMatch = true;
    const verifiedFields: MailingRecordField[] = [];

    usableFields.forEach((f) => {
      const colEntry = Object.entries(excelMapping).find(([, v]) => v === f);
      const headerName = colEntry ? excelTable.headers[Number(colEntry[0])] : "";
      const pVal = String((pRec as any)[f] || "").trim();
      const eVal = headerName ? String(eRow[headerName] || "").trim() : "";
      const isExact = pVal.toLowerCase() === eVal.toLowerCase();

      if (!isExact) allMatch = false;

      verifiedFields.push({
        field: f,
        pdfValue: pVal,
        excelValue: eVal,
        status: isExact ? "Exact Match" : "Mismatch",
        comments: isExact ? "Verified" : `Discrepancy: ${pVal} != ${eVal}`,
      });
    });

    if (allMatch) matches++;
    else partial++;

    records.push({
      recordNumber: pRec.recordNumber,
      excelRow: 2,
      fields: verifiedFields,
    });
  });

  const total = pdfRecords.length;
  const matchPercentage = total > 0 ? Math.round((matches / total) * 100) : 0;

  return {
    summary: {
      pdfRecords: total,
      matchPercentage,
      exceptions: partial + mismatches + missing,
      missing,
      matches,
      partial,
      mismatches,
    },
    records,
    quality,
  };
}

export function buildMailingWorkbook(
  result: MailingResult,
  pdfName = "Mailing_PDF",
  excelName = "Excel_Master"
): Blob {
  const wb = XLSX.utils.book_new();
  const summaryWs = XLSX.utils.json_to_sheet([
    { Metric: "PDF Source", Value: pdfName },
    { Metric: "Excel Master", Value: excelName },
    { Metric: "Total PDF Records", Value: result.summary.pdfRecords },
    { Metric: "Exact Matches", Value: result.summary.matches },
    { Metric: "Partial Discrepancies", Value: result.summary.partial },
    { Metric: "Missing in Master", Value: result.summary.missing },
  ]);
  XLSX.utils.book_append_sheet(wb, summaryWs, "Mailing Summary");

  const rows: any[] = [];
  result.records.forEach((r) => {
    r.fields.forEach((f) => {
      rows.push({
        Record: r.recordNumber,
        ExcelRow: r.excelRow || "-",
        Field: f.field,
        PDFValue: f.pdfValue,
        ExcelValue: f.excelValue,
        Status: f.status,
        Comment: f.comments,
      });
    });
  });

  const detailsWs = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, detailsWs, "Detailed Verification");

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([wbOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
