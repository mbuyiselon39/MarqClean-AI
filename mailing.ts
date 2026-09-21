import { strToU8, zipSync } from "fflate";
import {
  columnIndexToReference,
  escapeXml,
  extractPdfText,
  sanitize,
  similarity,
  standardizeAddress,
  standardizeEmail,
  standardizeName,
  standardizePhone,
  standardizePostal,
  type DataTable,
} from "./engine";

// ---------------------------------------------------------------------------
// Mailing platform fields (the master mailing schema)
// ---------------------------------------------------------------------------

export type MailingFieldKey =
  | "shareholderNumber"
  | "issuer"
  | "fullName"
  | "title"
  | "addressLine1"
  | "addressLine2"
  | "addressLine3"
  | "addressLine4"
  | "postalCode"
  | "revenueServiceNature"
  | "email"
  | "cellPhone"
  | "campaignType"
  | "fromReplyEmail"
  | "comments"
  | "notes";

export type MailingField = {
  key: MailingFieldKey;
  label: string;
  aliases: string[];
  kind: "name" | "title" | "address" | "postal" | "email" | "phone" | "generic";
};

export const MAILING_FIELDS: MailingField[] = [
  { key: "shareholderNumber", label: "Shareholder Number", kind: "generic", aliases: ["shareholder number", "shareholder no", "share no", "holder number", "account number", "sh no"] },
  { key: "issuer", label: "Issuer", kind: "generic", aliases: ["issuer", "company", "share", "instrument"] },
  { key: "fullName", label: "Full Name", kind: "name", aliases: ["full name", "name", "client name", "holder name", "surname and initials", "registered name"] },
  { key: "title", label: "Title", kind: "title", aliases: ["title", "salutation"] },
  { key: "addressLine1", label: "Address Line 1", kind: "address", aliases: ["address line 1", "address1", "addr1", "address 1", "line 1", "physical address", "postal address"] },
  { key: "addressLine2", label: "Address Line 2", kind: "address", aliases: ["address line 2", "address2", "addr2", "address 2", "line 2"] },
  { key: "addressLine3", label: "Address Line 3", kind: "address", aliases: ["address line 3", "address3", "addr3", "address 3", "line 3"] },
  { key: "addressLine4", label: "Address Line 4", kind: "address", aliases: ["address line 4", "address4", "addr4", "address 4", "line 4", "suburb", "city"] },
  { key: "postalCode", label: "Postal Code", kind: "postal", aliases: ["postal code", "postcode", "post code", "zip", "code"] },
  { key: "revenueServiceNature", label: "Revenue Service Nature", kind: "generic", aliases: ["revenue service nature", "record type", "service nature", "nature", "type"] },
  { key: "email", label: "Email Address", kind: "email", aliases: ["email", "e-mail", "email address", "mail"] },
  { key: "cellPhone", label: "Cell Phone Number", kind: "phone", aliases: ["cell", "cellphone", "cell phone", "cell number", "mobile", "mobile number", "contact number"] },
  { key: "campaignType", label: "Campaign Type", kind: "generic", aliases: ["campaign type", "campaign", "delivery method", "method", "channel"] },
  { key: "fromReplyEmail", label: "From and Reply Email", kind: "email", aliases: ["from and reply email", "reply email", "from email", "reply to"] },
  { key: "comments", label: "Comments", kind: "generic", aliases: ["comments", "comment"] },
  { key: "notes", label: "Notes", kind: "generic", aliases: ["notes", "note", "remarks"] },
];

export function getMailingField(key: string): MailingField | undefined {
  return MAILING_FIELDS.find((field) => field.key === key);
}

function normHeader(header: string): string {
  return header.toLowerCase().replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
}

export function autoDetectMailingMapping(table: DataTable): Record<number, string> {
  const mapping: Record<number, string> = {};
  const used = new Set<string>();

  table.headers.forEach((header, index) => {
    const cleaned = normHeader(header);
    let bestKey = "";
    let bestScore = 0;

    MAILING_FIELDS.forEach((field) => {
      if (used.has(field.key)) return;
      field.aliases.forEach((alias) => {
        let score = 0;
        if (cleaned === alias) score = 100;
        else if (cleaned.includes(alias) || alias.includes(cleaned)) score = 70;
        if (score > bestScore) {
          bestScore = score;
          bestKey = field.key;
        }
      });
    });

    if (bestScore >= 60 && bestKey) {
      mapping[index] = bestKey;
      used.add(bestKey);
    } else {
      mapping[index] = "";
    }
  });

  return mapping;
}

// ---------------------------------------------------------------------------
// PDF free-text -> structured mailing records
// ---------------------------------------------------------------------------

export type ExtractedMailingRecord = {
  fullName: string;
  title: string;
  addressLine1: string;
  addressLine2: string;
  addressLine3: string;
  addressLine4: string;
  postalCode: string;
  email: string;
  cellPhone: string;
  confidence: number;
  page: number;
  sourceText: string;
};

const TITLE_PATTERN = /^(mr|mrs|ms|miss|dr|prof|adv|sir|mev|mnr)\b\.?/i;
const PROVINCE_WORDS = ["gauteng", "kwazulu", "natal", "western cape", "eastern cape", "northern cape", "free state", "limpopo", "mpumalanga", "north west"];

function looksLikeName(token: string): boolean {
  return /[a-z]/i.test(token) && !/\d/.test(token) && token.length <= 40;
}

function findPostalCode(text: string): string {
  // South African postal codes are 4 digits, usually at the end
  const trailing = text.trim().match(/(\d{4})\s*$/);
  if (trailing) return trailing[1];
  const anywhere = text.match(/\b(\d{4})\b/g);
  return anywhere ? anywhere[anywhere.length - 1] : "";
}

function findEmail(text: string): string {
  const match = text.match(/[^\s@]+@[^\s@]+\.[^\s@]+/);
  return match ? match[0] : "";
}

function findCell(text: string): string {
  const match = text.match(/(\+?\d[\d\s().-]{8,}\d)/);
  return match ? match[0].trim() : "";
}

// Split a single free-text mailing block into structured fields.
export function structureMailingBlock(block: string, page: number): ExtractedMailingRecord {
  const original = sanitize(block);
  const email = findEmail(original);
  const cellPhone = findCell(original);
  const postalCode = findPostalCode(original);

  // Work on a copy without email so it does not pollute name/address
  let working = original.replace(email, " ").replace(/\s+/g, " ").trim();

  // Detect title
  let title = "";
  const titleMatch = working.match(TITLE_PATTERN);
  if (titleMatch) {
    title = titleMatch[1].replace(/\.$/, "");
    working = working.slice(titleMatch[0].length).trim();
  }

  // Address keyword anchors: text from the first address token onward is the address
  const addressAnchor = working.search(/\b(p\s*\.?\s*o\s*\.?\s*box|post\s?net|private\s?bag|unit|flat|no\.?|number|\d+\s|street|str\b|road|rd\b|ave|avenue|drive|dr\b|lane|ln\b|suite)\b/i);

  let namePart = "";
  let addressPart = "";

  if (addressAnchor > 0) {
    namePart = working.slice(0, addressAnchor).trim();
    addressPart = working.slice(addressAnchor).trim();
  } else {
    // Fallback: assume the first 2-4 capitalised words are the name
    const tokens = working.split(" ");
    const nameTokens: string[] = [];
    for (const token of tokens) {
      if (looksLikeName(token) && nameTokens.length < 4) nameTokens.push(token);
      else break;
    }
    namePart = nameTokens.join(" ");
    addressPart = working.slice(namePart.length).trim();
  }

  // Remove trailing postal code from the address part
  if (postalCode) {
    addressPart = addressPart.replace(new RegExp(`\\b${postalCode}\\b\\s*$`), "").trim();
  }

  // Split address into up to 4 lines using PO BOX and place-name heuristics
  const addressLines = splitAddressLines(addressPart);

  // Confidence heuristic
  let confidence = 50;
  if (namePart) confidence += 20;
  if (addressLines[0]) confidence += 15;
  if (postalCode) confidence += 15;
  confidence = Math.min(confidence, 99);

  return {
    fullName: sanitize(namePart).toUpperCase(),
    title: title.toUpperCase(),
    addressLine1: addressLines[0] ?? "",
    addressLine2: addressLines[1] ?? "",
    addressLine3: addressLines[2] ?? "",
    addressLine4: addressLines[3] ?? "",
    postalCode,
    email,
    cellPhone,
    confidence,
    page,
    sourceText: original,
  };
}

function splitAddressLines(address: string): string[] {
  const clean = sanitize(address);
  if (!clean) return [];

  // Normalise PO BOX so it stays on one line
  const normalized = clean.replace(/\bp\s*\.?\s*o\s*\.?\s*box\b/gi, "PO BOX");

  // If commas exist, prefer them as separators
  if (normalized.includes(",")) {
    return normalized.split(",").map((s) => s.trim().toUpperCase()).filter(Boolean).slice(0, 4);
  }

  const tokens = normalized.split(" ");
  const lines: string[] = [];
  let current: string[] = [];

  const flush = () => {
    if (current.length) {
      lines.push(current.join(" ").toUpperCase());
      current = [];
    }
  };

  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i];
    current.push(token);

    // Break after a PO BOX number
    if (/box/i.test(token) && /\d/.test(tokens[i + 1] ?? "")) {
      current.push(tokens[i + 1]);
      i += 1;
      flush();
      continue;
    }

    // Break when a token looks like a place-name boundary (all caps single word after several tokens)
    const next = tokens[i + 1];
    if (current.length >= 2 && next && looksLikeName(next) && next.length > 3 && PROVINCE_WORDS.every((p) => !next.toLowerCase().includes(p))) {
      // Heuristic: keep grouping; only flush if we already have 2+ words
      if (current.length >= 2 && lines.length < 3) flush();
    }
  }
  flush();

  // If everything collapsed into one line, try splitting by 2+ spaces already handled; otherwise return single
  return lines.slice(0, 4);
}

// Split raw PDF text into record blocks (one mailing entry per block)
function splitPdfIntoBlocks(text: string): Array<{ text: string; page: number }> {
  const lines = text.split(/\r?\n/);
  const blocks: Array<{ text: string; page: number }> = [];
  let current: string[] = [];
  let page = 1;

  const flush = () => {
    const joined = sanitize(current.join(" "));
    if (joined.length > 6) blocks.push({ text: joined, page });
    current = [];
  };

  lines.forEach((rawLine) => {
    const line = sanitize(rawLine);
    if (!line) {
      flush();
      return;
    }
    // A 4-digit postal code at the end typically ends a mailing block
    current.push(line);
    if (/\d{4}\s*$/.test(line)) flush();
  });
  flush();

  return blocks;
}

export async function extractMailingRecordsFromPdf(file: File): Promise<ExtractedMailingRecord[]> {
  const text = await extractPdfText(file);
  const blocks = splitPdfIntoBlocks(text);
  return blocks.map((block) => structureMailingBlock(block.text, block.page));
}

// Also allow structuring from an already-loaded text table (CSV/Word single-column free text)
export function extractMailingRecordsFromTable(table: DataTable): ExtractedMailingRecord[] {
  // If the table already has address-like columns, join each row into a text block
  return table.rows
    .map((row) => structureMailingBlock(row.filter(Boolean).join(" "), 1))
    .filter((record) => record.fullName || record.addressLine1);
}

// ---------------------------------------------------------------------------
// Verification of extracted records against the Excel master
// ---------------------------------------------------------------------------

export type MailingMatchStatus = "Exact Match" | "Partial Match" | "Mismatch" | "Missing";

export type MailingFieldResult = {
  field: string;
  pdfValue: string;
  excelValue: string;
  status: MailingMatchStatus;
  confidence: number;
  exceptionType: string;
  comments: string;
};

export type MailingRecordResult = {
  recordNumber: number;
  sourceFile: string;
  excelRow: number | null;
  shareholderNumber: string;
  issuer: string;
  extracted: ExtractedMailingRecord;
  fields: MailingFieldResult[];
  overall: MailingMatchStatus;
};

export type MailingMapping = Record<number, string>; // excel column index -> mailing field key

export type MailingCompareConfig = {
  fieldKey: MailingFieldKey;
};

export type MailingSummary = {
  pdfRecords: number;
  excelRecords: number;
  matches: number;
  partial: number;
  mismatches: number;
  missing: number;
  exceptions: number;
  matchPercentage: number;
};

export type MailingResult = {
  records: MailingRecordResult[];
  summary: MailingSummary;
  quality: Array<{ location: string; issue: string }>;
  excelMaster: DataTable;
  mapping: MailingMapping;
  selectedFields: MailingFieldKey[];
};

function excelColumnFor(mapping: MailingMapping, key: string): number {
  const entry = Object.entries(mapping).find(([, value]) => value === key);
  return entry ? Number(entry[0]) : -1;
}

function pdfValueFor(record: ExtractedMailingRecord, key: MailingFieldKey): string {
  switch (key) {
    case "fullName": return record.fullName;
    case "title": return record.title;
    case "addressLine1": return record.addressLine1;
    case "addressLine2": return record.addressLine2;
    case "addressLine3": return record.addressLine3;
    case "addressLine4": return record.addressLine4;
    case "postalCode": return record.postalCode;
    case "email": return record.email;
    case "cellPhone": return record.cellPhone;
    default: return "";
  }
}

function standardizeMailing(value: string, kind: MailingField["kind"]): string {
  switch (kind) {
    case "name": return standardizeName(value);
    case "title": return sanitize(value).toLowerCase().replace(/\./g, "");
    case "address": return standardizeAddress(value);
    case "postal": return standardizePostal(value);
    case "email": return standardizeEmail(value);
    case "phone": return standardizePhone(value);
    default: return sanitize(value).toLowerCase();
  }
}

function compareMailing(pdfValue: string, excelValue: string, kind: MailingField["kind"]): { status: MailingMatchStatus; score: number } {
  const pdfEmpty = !sanitize(pdfValue);
  const excelEmpty = !sanitize(excelValue);
  if (pdfEmpty && excelEmpty) return { status: "Missing", score: 0 };
  if (pdfEmpty || excelEmpty) return { status: "Missing", score: 0 };

  const p = standardizeMailing(pdfValue, kind);
  const e = standardizeMailing(excelValue, kind);
  if (p === e) return { status: "Exact Match", score: 1 };

  // For addresses, also compare as combined token overlap
  const score = similarity(p, e);
  if (score >= 0.82) return { status: "Partial Match", score };
  return { status: "Mismatch", score };
}

function buildExcelRowKey(row: string[], mapping: MailingMapping, key: MailingFieldKey, kind: MailingField["kind"]): string {
  const col = excelColumnFor(mapping, key);
  if (col < 0) return "";
  return standardizeMailing(row[col] ?? "", kind);
}

export function runMailingVerification(
  extracted: ExtractedMailingRecord[],
  excel: DataTable,
  mapping: MailingMapping,
  selectedFields: MailingFieldKey[],
  matchKey: MailingFieldKey,
  sourceFile: string
): MailingResult {
  const keyField = getMailingField(matchKey);
  const keyKind = keyField?.kind ?? "generic";

  // Index excel rows by the matching key
  const excelIndex = new Map<string, number[]>();
  excel.rows.forEach((row, index) => {
    const key = buildExcelRowKey(row, mapping, matchKey, keyKind);
    if (!key) return;
    const bucket = excelIndex.get(key) ?? [];
    bucket.push(index);
    excelIndex.set(key, bucket);
  });

  const shareholderCol = excelColumnFor(mapping, "shareholderNumber");
  const issuerCol = excelColumnFor(mapping, "issuer");
  const emailCol = excelColumnFor(mapping, "email");
  const campaignCol = excelColumnFor(mapping, "campaignType");
  const cellCol = excelColumnFor(mapping, "cellPhone");

  const records: MailingRecordResult[] = [];
  const quality: Array<{ location: string; issue: string }> = [];
  const matchedExcelRows = new Set<number>();

  const summary: MailingSummary = {
    pdfRecords: extracted.length, excelRecords: excel.rows.length,
    matches: 0, partial: 0, mismatches: 0, missing: 0, exceptions: 0, matchPercentage: 0,
  };

  extracted.forEach((record, recordIndex) => {
    const pdfKeyValue = pdfValueFor(record, matchKey);
    const stdKey = standardizeMailing(pdfKeyValue, keyKind);
    const candidates = stdKey ? excelIndex.get(stdKey) ?? [] : [];
    const excelRowIndex = candidates.find((i) => !matchedExcelRows.has(i)) ?? candidates[0];
    const hasExcel = excelRowIndex !== undefined;
    if (hasExcel) matchedExcelRows.add(excelRowIndex);

    const excelRow = hasExcel ? excel.rows[excelRowIndex] : [];
    const shareholderNumber = hasExcel && shareholderCol >= 0 ? sanitize(excelRow[shareholderCol] ?? "") : "";
    const issuer = hasExcel && issuerCol >= 0 ? sanitize(excelRow[issuerCol] ?? "") : "";

    const fieldResults: MailingFieldResult[] = [];
    const rank: Record<MailingMatchStatus, number> = { "Exact Match": 0, "Partial Match": 1, "Missing": 2, "Mismatch": 3 };
    let overall: MailingMatchStatus = "Exact Match";

    selectedFields.forEach((fieldKey) => {
      const field = getMailingField(fieldKey);
      const kind = field?.kind ?? "generic";
      const pdfValue = pdfValueFor(record, fieldKey);
      const col = excelColumnFor(mapping, fieldKey);
      const excelValue = hasExcel && col >= 0 ? sanitize(excelRow[col] ?? "") : "";

      const comparison = hasExcel ? compareMailing(pdfValue, excelValue, kind) : { status: "Missing" as MailingMatchStatus, score: 0 };

      let exceptionType = "";
      let comments = "";
      if (comparison.status === "Missing") {
        exceptionType = !sanitize(pdfValue) ? "Missing in PDF" : !sanitize(excelValue) ? "Missing in Excel" : "No matching Excel record";
        comments = hasExcel ? "One side is blank" : "No matching Excel row found";
      } else if (comparison.status === "Partial Match") {
        exceptionType = "Partial";
        comments = `Similarity ${(comparison.score * 100).toFixed(0)}% - review`;
      } else if (comparison.status === "Mismatch") {
        exceptionType = "Mismatch";
        comments = "Values differ - investigate";
      }

      fieldResults.push({
        field: field?.label ?? fieldKey,
        pdfValue,
        excelValue,
        status: comparison.status,
        confidence: Math.round(comparison.score * 100),
        exceptionType,
        comments,
      });

      if (rank[comparison.status] > rank[overall]) overall = comparison.status;
    });

    // Campaign type validation
    if (hasExcel && campaignCol >= 0) {
      const campaign = sanitize(excelRow[campaignCol] ?? "").toLowerCase();
      const emailVal = emailCol >= 0 ? sanitize(excelRow[emailCol] ?? "") : "";
      const cellVal = cellCol >= 0 ? sanitize(excelRow[cellCol] ?? "") : "";
      const comments = getMailingField("comments");
      const commentsCol = comments ? excelColumnFor(mapping, "comments") : -1;
      const commentText = commentsCol >= 0 ? sanitize(excelRow[commentsCol] ?? "").toLowerCase() : "";

      if (campaign.includes("email") && !emailVal) {
        quality.push({ location: `Excel row ${excelRowIndex! + 2}`, issue: "Campaign Type is Email but Email Address is blank" });
      }
      if (campaign.includes("email") && /obtain|to be obtained|await|no email/.test(commentText)) {
        quality.push({ location: `Excel row ${excelRowIndex! + 2}`, issue: "Campaign is Email but comments indicate email still to be obtained" });
      }
      if (campaign.includes("post") && !record.addressLine1) {
        quality.push({ location: `PDF record ${recordIndex + 1}`, issue: "Campaign Type is Post but address details are incomplete" });
      }
      if (campaign.includes("email") && emailVal && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailVal)) {
        quality.push({ location: `Excel row ${excelRowIndex! + 2}`, issue: "Invalid email format for an Email campaign" });
      }
      if (cellVal && /e\+|E\+/.test(cellVal)) {
        quality.push({ location: `Excel row ${excelRowIndex! + 2}`, issue: "Cell number stored as scientific notation" });
      }
    }

    const finalOverall: MailingMatchStatus = hasExcel ? overall : "Missing";

    records.push({
      recordNumber: recordIndex + 1,
      sourceFile,
      excelRow: hasExcel ? excelRowIndex! + 2 : null,
      shareholderNumber,
      issuer,
      extracted: record,
      fields: fieldResults,
      overall: finalOverall,
    });

    const overallValue = finalOverall as string;
    if (overallValue === "Exact Match") summary.matches += 1;
    else if (overallValue === "Partial Match") summary.partial += 1;
    else if (overallValue === "Mismatch") summary.mismatches += 1;
    else summary.missing += 1;
  });

  // Duplicate detection in excel master (names and shareholder numbers)
  const nameSeen = new Map<string, number>();
  const shSeen = new Map<string, number>();
  const nameCol = excelColumnFor(mapping, "fullName");
  excel.rows.forEach((row, index) => {
    if (nameCol >= 0) {
      const name = standardizeName(row[nameCol] ?? "");
      if (name) {
        const count = (nameSeen.get(name) ?? 0) + 1;
        nameSeen.set(name, count);
        if (count === 2) quality.push({ location: `Excel row ${index + 2}`, issue: "Duplicate full name in master" });
      }
    }
    if (shareholderCol >= 0) {
      const sh = sanitize(row[shareholderCol] ?? "").toUpperCase();
      if (sh) {
        const count = (shSeen.get(sh) ?? 0) + 1;
        shSeen.set(sh, count);
        if (count === 2) quality.push({ location: `Excel row ${index + 2}`, issue: "Duplicate shareholder number in master" });
      }
    }
    // Missing / invalid checks
    if (emailCol >= 0) {
      const email = sanitize(row[emailCol] ?? "");
      if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) quality.push({ location: `Excel row ${index + 2}`, issue: "Invalid email format" });
    }
    const postalCol = excelColumnFor(mapping, "postalCode");
    if (postalCol >= 0) {
      const postal = standardizePostal(row[postalCol] ?? "");
      if (postal && postal.length !== 4) quality.push({ location: `Excel row ${index + 2}`, issue: "Postal code is not 4 digits" });
    }
  });

  summary.exceptions = summary.mismatches + summary.missing;
  const denom = summary.pdfRecords || 1;
  summary.matchPercentage = Math.round((summary.matches / denom) * 1000) / 10;

  return { records, summary, quality, excelMaster: excel, mapping, selectedFields };
}

// ---------------------------------------------------------------------------
// 6-sheet colour-coded Excel report
// ---------------------------------------------------------------------------

const STATUS_STYLE: Record<string, number> = {
  "Exact Match": 3, "Partial Match": 4, "Mismatch": 5, "Missing": 6,
};

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/><color rgb="FF0F172A"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font><font><b/><sz val="14"/><name val="Calibri"/><color rgb="FF0F172A"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCDEFD6"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCFE3FB"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF9C9CC"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE2BD"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>`;

type Cell = { text?: string; number?: number; style: number };

function cellXml(ref: string, cell: Cell): string {
  if (cell.number !== undefined) return `<c r="${ref}" s="${cell.style}"><v>${cell.number}</v></c>`;
  const text = cell.text ?? "";
  if (!text) return `<c r="${ref}" s="${cell.style}"/>`;
  return `<c r="${ref}" t="inlineStr" s="${cell.style}"><is><t>${escapeXml(text)}</t></is></c>`;
}

function sheetXml(rows: Cell[][]): string {
  const cols = rows.reduce((max, row) => Math.max(max, row.length), 1);
  const body = rows
    .map((row, r) => `<row r="${r + 1}">${row.map((cell, c) => cellXml(`${columnIndexToReference(c)}${r + 1}`, cell)).join("")}</row>`)
    .join("");
  const colDefs = Array.from({ length: cols }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="24" customWidth="1"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${colDefs}</cols><sheetData>${body}</sheetData></worksheet>`;
}

export function buildMailingWorkbook(result: MailingResult, pdfName: string, excelName: string): Uint8Array {
  const { summary, records, quality, excelMaster, mapping, selectedFields } = result;

  // Sheet 1: Executive Summary
  const s1: Cell[][] = [
    [{ text: "Mailing Verification - Executive Summary", style: 2 }],
    [{ text: `PDF source: ${pdfName}`, style: 0 }],
    [{ text: `Excel master: ${excelName}`, style: 0 }],
    [],
    [{ text: "Metric", style: 1 }, { text: "Value", style: 1 }],
    [{ text: "Total PDF Records Extracted", style: 0 }, { number: summary.pdfRecords, style: 0 }],
    [{ text: "Total Excel Records Checked", style: 0 }, { number: summary.excelRecords, style: 0 }],
    [{ text: "Total Matches", style: 3 }, { number: summary.matches, style: 0 }],
    [{ text: "Total Partial Matches", style: 4 }, { number: summary.partial, style: 0 }],
    [{ text: "Total Mismatches", style: 5 }, { number: summary.mismatches, style: 0 }],
    [{ text: "Total Missing Values", style: 6 }, { number: summary.missing, style: 0 }],
    [{ text: "Overall Match Percentage", style: 0 }, { number: summary.matchPercentage, style: 0 }],
    [{ text: "Total Exceptions", style: 0 }, { number: summary.exceptions, style: 0 }],
  ];

  // Sheet 2: Detailed Verification Report
  const detailHeader: Cell[] = [
    { text: "Record Number", style: 1 }, { text: "Source File", style: 1 }, { text: "Excel Row", style: 1 },
    { text: "Shareholder Number", style: 1 }, { text: "Issuer", style: 1 }, { text: "Field Compared", style: 1 },
    { text: "PDF Extracted Value", style: 1 }, { text: "Excel Value", style: 1 }, { text: "Match Status", style: 1 },
    { text: "Confidence Score", style: 1 }, { text: "Exception Type", style: 1 }, { text: "Comments", style: 1 },
  ];
  const s2: Cell[][] = [detailHeader];
  records.forEach((record) => {
    record.fields.forEach((field) => {
      s2.push([
        { number: record.recordNumber, style: 0 },
        { text: record.sourceFile, style: 0 },
        { text: record.excelRow ? String(record.excelRow) : "", style: 0 },
        { text: record.shareholderNumber, style: 0 },
        { text: record.issuer, style: 0 },
        { text: field.field, style: 0 },
        { text: field.pdfValue, style: 0 },
        { text: field.excelValue, style: 0 },
        { text: field.status, style: STATUS_STYLE[field.status] ?? 0 },
        { number: field.confidence, style: 0 },
        { text: field.exceptionType, style: 0 },
        { text: field.comments, style: 0 },
      ]);
    });
  });

  // Sheet 3: Exceptions Only
  const s3: Cell[][] = [detailHeader];
  records.forEach((record) => {
    record.fields.forEach((field) => {
      if (field.status === "Exact Match") return;
      s3.push([
        { number: record.recordNumber, style: 0 },
        { text: record.sourceFile, style: 0 },
        { text: record.excelRow ? String(record.excelRow) : "", style: 0 },
        { text: record.shareholderNumber, style: 0 },
        { text: record.issuer, style: 0 },
        { text: field.field, style: 0 },
        { text: field.pdfValue, style: 0 },
        { text: field.excelValue, style: 0 },
        { text: field.status, style: STATUS_STYLE[field.status] ?? 0 },
        { number: field.confidence, style: 0 },
        { text: field.exceptionType, style: 0 },
        { text: field.comments, style: 0 },
      ]);
    });
  });

  // Sheet 4: PDF Extraction Output
  const s4: Cell[][] = [[
    { text: "Extracted Full Name", style: 1 }, { text: "Extracted Title", style: 1 },
    { text: "Address Line 1", style: 1 }, { text: "Address Line 2", style: 1 },
    { text: "Address Line 3", style: 1 }, { text: "Address Line 4", style: 1 },
    { text: "Postal Code", style: 1 }, { text: "Email", style: 1 }, { text: "Cell Phone", style: 1 },
    { text: "Extraction Confidence", style: 1 }, { text: "PDF Page", style: 1 }, { text: "Source Text", style: 1 },
  ]];
  records.forEach((record) => {
    const e = record.extracted;
    s4.push([
      { text: e.fullName, style: 0 }, { text: e.title, style: 0 },
      { text: e.addressLine1, style: 0 }, { text: e.addressLine2, style: 0 },
      { text: e.addressLine3, style: 0 }, { text: e.addressLine4, style: 0 },
      { text: e.postalCode, style: 0 }, { text: e.email, style: 0 }, { text: e.cellPhone, style: 0 },
      { number: e.confidence, style: 0 }, { number: e.page, style: 0 }, { text: e.sourceText, style: 0 },
    ]);
  });

  // Sheet 5: Excel Master Data Used
  const usedRows = new Set(records.map((r) => (r.excelRow ? r.excelRow - 2 : -1)).filter((i) => i >= 0));
  const mappedHeaders = excelMaster.headers.map((header, index) => {
    const key = mapping[index];
    const field = key ? getMailingField(key) : undefined;
    return field ? `${header} -> ${field.label}` : header;
  });
  const s5: Cell[][] = [mappedHeaders.map((h) => ({ text: h, style: 1 }))];
  Array.from(usedRows).sort((a, b) => a - b).forEach((rowIndex) => {
    s5.push(excelMaster.rows[rowIndex].map((value) => ({ text: value, style: 0 })));
  });
  if (s5.length === 1) {
    excelMaster.rows.slice(0, 200).forEach((row) => s5.push(row.map((value) => ({ text: value, style: 0 }))));
  }

  // Sheet 6: Data Quality Issues
  const s6: Cell[][] = [[{ text: "Location", style: 1 }, { text: "Issue", style: 1 }]];
  quality.forEach((issue) => s6.push([{ text: issue.location, style: 0 }, { text: issue.issue, style: 6 }]));

  void selectedFields;

  const sheetNames = ["Executive Summary", "Detailed Verification", "Exceptions Only", "PDF Extraction Output", "Excel Master Data", "Data Quality Issues"];
  const sheetData = [s1, s2, s3, s4, s5, s6].map(sheetXml);

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetNames.map((name, i) => `<sheet name="${escapeXml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`;
  const relEntries = sheetNames.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("");
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntries}<Relationship Id="rId${sheetNames.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const overrides = sheetNames.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/styles.xml": strToU8(STYLES_XML),
  };
  sheetData.forEach((xml, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(xml);
  });

  return zipSync(files);
}
