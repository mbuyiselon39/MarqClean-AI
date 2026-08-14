import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Smart Sheet Manager — real workbook analysis and rebuilding, powered by
// SheetJS (Community Edition). Every statistic below is computed from the
// actual parsed workbook, not simulated.
//
// Honest capability notes (SheetJS Community Edition limitations):
// - Formulas, values, number formats, merged cells, defined names, and
//   column widths round-trip through read -> write.
// - A workbook's VBA project (.xlsm) is preserved as an opaque passthrough
//   blob (wb.vbaraw) when present, as long as the output is written back
//   as .xlsm.
// - Charts and PivotTables are NOT preserved on write by the free/community
//   engine — the underlying XML for these is not part of its parsed model.
//   The UI surfaces this clearly rather than silently dropping them.
// ---------------------------------------------------------------------------

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export type SheetInfo = {
  name: string;
  index: number;
  visibility: SheetVisibility;
  rowCount: number;
  colCount: number;
  nonEmptyCellCount: number;
  isEmpty: boolean;
  hasFormulas: boolean;
  hasMergedCells: boolean;
  contentHash: string;
  duplicateOf: string | null;
  referencedByFormulaCount: number;
  keepScore: number;
  keepReasons: string[];
};

export type WorkbookAnalysis = {
  sheets: SheetInfo[];
  totalSheets: number;
  visibleCount: number;
  hiddenCount: number;
  veryHiddenCount: number;
  emptyCount: number;
  duplicateGroups: string[][];
  hasVbaProject: boolean;
  originalSizeBytes: number;
};

const KEEP_NAME_SIGNALS: Array<{ pattern: RegExp; weight: number; reason: string }> = [
  { pattern: /dashboard/i, weight: 30, reason: "Name suggests a dashboard" },
  { pattern: /summary|overview/i, weight: 25, reason: "Name suggests a summary sheet" },
  { pattern: /final|output|result/i, weight: 15, reason: "Name suggests a final output sheet" },
  { pattern: /reconcil/i, weight: 15, reason: "Name suggests a reconciliation sheet" },
  { pattern: /report/i, weight: 12, reason: "Name suggests a report" },
  { pattern: /^sheet\d+$/i, weight: -20, reason: "Default, unrenamed sheet name" },
  { pattern: /^copy of|copy\(\d+\)|_copy$/i, weight: -15, reason: "Name suggests a duplicate/backup copy" },
  { pattern: /archive|old|temp|test|backup|draft/i, weight: -18, reason: "Name suggests archived or temporary content" },
];

function hashString(input: string): string {
  // Small, fast, deterministic non-cryptographic hash (FNV-1a) — enough to
  // group sheets with identical serialized content without pulling in a
  // crypto dependency for what is purely a duplicate-detection heuristic.
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16);
}

function getSheetVisibility(workbook: XLSX.WorkBook, sheetName: string): SheetVisibility {
  const entry = workbook.Workbook?.Sheets?.find((s) => s.name === sheetName);
  if (!entry || !entry.Hidden) return "visible";
  return entry.Hidden === 2 ? "veryHidden" : "hidden";
}

function serializeSheetContent(sheet: XLSX.WorkSheet): string {
  // Serialize values only (not styles) so genuinely identical data is
  // detected as a duplicate even if formatting differs slightly.
  const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
  return csv.trim();
}

export function analyzeWorkbook(workbook: XLSX.WorkBook, originalSizeBytes: number): WorkbookAnalysis {
  const contentToNames = new Map<string, string[]>();
  const referencedCounts = new Map<string, number>();

  // Scan every formula in every sheet for cross-sheet references, e.g.
  // ='Summary'!A1 or =SUM(Dashboard!B2:B10), so we can genuinely detect
  // which sheets are actually relied upon elsewhere in the workbook.
  workbook.SheetNames.forEach((name) => {
    const sheet = workbook.Sheets[name];
    if (!sheet) return;
    Object.keys(sheet).forEach((cellRef) => {
      if (cellRef.startsWith("!")) return;
      const cell = sheet[cellRef] as XLSX.CellObject;
      if (!cell?.f) return;
      workbook.SheetNames.forEach((otherName) => {
        if (otherName === name) return;
        const escaped = otherName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const refPattern = new RegExp(`(^|[^A-Za-z0-9_])(('${escaped}')|(${escaped}))!`, "i");
        if (refPattern.test(cell.f as string)) {
          referencedCounts.set(otherName, (referencedCounts.get(otherName) ?? 0) + 1);
        }
      });
    });
  });

  const sheets: SheetInfo[] = workbook.SheetNames.map((name, index) => {
    const sheet = workbook.Sheets[name];
    const ref = sheet["!ref"];
    const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: -1, c: -1 } };
    const rowCount = Math.max(0, range.e.r - range.s.r + 1);
    const colCount = Math.max(0, range.e.c - range.s.c + 1);

    let nonEmptyCellCount = 0;
    let hasFormulas = false;
    Object.keys(sheet).forEach((cellRef) => {
      if (cellRef.startsWith("!")) return;
      const cell = sheet[cellRef] as XLSX.CellObject;
      if (cell?.f) hasFormulas = true;
      if (cell?.v !== undefined && cell?.v !== "") nonEmptyCellCount++;
    });

    const contentHash = hashString(serializeSheetContent(sheet));
    const group = contentToNames.get(contentHash) ?? [];
    group.push(name);
    contentToNames.set(contentHash, group);

    const hasMergedCells = Array.isArray(sheet["!merges"]) && sheet["!merges"].length > 0;
    const isEmpty = nonEmptyCellCount === 0;

    let keepScore = 40; // neutral baseline
    const keepReasons: string[] = [];
    KEEP_NAME_SIGNALS.forEach(({ pattern, weight, reason }) => {
      if (pattern.test(name)) {
        keepScore += weight;
        keepReasons.push(reason);
      }
    });
    if (hasFormulas) {
      keepScore += 10;
      keepReasons.push("Contains formulas");
    }
    const referencedByFormulaCount = referencedCounts.get(name) ?? 0;
    if (referencedByFormulaCount > 0) {
      keepScore += Math.min(20, referencedByFormulaCount * 5);
      keepReasons.push(`Referenced by formulas in ${referencedByFormulaCount} other cell${referencedByFormulaCount === 1 ? "" : "s"}`);
    }
    if (hasMergedCells) {
      keepScore += 5;
      keepReasons.push("Contains merged cells (often used in dashboards/headers)");
    }
    if (isEmpty) {
      keepScore -= 40;
      keepReasons.push("Sheet is empty");
    }
    if (getSheetVisibility(workbook, name) !== "visible") {
      keepScore -= 10;
      keepReasons.push("Sheet is hidden");
    }
    keepScore = Math.max(0, Math.min(100, keepScore));

    return {
      name,
      index,
      visibility: getSheetVisibility(workbook, name),
      rowCount,
      colCount,
      nonEmptyCellCount,
      isEmpty,
      hasFormulas,
      hasMergedCells,
      contentHash,
      duplicateOf: null,
      referencedByFormulaCount,
      keepScore,
      keepReasons,
    };
  });

  // Resolve duplicate groups: the first sheet in a content group is the
  // "original"; every later sheet with identical content is flagged.
  const duplicateGroups: string[][] = [];
  contentToNames.forEach((names) => {
    if (names.length > 1) {
      duplicateGroups.push(names);
      names.slice(1).forEach((dupeName) => {
        const info = sheets.find((s) => s.name === dupeName);
        if (info) info.duplicateOf = names[0];
      });
    }
  });

  const visibleCount = sheets.filter((s) => s.visibility === "visible").length;
  const hiddenCount = sheets.filter((s) => s.visibility === "hidden").length;
  const veryHiddenCount = sheets.filter((s) => s.visibility === "veryHidden").length;
  const emptyCount = sheets.filter((s) => s.isEmpty).length;

  return {
    sheets,
    totalSheets: sheets.length,
    visibleCount,
    hiddenCount,
    veryHiddenCount,
    emptyCount,
    duplicateGroups,
    hasVbaProject: Boolean((workbook as XLSX.WorkBook & { vbaraw?: unknown }).vbaraw),
    originalSizeBytes,
  };
}

export type FilterResult = {
  bytes: Uint8Array;
  keptSheets: string[];
  removedSheets: string[];
  newSizeBytes: number;
  elapsedMs: number;
  usedXlsm: boolean;
};

/** Rebuilds the workbook containing only the given sheet names, in their original relative order unless reordered. */
export function buildFilteredWorkbook(workbook: XLSX.WorkBook, keepSheetNames: string[]): FilterResult {
  const start = performance.now();
  if (keepSheetNames.length === 0) {
    throw new Error("At least one worksheet must remain — cannot produce a workbook with zero sheets.");
  }

  const keepSet = new Set(keepSheetNames);
  const removedSheets = workbook.SheetNames.filter((name) => !keepSet.has(name));

  const nextWorkbook: XLSX.WorkBook = {
    ...workbook,
    SheetNames: keepSheetNames,
    Sheets: Object.fromEntries(keepSheetNames.map((name) => [name, workbook.Sheets[name]])),
  };
  if (nextWorkbook.Workbook?.Sheets) {
    nextWorkbook.Workbook = {
      ...nextWorkbook.Workbook,
      Sheets: nextWorkbook.Workbook.Sheets.filter((s) => keepSet.has(s.name ?? "")),
    };
  }

  const hasVba = Boolean((workbook as XLSX.WorkBook & { vbaraw?: unknown }).vbaraw);
  const bookType: XLSX.BookType = hasVba ? "xlsm" : "xlsx";
  const out = XLSX.write(nextWorkbook, { type: "array", bookType, cellStyles: true });
  const bytes = new Uint8Array(out as ArrayBuffer);

  return {
    bytes,
    keptSheets: keepSheetNames,
    removedSheets,
    newSizeBytes: bytes.byteLength,
    elapsedMs: performance.now() - start,
    usedXlsm: hasVba,
  };
}

export function renameSheet(workbook: XLSX.WorkBook, oldName: string, newName: string): void {
  const index = workbook.SheetNames.indexOf(oldName);
  if (index === -1) return;
  const sanitized = newName.trim().slice(0, 31).replace(/[\\/*?:[\]]/g, "-") || oldName;
  workbook.SheetNames[index] = sanitized;
  workbook.Sheets[sanitized] = workbook.Sheets[oldName];
  if (sanitized !== oldName) delete workbook.Sheets[oldName];
  const entry = workbook.Workbook?.Sheets?.find((s) => s.name === oldName);
  if (entry) entry.name = sanitized;
}

export function reorderSheets(workbook: XLSX.WorkBook, orderedNames: string[]): void {
  workbook.SheetNames = orderedNames.filter((n) => workbook.SheetNames.includes(n));
}

export function setVisibility(workbook: XLSX.WorkBook, sheetName: string, visibility: SheetVisibility): void {
  if (!workbook.Workbook) workbook.Workbook = {};
  if (!workbook.Workbook.Sheets) workbook.Workbook.Sheets = workbook.SheetNames.map((name) => ({ name }));
  const entry = workbook.Workbook.Sheets.find((s) => s.name === sheetName);
  const hiddenValue = visibility === "visible" ? 0 : visibility === "hidden" ? 1 : 2;
  if (entry) entry.Hidden = hiddenValue;
  else workbook.Workbook.Sheets.push({ name: sheetName, Hidden: hiddenValue });
}

export function buildRemovedSheetsLog(analysis: WorkbookAnalysis, removedSheets: string[]): string {
  const lines = [
    "Smart Sheet Manager — Removed Sheets Log",
    `Generated: ${new Date().toISOString()}`,
    `Original sheet count: ${analysis.totalSheets}`,
    `Removed sheet count: ${removedSheets.length}`,
    `Remaining sheet count: ${analysis.totalSheets - removedSheets.length}`,
    "",
    "Sheet Name,Visibility,Rows,Columns,Non-Empty Cells,Had Formulas,Was Empty,Duplicate Of",
  ];
  removedSheets.forEach((name) => {
    const info = analysis.sheets.find((s) => s.name === name);
    if (!info) {
      lines.push(`${name},unknown,,,,,,`);
      return;
    }
    lines.push(
      [
        info.name,
        info.visibility,
        info.rowCount,
        info.colCount,
        info.nonEmptyCellCount,
        info.hasFormulas ? "Yes" : "No",
        info.isEmpty ? "Yes" : "No",
        info.duplicateOf ?? "",
      ].join(",")
    );
  });
  return lines.join("\n");
}
