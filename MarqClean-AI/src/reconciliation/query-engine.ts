import type { DataTable } from "./engine";
import { mergeAppend, mergeJoin, removeDuplicatesByColumns } from "./toolbox";
import { textSplit, concatColumns, filterTableByCriterion, sortTableByColumn, type Criterion, type CriteriaOp } from "./excel-functions";

// ---------------------------------------------------------------------------
// Power Query-style step pipeline.
//
// Honest scope: this reproduces the *result* of a Power Query transformation
// (cleaned, combined, reshaped data with a full step-by-step log) by composing
// the platform's existing, already-tested transform functions in sequence. It
// does not generate real Power Query "M" code, and a saved recipe here will
// not appear as a live, refreshable query inside Excel's own Data > Queries
// & Connections pane — see the in-product capability notice for the full
// wording surfaced to users.
// ---------------------------------------------------------------------------

export type QueryStepType =
  | "choose-columns"
  | "change-type"
  | "split-column"
  | "merge-columns"
  | "remove-duplicates"
  | "remove-blank-rows"
  | "filter-rows"
  | "sort"
  | "append"
  | "join";

export type DataTypeTarget = "text" | "number" | "date-iso";

export type QueryStep = {
  id: string;
  type: QueryStepType;
  label: string;
  params: Record<string, string | number | boolean>;
};

export type QueryStepOutcome = {
  step: QueryStep;
  table: DataTable;
  message: string;
};

function coerceType(value: string, target: DataTypeTarget): string {
  const trimmed = value.trim();
  if (target === "text") return trimmed;
  if (target === "number") {
    const cleaned = trimmed.replace(/[^0-9.\-]/g, "");
    const n = Number(cleaned);
    return Number.isFinite(n) && cleaned !== "" ? String(n) : "";
  }
  // date-iso
  const d = new Date(trimmed);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

function isBlankOrError(value: string): boolean {
  const v = value.trim();
  if (v === "") return true;
  return /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NULL!|NUM!|CALC!|SPILL!)$/i.test(v);
}

/** Runs one step against a table. `lookupTable` resolves other uploaded source tables by name, needed for append/join. */
export function runQueryStep(table: DataTable, step: QueryStep, lookupTable: (name: string) => DataTable | undefined): QueryStepOutcome {
  switch (step.type) {
    case "choose-columns": {
      const keep = String(step.params.columns ?? "").split(",").map((s) => Number(s)).filter((n) => Number.isInteger(n));
      const headers = keep.map((i) => table.headers[i]);
      const rows = table.rows.map((row) => keep.map((i) => row[i] ?? ""));
      return { step, table: { headers, rows, sourceName: table.sourceName }, message: `Kept ${keep.length} of ${table.headers.length} columns.` };
    }
    case "change-type": {
      const col = Number(step.params.column);
      const target = step.params.target as DataTypeTarget;
      const rows = table.rows.map((row) => {
        const next = row.slice();
        next[col] = coerceType(row[col] ?? "", target);
        return next;
      });
      const failed = rows.filter((row, i) => row[col] === "" && (table.rows[i][col] ?? "").trim() !== "").length;
      return {
        step,
        table: { ...table, rows },
        message: failed > 0 ? `Converted ${table.headers[col]} to ${target}. ${failed} value(s) could not convert and became blank.` : `Converted ${table.headers[col]} to ${target}.`,
      };
    }
    case "split-column": {
      const col = Number(step.params.column);
      const delimiter = String(step.params.delimiter ?? ",");
      const result = textSplit(table, col, delimiter);
      return { step, table: result.table!, message: result.message };
    }
    case "merge-columns": {
      const colA = Number(step.params.columnA);
      const colB = Number(step.params.columnB);
      const sep = String(step.params.separator ?? " ");
      const result = concatColumns(table, colA, colB, sep);
      return { step, table: result.table!, message: result.message };
    }
    case "remove-duplicates": {
      const cols = String(step.params.columns ?? "").split(",").map((s) => Number(s)).filter((n) => Number.isInteger(n) && n >= 0);
      const result = removeDuplicatesByColumns(table, cols);
      return { step, table: result.table, message: `Removed ${result.removed} duplicate row(s). ${result.table.rows.length} rows remain.` };
    }
    case "remove-blank-rows": {
      const col = Number(step.params.column);
      const rows = table.rows.filter((row) => !isBlankOrError(row[col] ?? ""));
      return { step, table: { ...table, rows }, message: `Removed ${table.rows.length - rows.length} blank/error row(s). ${rows.length} rows remain.` };
    }
    case "filter-rows": {
      const criterion: Criterion = { col: Number(step.params.column), op: step.params.op as CriteriaOp, value: String(step.params.value ?? "") };
      const result = filterTableByCriterion(table, criterion);
      return { step, table: result.table!, message: result.message };
    }
    case "sort": {
      const col = Number(step.params.column);
      const dir = step.params.direction as "asc" | "desc";
      const result = sortTableByColumn(table, col, dir);
      return { step, table: result.table!, message: result.message };
    }
    case "append": {
      const other = lookupTable(String(step.params.source));
      if (!other) return { step, table, message: "Referenced source table was not found; step skipped." };
      const merged = mergeAppend([table, other]);
      return { step, table: merged, message: `Appended ${other.rows.length} rows from "${other.sourceName}". ${merged.rows.length} total rows.` };
    }
    case "join": {
      const other = lookupTable(String(step.params.source));
      if (!other) return { step, table, message: "Referenced source table was not found; step skipped." };
      const leftKey = Number(step.params.leftKey);
      const rightKey = Number(step.params.rightKey);
      const merged = mergeJoin(table, other, leftKey, rightKey);
      return { step, table: merged, message: `Joined with "${other.sourceName}" on ${table.headers[leftKey]} = ${other.headers[rightKey]}.` };
    }
    default:
      return { step, table, message: "Unknown step type; skipped." };
  }
}

export function runQueryPipeline(source: DataTable, steps: QueryStep[], lookupTable: (name: string) => DataTable | undefined): QueryStepOutcome[] {
  const outcomes: QueryStepOutcome[] = [];
  let current = source;
  steps.forEach((step) => {
    const outcome = runQueryStep(current, step, lookupTable);
    outcomes.push(outcome);
    current = outcome.table;
  });
  return outcomes;
}

export function buildTransformationSummary(sourceName: string, outcomes: QueryStepOutcome[]): string {
  const lines = [
    "MarqClean AI — Power Query-style Transformation Summary",
    `Generated: ${new Date().toISOString()}`,
    `Source: ${sourceName}`,
    "",
    "This documents every step applied, in order, and its result. It is a transformation log, not Power Query \"M\" code — it will not appear as a live, refreshable query inside Excel.",
    "",
  ];
  outcomes.forEach((o, i) => {
    lines.push(`${i + 1}. ${o.step.label}`);
    lines.push(`   Result: ${o.message}`);
    lines.push(`   Output: ${o.table.rows.length} rows x ${o.table.headers.length} columns`);
    lines.push("");
  });
  return lines.join("\n");
}

export function describeStep(step: QueryStep): string {
  return step.label;
}
