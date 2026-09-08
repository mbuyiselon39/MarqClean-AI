import type { DataTable } from "./engine";

// ---------------------------------------------------------------------------
// Power Pivot-style data modelling.
//
// Honest scope: this is NOT an in-memory columnar database (xVelocity) and
// does not support dynamic, arbitrary filter-context re-aggregation the way
// real Power Pivot measures do inside Excel. What it does provide, genuinely:
//   - Relationship metadata between uploaded tables, with real validation
//     against the actual data (not a cosmetic checkbox).
//   - A simple relationship diagram rendered from that metadata.
//   - Row-level calculated columns (delegated to the existing Advanced Excel
//     Functions engine — no new evaluation logic needed).
//   - A fixed set of cross-table aggregate "measures" (sum/average/count/
//     min/max of a column in the "many" table, grouped by the matching key
//     in the "one" table), computed once when the model is built rather than
//     dynamically re-evaluated against an arbitrary filter context.
// ---------------------------------------------------------------------------

export type Cardinality = "one-to-many" | "one-to-one";

export type Relationship = {
  id: string;
  oneTable: string;
  oneKey: number;
  manyTable: string;
  manyKey: number;
  cardinality: Cardinality;
};

export type RelationshipConflict = {
  relationshipId: string;
  severity: "error" | "warning";
  message: string;
};

function columnValues(table: DataTable, col: number): string[] {
  return table.rows.map((row) => (row[col] ?? "").trim().toLowerCase());
}

function originalCasedValues(table: DataTable, col: number): string[] {
  return table.rows.map((row) => (row[col] ?? "").trim());
}

function hasDuplicates(values: string[], originals: string[]): { duplicate: boolean; sampleKey?: string } {
  const seen = new Set<string>();
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v === "") continue;
    if (seen.has(v)) return { duplicate: true, sampleKey: originals[i] };
    seen.add(v);
  }
  return { duplicate: false };
}

/** Validates a declared relationship against the actual uploaded data. Returns real conflicts, not a cosmetic pass/fail. */
export function validateRelationship(tables: Map<string, DataTable>, rel: Relationship): RelationshipConflict[] {
  const conflicts: RelationshipConflict[] = [];
  const oneTable = tables.get(rel.oneTable);
  const manyTable = tables.get(rel.manyTable);

  if (!oneTable || !manyTable) {
    conflicts.push({ relationshipId: rel.id, severity: "error", message: "One or both tables in this relationship could not be found." });
    return conflicts;
  }

  const oneValues = columnValues(oneTable, rel.oneKey);
  const manyValues = columnValues(manyTable, rel.manyKey);
  const oneOriginals = originalCasedValues(oneTable, rel.oneKey);
  const manyOriginals = originalCasedValues(manyTable, rel.manyKey);
  const oneDupe = hasDuplicates(oneValues, oneOriginals);
  const manyDupe = hasDuplicates(manyValues, manyOriginals);

  if (oneDupe.duplicate) {
    conflicts.push({
      relationshipId: rel.id,
      severity: "error",
      message: `"${oneTable.headers[rel.oneKey]}" in ${rel.oneTable} has duplicate values (e.g. "${oneDupe.sampleKey}") but was declared as the "one" side of a ${rel.cardinality} relationship, which requires unique keys.`,
    });
  }

  if (rel.cardinality === "one-to-one" && manyDupe.duplicate) {
    conflicts.push({
      relationshipId: rel.id,
      severity: "error",
      message: `"${manyTable.headers[rel.manyKey]}" in ${rel.manyTable} has duplicate values (e.g. "${manyDupe.sampleKey}"), so this is not actually a one-to-one relationship — declare it as one-to-many instead.`,
    });
  }

  const oneSet = new Set(oneValues.filter(Boolean));
  const unmatchedIndex = manyValues.findIndex((v, i) => v && !oneSet.has(v) && manyOriginals[i]);
  const unmatchedCount = new Set(manyValues.filter((v) => v && !oneSet.has(v))).size;
  if (unmatchedCount > 0) {
    conflicts.push({
      relationshipId: rel.id,
      severity: "warning",
      message: `${unmatchedCount} distinct value(s) in ${rel.manyTable}."${manyTable.headers[rel.manyKey]}" do not match any row in ${rel.oneTable}."${oneTable.headers[rel.oneKey]}" (e.g. "${manyOriginals[unmatchedIndex]}").`,
    });
  }

  return conflicts;
}

export type MeasureAggFn = "sum" | "average" | "count" | "min" | "max";

export type Measure = {
  id: string;
  name: string;
  relationshipId: string;
  aggColumn: number;
  aggFn: MeasureAggFn;
};

function aggregate(values: number[], fn: MeasureAggFn): number {
  if (fn === "count") return values.length;
  if (values.length === 0) return 0;
  if (fn === "sum") return values.reduce((a, b) => a + b, 0);
  if (fn === "average") return values.reduce((a, b) => a + b, 0) / values.length;
  if (fn === "min") return Math.min(...values);
  return Math.max(...values);
}

/** Computes a cross-table measure once (not dynamically re-evaluated against an arbitrary filter context) and returns the "one" table with the measure appended as a new column. */
export function computeMeasure(tables: Map<string, DataTable>, relationships: Relationship[], measure: Measure): DataTable {
  const rel = relationships.find((r) => r.id === measure.relationshipId);
  if (!rel) throw new Error("The relationship this measure depends on no longer exists.");
  const oneTable = tables.get(rel.oneTable);
  const manyTable = tables.get(rel.manyTable);
  if (!oneTable || !manyTable) throw new Error("One or both tables in this measure's relationship could not be found.");

  const byKey = new Map<string, number[]>();
  manyTable.rows.forEach((row) => {
    const key = (row[rel.manyKey] ?? "").trim().toLowerCase();
    const raw = (row[measure.aggColumn] ?? "").replace(/[^0-9.\-]/g, "");
    const n = Number(raw);
    if (!Number.isFinite(n)) return;
    const list = byKey.get(key) ?? [];
    list.push(n);
    byKey.set(key, list);
  });

  const headers = [...oneTable.headers, measure.name];
  const rows = oneTable.rows.map((row) => {
    const key = (row[rel.oneKey] ?? "").trim().toLowerCase();
    const values = byKey.get(key) ?? [];
    return [...row, values.length > 0 || measure.aggFn === "count" ? String(aggregate(values, measure.aggFn)) : ""];
  });

  return { headers, rows, sourceName: oneTable.sourceName };
}
