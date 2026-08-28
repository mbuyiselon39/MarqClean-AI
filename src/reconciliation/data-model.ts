import { DataTable } from "./engine";

export type Cardinality = "one-to-many" | "many-to-one" | "one-to-one" | "many-to-many";
export type MeasureAggFn = "sum" | "average" | "count" | "min" | "max" | "distinctcount";

export interface Relationship {
  id: string;
  oneTable: string;
  oneKey: number;
  manyTable: string;
  manyKey: number;
  cardinality: Cardinality;
}

export interface RelationshipConflict {
  severity: "error" | "warning" | "info";
  message: string;
}

export interface Measure {
  id: string;
  name: string;
  relationshipId: string;
  aggColumn: number;
  aggFn: MeasureAggFn;
}

export function validateRelationship(
  tableMap: Map<string, DataTable>,
  rel: Relationship
): RelationshipConflict[] {
  const conflicts: RelationshipConflict[] = [];
  const tOne = tableMap.get(rel.oneTable);
  const tMany = tableMap.get(rel.manyTable);

  if (!tOne) {
    conflicts.push({ severity: "error", message: `Table "${rel.oneTable}" is missing.` });
    return conflicts;
  }
  if (!tMany) {
    conflicts.push({ severity: "error", message: `Table "${rel.manyTable}" is missing.` });
    return conflicts;
  }

  const kOne = tOne.headers[rel.oneKey];
  const kMany = tMany.headers[rel.manyKey];

  if (!kOne) conflicts.push({ severity: "error", message: `Key column index ${rel.oneKey} in "${rel.oneTable}" is invalid.` });
  if (!kMany) conflicts.push({ severity: "error", message: `Key column index ${rel.manyKey} in "${rel.manyTable}" is invalid.` });

  if (kOne && kMany) {
    // Check primary key uniqueness on "one" side
    const oneVals = tOne.rows.map((r) => (r[kOne] || "").trim().toLowerCase()).filter(Boolean);
    const uniqueOneVals = new Set(oneVals);
    if (uniqueOneVals.size < oneVals.length) {
      conflicts.push({
        severity: "error",
        message: `Duplicate keys found in "${rel.oneTable}.${kOne}". The "one" side of a relationship requires unique primary keys.`,
      });
    }

    // Check orphan records on "many" side
    const manyVals = tMany.rows.map((r) => (r[kMany] || "").trim().toLowerCase()).filter(Boolean);
    const orphans = manyVals.filter((v) => !uniqueOneVals.has(v));
    if (orphans.length > 0) {
      conflicts.push({
        severity: "warning",
        message: `${orphans.length} row(s) in "${rel.manyTable}.${kMany}" have foreign keys not found in "${rel.oneTable}.${kOne}".`,
      });
    }
  }

  return conflicts;
}

export function computeMeasure(
  tableMap: Map<string, DataTable>,
  relationships: Relationship[],
  measure: Measure
): DataTable {
  const rel = relationships.find((r) => r.id === measure.relationshipId);
  if (!rel) {
    throw new Error(`Relationship with ID "${measure.relationshipId}" not found.`);
  }

  const tOne = tableMap.get(rel.oneTable);
  const tMany = tableMap.get(rel.manyTable);

  if (!tOne || !tMany) {
    throw new Error("One or both tables in the relationship are missing.");
  }

  const kOne = tOne.headers[rel.oneKey];
  const kMany = tMany.headers[rel.manyKey];
  const aggCol = tMany.headers[measure.aggColumn];

  if (!kOne || !kMany || !aggCol) {
    throw new Error("Specified key or aggregation column does not exist.");
  }

  // Group rows in many-side by the foreign key
  const groups = new Map<string, string[]>();
  tMany.rows.forEach((r) => {
    const fk = (r[kMany] || "").trim().toLowerCase();
    if (!groups.has(fk)) groups.set(fk, []);
    groups.get(fk)!.push(r[aggCol] || "");
  });

  const newHeader = measure.name;
  const newRows = tOne.rows.map((r) => {
    const pk = (r[kOne] || "").trim().toLowerCase();
    const cellVals = groups.get(pk) || [];
    const nums = cellVals.map((v) => parseFloat(v)).filter((n) => !isNaN(n));

    let aggResult = "";
    if (measure.aggFn === "sum") {
      const sum = nums.reduce((a, b) => a + b, 0);
      aggResult = String(Number(sum.toFixed(2)));
    } else if (measure.aggFn === "average") {
      const avg = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
      aggResult = String(Number(avg.toFixed(2)));
    } else if (measure.aggFn === "count") {
      aggResult = String(cellVals.length);
    } else if (measure.aggFn === "distinctcount") {
      aggResult = String(new Set(cellVals.filter(Boolean)).size);
    } else if (measure.aggFn === "min") {
      aggResult = nums.length > 0 ? String(Math.min(...nums)) : "";
    } else if (measure.aggFn === "max") {
      aggResult = nums.length > 0 ? String(Math.max(...nums)) : "";
    }

    return { ...r, [newHeader]: aggResult };
  });

  return {
    name: tOne.name,
    headers: [...tOne.headers, newHeader],
    rows: newRows,
  };
}
