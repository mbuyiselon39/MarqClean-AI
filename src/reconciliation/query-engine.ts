import { DataTable } from "./engine";

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

export type DataTypeTarget = "text" | "number" | "date" | "boolean" | "currency" | "percent";

export interface QueryStep {
  id: string;
  type: QueryStepType;
  label: string;
  params: Record<string, any>;
}

export interface StepOutcome {
  step: QueryStep;
  table: DataTable;
  rowsCount: number;
  colsCount: number;
  durationMs: number;
}

export function runQueryPipeline(
  initialTable: DataTable,
  steps: QueryStep[],
  lookupTable?: (name: string) => DataTable | undefined
): StepOutcome[] {
  const outcomes: StepOutcome[] = [];
  let current: DataTable = {
    name: initialTable.name,
    headers: [...initialTable.headers],
    rows: initialTable.rows.map((r) => ({ ...r })),
  };

  for (const step of steps) {
    const start = performance.now();
    let nextTable: DataTable = {
      name: current.name,
      headers: [...current.headers],
      rows: current.rows.map((r) => ({ ...r })),
    };

    switch (step.type) {
      case "choose-columns": {
        const rawCols = String(step.params.columns || "");
        const indices = rawCols
          .split(",")
          .map((n) => parseInt(n.trim(), 10))
          .filter((i) => !isNaN(i) && i >= 0 && i < current.headers.length);
        const keptHeaders = indices.map((i) => current.headers[i]);
        nextTable.headers = keptHeaders;
        nextTable.rows = current.rows.map((r) => {
          const newR: Record<string, string> = {};
          keptHeaders.forEach((h) => {
            newR[h] = r[h] || "";
          });
          return newR;
        });
        break;
      }
      case "change-type": {
        const colIdx = Number(step.params.column);
        const target = step.params.target as DataTypeTarget;
        const colName = current.headers[colIdx];
        if (colName) {
          nextTable.rows = current.rows.map((r) => {
            let val = (r[colName] || "").trim();
            if (target === "number") {
              const num = parseFloat(val.replace(/[^0-9.-]/g, ""));
              val = isNaN(num) ? "" : String(num);
            } else if (target === "currency") {
              const num = parseFloat(val.replace(/[^0-9.-]/g, ""));
              val = isNaN(num) ? "" : `$${num.toFixed(2)}`;
            } else if (target === "percent") {
              const num = parseFloat(val.replace(/[^0-9.-]/g, ""));
              val = isNaN(num) ? "" : `${(num * 100).toFixed(1)}%`;
            } else if (target === "boolean") {
              const lower = val.toLowerCase();
              val = lower === "true" || lower === "1" || lower === "yes" ? "TRUE" : "FALSE";
            }
            return { ...r, [colName]: val };
          });
        }
        break;
      }
      case "split-column": {
        const colIdx = Number(step.params.column);
        const delimiter = String(step.params.delimiter || ",");
        const colName = current.headers[colIdx];
        if (colName) {
          let maxSplits = 1;
          const rowSplits = current.rows.map((r) => {
            const parts = (r[colName] || "").split(delimiter).map((p) => p.trim());
            if (parts.length > maxSplits) maxSplits = parts.length;
            return { r, parts };
          });
          const newColHeaders = Array.from({ length: maxSplits }, (_, i) => `${colName}_${i + 1}`);
          nextTable.headers = [...current.headers, ...newColHeaders];
          nextTable.rows = rowSplits.map(({ r, parts }) => {
            const newR = { ...r };
            newColHeaders.forEach((nh, i) => {
              newR[nh] = parts[i] || "";
            });
            return newR;
          });
        }
        break;
      }
      case "merge-columns": {
        const colA = current.headers[Number(step.params.columnA)];
        const colB = current.headers[Number(step.params.columnB)];
        const sep = String(step.params.separator || " ");
        if (colA && colB) {
          const mergedName = `${colA}_${colB}`;
          nextTable.headers = [...current.headers, mergedName];
          nextTable.rows = current.rows.map((r) => ({
            ...r,
            [mergedName]: `${r[colA] || ""}${sep}${r[colB] || ""}`,
          }));
        }
        break;
      }
      case "remove-duplicates": {
        const rawCols = String(step.params.columns || "");
        const indices = rawCols
          .split(",")
          .map((n) => parseInt(n.trim(), 10))
          .filter((i) => !isNaN(i) && i >= 0 && i < current.headers.length);
        const keys = indices.length > 0 ? indices.map((i) => current.headers[i]) : current.headers;
        const seen = new Set<string>();
        nextTable.rows = current.rows.filter((r) => {
          const sig = keys.map((k) => (r[k] || "").trim().toLowerCase()).join("||");
          if (seen.has(sig)) return false;
          seen.add(sig);
          return true;
        });
        break;
      }
      case "remove-blank-rows": {
        const colName = current.headers[Number(step.params.column)];
        if (colName) {
          nextTable.rows = current.rows.filter((r) => (r[colName] || "").trim() !== "");
        }
        break;
      }
      case "filter-rows": {
        const colName = current.headers[Number(step.params.column)];
        const op = step.params.op;
        const val = String(step.params.value || "").toLowerCase();
        if (colName) {
          nextTable.rows = current.rows.filter((r) => {
            const cell = (r[colName] || "").toLowerCase();
            if (op === "=") return cell === val;
            if (op === "!=") return cell !== val;
            if (op === "contains") return cell.includes(val);
            const nCell = parseFloat(cell);
            const nVal = parseFloat(val);
            if (!isNaN(nCell) && !isNaN(nVal)) {
              if (op === ">") return nCell > nVal;
              if (op === "<") return nCell < nVal;
              if (op === ">=") return nCell >= nVal;
              if (op === "<=") return nCell <= nVal;
            }
            return true;
          });
        }
        break;
      }
      case "sort": {
        const colName = current.headers[Number(step.params.column)];
        const dir = step.params.direction === "desc" ? -1 : 1;
        if (colName) {
          nextTable.rows = [...current.rows].sort((a, b) => {
            const vA = (a[colName] || "").trim();
            const vB = (b[colName] || "").trim();
            const nA = parseFloat(vA);
            const nB = parseFloat(vB);
            if (!isNaN(nA) && !isNaN(nB)) return (nA - nB) * dir;
            return vA.localeCompare(vB) * dir;
          });
        }
        break;
      }
      case "append": {
        const otherName = String(step.params.source || "");
        const other = lookupTable ? lookupTable(otherName) : undefined;
        if (other) {
          const unionHeaders = Array.from(new Set([...current.headers, ...other.headers]));
          nextTable.headers = unionHeaders;
          const rows1 = current.rows.map((r) => {
            const newR: Record<string, string> = {};
            unionHeaders.forEach((h) => (newR[h] = r[h] || ""));
            return newR;
          });
          const rows2 = other.rows.map((r) => {
            const newR: Record<string, string> = {};
            unionHeaders.forEach((h) => (newR[h] = r[h] || ""));
            return newR;
          });
          nextTable.rows = [...rows1, ...rows2];
        }
        break;
      }
      case "join": {
        const otherName = String(step.params.source || "");
        const other = lookupTable ? lookupTable(otherName) : undefined;
        const leftKey = current.headers[Number(step.params.leftKey)];
        const rightKey = other ? other.headers[Number(step.params.rightKey)] : undefined;
        if (other && leftKey && rightKey) {
          const rightHeaders = other.headers.filter((h) => h !== rightKey);
          nextTable.headers = [...current.headers, ...rightHeaders.map((h) => `${otherName}_${h}`)];
          const lookup = new Map<string, Record<string, string>>();
          other.rows.forEach((r) => {
            const key = (r[rightKey] || "").trim().toLowerCase();
            if (!lookup.has(key)) lookup.set(key, r);
          });
          nextTable.rows = current.rows.map((r) => {
            const key = (r[leftKey] || "").trim().toLowerCase();
            const match = lookup.get(key);
            const combined = { ...r };
            rightHeaders.forEach((h) => {
              combined[`${otherName}_${h}`] = match ? match[h] || "" : "";
            });
            return combined;
          });
        }
        break;
      }
    }

    const duration = Math.round(performance.now() - start);
    outcomes.push({
      step,
      table: nextTable,
      rowsCount: nextTable.rows.length,
      colsCount: nextTable.headers.length,
      durationMs: duration,
    });
    current = nextTable;
  }

  return outcomes;
}

export function buildTransformationSummary(primaryName: string, outcomes: StepOutcome[]): string {
  if (outcomes.length === 0) return `Transformation summary for ${primaryName}: No steps applied.`;
  const lines = [`Power Query Pipeline Summary for: ${primaryName}`, `Executed ${outcomes.length} applied steps.`, ""];
  outcomes.forEach((o, i) => {
    lines.push(`Step ${i + 1}: ${o.step.label}`);
    lines.push(`  Type: ${o.step.type} | Duration: ${o.durationMs}ms`);
    lines.push(`  Resulting shape: ${o.rowsCount} rows × ${o.colsCount} columns`);
    lines.push("");
  });
  return lines.join("\n");
}
