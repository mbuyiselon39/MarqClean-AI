// ---------------------------------------------------------------------------
// DAX-equivalent formula reference.
//
// Honest scope: this is a reference and pattern library, NOT a DAX engine.
// Formulas are explained and mapped to an equivalent calculation this
// platform can actually run (via the Advanced Excel Functions engine or the
// Power Pivot-style measures above) — they are not parsed, executed, or
// validated as real DAX, and are not guaranteed to work if pasted into
// Power BI or Excel's Power Pivot without adjustment. Syntax checking below
// is shape-only (balanced parentheses, a recognized function name, argument
// count in range) — it is not semantic validation of row/filter context.
// ---------------------------------------------------------------------------

export type DaxCategory = "Aggregation" | "Filtering" | "Time Intelligence" | "Relationship" | "Financial";

export type DaxPattern = {
  name: string;
  category: DaxCategory;
  minArgs: number;
  maxArgs: number;
  daxSyntax: string;
  daxExample: string;
  equivalent: string;
};

export const DAX_PATTERNS: DaxPattern[] = [
  { name: "SUM", category: "Aggregation", minArgs: 1, maxArgs: 1, daxSyntax: "SUM(<column>)", daxExample: "SUM(Orders[Amount])", equivalent: "Use a Power Pivot-style measure with the Sum aggregation, or SUMIFS in Advanced Excel Functions for a single-table conditional total." },
  { name: "AVERAGE", category: "Aggregation", minArgs: 1, maxArgs: 1, daxSyntax: "AVERAGE(<column>)", daxExample: "AVERAGE(Orders[Amount])", equivalent: "Use a Power Pivot-style measure with the Average aggregation, or AVERAGEIFS in Advanced Excel Functions." },
  { name: "COUNT", category: "Aggregation", minArgs: 1, maxArgs: 1, daxSyntax: "COUNT(<column>)", daxExample: "COUNT(Orders[OrderID])", equivalent: "Use a Power Pivot-style measure with the Count aggregation, or COUNTIFS in Advanced Excel Functions." },
  { name: "DISTINCTCOUNT", category: "Aggregation", minArgs: 1, maxArgs: 1, daxSyntax: "DISTINCTCOUNT(<column>)", daxExample: "DISTINCTCOUNT(Orders[CustomerID])", equivalent: "Use the UNIQUE function in Advanced Excel Functions on the column, then count the resulting rows." },
  { name: "SUMX", category: "Aggregation", minArgs: 2, maxArgs: 2, daxSyntax: "SUMX(<table>, <expression>)", daxExample: "SUMX(Orders, Orders[Qty] * Orders[UnitPrice])", equivalent: "Use the Power Query-style workflow to add a calculated column (Qty * UnitPrice) as a step, then SUM that new column." },
  { name: "RANKX", category: "Aggregation", minArgs: 2, maxArgs: 3, daxSyntax: "RANKX(<table>, <expression>, [order])", daxExample: "RANKX(Orders, Orders[Amount])", equivalent: "Use the RANK function in Advanced Excel Functions on the relevant column." },
  { name: "CALCULATE", category: "Filtering", minArgs: 2, maxArgs: 6, daxSyntax: "CALCULATE(<expression>, <filter1>, [filter2], ...)", daxExample: "CALCULATE(SUM(Orders[Amount]), Orders[Region] = \"East\")", equivalent: "Use Power Query's Filter Rows step to isolate the condition first, then apply a Power Pivot-style measure or SUMIFS to the filtered result." },
  { name: "FILTER", category: "Filtering", minArgs: 2, maxArgs: 2, daxSyntax: "FILTER(<table>, <condition>)", daxExample: "FILTER(Orders, Orders[Amount] > 100)", equivalent: "Directly equivalent to the Power Query-style workflow's Filter Rows step." },
  { name: "ALL", category: "Filtering", minArgs: 1, maxArgs: 1, daxSyntax: "ALL(<table or column>)", daxExample: "ALL(Orders)", equivalent: "Removing filters mid-calculation has no direct equivalent here — recompute the measure from the unfiltered source table instead." },
  { name: "DIVIDE", category: "Filtering", minArgs: 2, maxArgs: 3, daxSyntax: "DIVIDE(<numerator>, <denominator>, [alternate])", daxExample: "DIVIDE(SUM(Orders[Profit]), SUM(Orders[Amount]), 0)", equivalent: "Use IFERROR around a division in Advanced Excel Functions to return a fallback value instead of a divide-by-zero error." },
  { name: "TOTALYTD", category: "Time Intelligence", minArgs: 2, maxArgs: 3, daxSyntax: "TOTALYTD(<expression>, <dates column>)", daxExample: "TOTALYTD(SUM(Orders[Amount]), Orders[Date])", equivalent: "Use Power Query's Filter Rows step (date >= start of year, date <= current date) then a Sum measure over the filtered rows." },
  { name: "DATESYTD", category: "Time Intelligence", minArgs: 1, maxArgs: 2, daxSyntax: "DATESYTD(<dates column>)", daxExample: "DATESYTD(Orders[Date])", equivalent: "Use the YEAR/MONTH/DAY date functions in Advanced Excel Functions to derive a year-to-date flag column, then filter on it." },
  { name: "SAMEPERIODLASTYEAR", category: "Time Intelligence", minArgs: 1, maxArgs: 1, daxSyntax: "SAMEPERIODLASTYEAR(<dates column>)", daxExample: "SAMEPERIODLASTYEAR(Orders[Date])", equivalent: "Use DATEDIF/date-part functions to shift a date column back one year, then filter or join on that shifted value." },
  { name: "RELATED", category: "Relationship", minArgs: 1, maxArgs: 1, daxSyntax: "RELATED(<column in the one table>)", daxExample: "RELATED(Customers[Region])", equivalent: "Directly equivalent to a Power Query-style Join step, or a Power Pivot-style relationship pulling a column across." },
  { name: "RELATEDTABLE", category: "Relationship", minArgs: 1, maxArgs: 1, daxSyntax: "RELATEDTABLE(<many table>)", daxExample: "RELATEDTABLE(Orders)", equivalent: "Directly equivalent to a Power Pivot-style measure aggregating the many-side table, grouped by the one-side key." },
  { name: "USERELATIONSHIP", category: "Relationship", minArgs: 2, maxArgs: 2, daxSyntax: "USERELATIONSHIP(<col1>, <col2>)", daxExample: "USERELATIONSHIP(Orders[ShipDate], Calendar[Date])", equivalent: "Define a second, separate relationship between the same two tables on the alternate key columns, and build a separate measure from it." },
  { name: "XNPV", category: "Financial", minArgs: 3, maxArgs: 3, daxSyntax: "XNPV(<rate>, <values>, <dates>)", daxExample: "XNPV(0.1, CashFlows[Amount], CashFlows[Date])", equivalent: "Directly available as XNPV in Advanced Excel Functions." },
  { name: "XIRR", category: "Financial", minArgs: 2, maxArgs: 3, daxSyntax: "XIRR(<values>, <dates>, [guess])", daxExample: "XIRR(CashFlows[Amount], CashFlows[Date])", equivalent: "Directly available as XIRR in Advanced Excel Functions." },
  { name: "PMT", category: "Financial", minArgs: 3, maxArgs: 5, daxSyntax: "PMT(<rate>, <nper>, <pv>, [fv], [type])", daxExample: "PMT(0.05/12, 60, -20000)", equivalent: "Directly available as PMT in Advanced Excel Functions." },
];

export const DAX_CATEGORIES: DaxCategory[] = ["Aggregation", "Filtering", "Time Intelligence", "Relationship", "Financial"];

export type SyntaxCheckResult = {
  ok: boolean;
  issues: string[];
  matchedFunction: DaxPattern | null;
};

/** Shape-only syntax check: balanced parens, a recognized function name, and a plausible argument count. This does NOT evaluate the formula. */
export function checkDaxSyntaxShape(formula: string): SyntaxCheckResult {
  const issues: string[] = [];
  const trimmed = formula.trim();

  if (!trimmed) {
    return { ok: false, issues: ["Enter a formula to check."], matchedFunction: null };
  }

  let depth = 0;
  for (const ch of trimmed) {
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (depth < 0) {
      issues.push("Unmatched closing parenthesis.");
      break;
    }
  }
  if (depth > 0) issues.push(`${depth} unclosed opening parenthesis(es).`);

  const nameMatch = trimmed.match(/^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/);
  const functionName = nameMatch?.[1]?.toUpperCase();
  const matched = functionName ? DAX_PATTERNS.find((p) => p.name === functionName) ?? null : null;

  if (!functionName) {
    issues.push("Could not find a function name followed by an opening parenthesis, e.g. SUM(...).");
  } else if (!matched) {
    issues.push(`"${functionName}" is not in this platform's DAX pattern reference. It may still be valid DAX — only the functions listed in the reference below are checked here.`);
  } else if (depth === 0) {
    const openIdx = trimmed.indexOf("(");
    const closeIdx = trimmed.lastIndexOf(")");
    if (openIdx !== -1 && closeIdx > openIdx) {
      const inner = trimmed.slice(openIdx + 1, closeIdx);
      const argCount = inner.trim() === "" ? 0 : splitTopLevelArgs(inner).length;
      if (argCount < matched.minArgs || argCount > matched.maxArgs) {
        const range = matched.minArgs === matched.maxArgs ? `exactly ${matched.minArgs}` : `${matched.minArgs}-${matched.maxArgs}`;
        issues.push(`${functionName} expects ${range} argument(s); found ${argCount}.`);
      }
    }
  }

  return { ok: issues.length === 0, issues, matchedFunction: matched };
}

function splitTopLevelArgs(input: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let current = "";
  for (const ch of input) {
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (ch === "," && depth === 0) {
      args.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  if (current.trim() !== "") args.push(current);
  return args;
}
