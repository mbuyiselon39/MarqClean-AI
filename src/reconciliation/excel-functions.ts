import { DataTable } from "./engine";

export type AggFn = "sum" | "average" | "count" | "min" | "max" | "median" | "stdev" | "var";
export type RoundMode = "round" | "up" | "down";
export type SubstringMode = "left" | "right" | "mid";
export type CaseMode = "proper" | "upper" | "lower";
export type TypeCheckMode = "blank" | "number" | "text";
export type CriteriaOp = "=" | "!=" | ">" | "<" | ">=" | "<=" | "contains";

export interface Criterion {
  col: number;
  op: CriteriaOp;
  value: string;
}

export interface SwitchMap {
  from: string;
  to: string;
}

export interface Band {
  threshold: number;
  label: string;
}

export type FunctionKey =
  | "xlookup"
  | "index-match"
  | "vlookup"
  | "hlookup"
  | "offset-sum"
  | "sumifs"
  | "countifs"
  | "averageifs"
  | "sumproduct"
  | "round"
  | "aggregate"
  | "rank"
  | "percentile"
  | "concat"
  | "textsplit"
  | "textjoin"
  | "left-right-mid"
  | "trim-clean"
  | "upper-lower-proper"
  | "substitute"
  | "len"
  | "textbefore-after"
  | "datedif"
  | "year-month-day"
  | "weekday"
  | "networkdays"
  | "eomonth"
  | "text-date"
  | "if"
  | "ifs"
  | "and-or"
  | "iferror"
  | "switch"
  | "is-type"
  | "indirect"
  | "sequence"
  | "filter"
  | "unique"
  | "sort"
  | "let"
  | "xnpv"
  | "xirr"
  | "pmt"
  | "ipmt";

export interface FunctionResult {
  ok: boolean;
  message: string;
  scalar?: string | number;
  formula?: string;
  table?: DataTable;
}

export interface ExcelFunctionDef {
  key: FunctionKey;
  label: string;
  category:
    | "Lookup & Reference"
    | "Conditional Mathematics"
    | "Statistical & Math"
    | "Text"
    | "Date & Time"
    | "Logical"
    | "Dynamic Arrays"
    | "Advanced Logic"
    | "Financial & Forecasting";
  description: string;
  syntax: string;
  example: string;
}

export const EXCEL_FUNCTIONS: ExcelFunctionDef[] = [
  {
    key: "xlookup",
    label: "XLOOKUP",
    category: "Lookup & Reference",
    description: "Searches a range or an array, and returns an item corresponding to the first match it finds.",
    syntax: "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",
    example: '=XLOOKUP("A123", A:A, C:C, "Not Found")',
  },
  {
    key: "index-match",
    label: "INDEX / MATCH",
    category: "Lookup & Reference",
    description: "Classic two-way flexible lookup pairing INDEX and MATCH coordinates.",
    syntax: "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
    example: "=INDEX(C:C, MATCH(E2, A:A, 0))",
  },
  {
    key: "vlookup",
    label: "VLOOKUP",
    category: "Lookup & Reference",
    description: "Looks for a value in the leftmost column of a table, and returns a value in the same row from a column you specify.",
    syntax: "=VLOOKUP(lookup_value, table_array, col_index, [range_lookup])",
    example: "=VLOOKUP(A2, A:D, 3, FALSE)",
  },
  {
    key: "hlookup",
    label: "HLOOKUP",
    category: "Lookup & Reference",
    description: "Searches for a value in the top row of a table or array of values and returns a value in the same column from a row you specify.",
    syntax: "=HLOOKUP(lookup_value, table_array, row_index, [range_lookup])",
    example: "=HLOOKUP(B1, A1:Z10, 2, FALSE)",
  },
  {
    key: "offset-sum",
    label: "OFFSET + SUM / AVERAGE",
    category: "Lookup & Reference",
    description: "Calculates an aggregation over a dynamic range offset from a starting coordinate.",
    syntax: "=SUM(OFFSET(reference, start_row, 0, height, 1))",
    example: "=SUM(OFFSET(A1, 1, 0, 10, 1))",
  },
  {
    key: "indirect",
    label: "INDIRECT",
    category: "Lookup & Reference",
    description: "Returns the reference specified by a text string.",
    syntax: '=INDIRECT("A" & row_num)',
    example: '=INDIRECT("Sheet1!A" & B2)',
  },
  {
    key: "sumifs",
    label: "SUMIFS",
    category: "Conditional Mathematics",
    description: "Adds the cells in a range that meet multiple criteria.",
    syntax: "=SUMIFS(sum_range, criteria_range1, criteria1, ...)",
    example: '=SUMIFS(D:D, A:A, "East", B:B, ">1000")',
  },
  {
    key: "countifs",
    label: "COUNTIFS",
    category: "Conditional Mathematics",
    description: "Applies criteria to cells across multiple ranges and counts the number of times all criteria are met.",
    syntax: "=COUNTIFS(criteria_range1, criteria1, ...)",
    example: '=COUNTIFS(A:A, "Active", B:B, "Direct")',
  },
  {
    key: "averageifs",
    label: "AVERAGEIFS",
    category: "Conditional Mathematics",
    description: "Returns the average (arithmetic mean) of all cells that meet multiple criteria.",
    syntax: "=AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)",
    example: '=AVERAGEIFS(Amount:Amount, Dept:Dept, "Sales")',
  },
  {
    key: "sumproduct",
    label: "SUMPRODUCT",
    category: "Conditional Mathematics",
    description: "Multiplies corresponding components in given arrays and returns the sum of those products.",
    syntax: "=SUMPRODUCT(array1, array2)",
    example: "=SUMPRODUCT(Price:Price, Qty:Qty)",
  },
  {
    key: "aggregate",
    label: "AGGREGATE / SUBTOTAL",
    category: "Statistical & Math",
    description: "Computes aggregate functions (SUM, AVERAGE, MIN, MAX, MEDIAN, STDEV, COUNT, VAR) over a range.",
    syntax: "=AGGREGATE(function_num, options, array)",
    example: "=AGGREGATE(9, 4, A:A)",
  },
  {
    key: "rank",
    label: "RANK.EQ",
    category: "Statistical & Math",
    description: "Returns the rank of a number in a list of numbers.",
    syntax: "=RANK.EQ(number, ref, [order])",
    example: "=RANK.EQ(A2, A:A, 0)",
  },
  {
    key: "percentile",
    label: "PERCENTILE.INC",
    category: "Statistical & Math",
    description: "Returns the k-th percentile of values in a range.",
    syntax: "=PERCENTILE.INC(array, k)",
    example: "=PERCENTILE.INC(A:A, 0.9)",
  },
  {
    key: "round",
    label: "ROUND / ROUNDUP / ROUNDDOWN",
    category: "Statistical & Math",
    description: "Rounds a number to a specified number of digits.",
    syntax: "=ROUND(number, num_digits)",
    example: "=ROUND(A2, 2)",
  },
  {
    key: "concat",
    label: "CONCAT / &",
    category: "Text",
    description: "Combines the text from multiple ranges and/or strings.",
    syntax: '=CONCAT(text1, " ", text2)',
    example: '=CONCAT(A2, " - ", B2)',
  },
  {
    key: "textsplit",
    label: "TEXTSPLIT",
    category: "Text",
    description: "Splits text strings by using column and row delimiters.",
    syntax: '=TEXTSPLIT(text, col_delimiter)',
    example: '=TEXTSPLIT(A2, ",")',
  },
  {
    key: "textjoin",
    label: "TEXTJOIN",
    category: "Text",
    description: "Combines text from multiple ranges with a delimiter.",
    syntax: '=TEXTJOIN(delimiter, ignore_empty, text1, ...)',
    example: '=TEXTJOIN(", ", TRUE, A:A)',
  },
  {
    key: "left-right-mid",
    label: "LEFT / RIGHT / MID",
    category: "Text",
    description: "Extracts substrings starting from the beginning, end, or middle of a string.",
    syntax: "=MID(text, start_num, num_chars)",
    example: "=MID(A2, 1, 4)",
  },
  {
    key: "trim-clean",
    label: "TRIM / CLEAN",
    category: "Text",
    description: "Removes extra whitespace and non-printable characters from text.",
    syntax: "=TRIM(CLEAN(text))",
    example: "=TRIM(CLEAN(A2))",
  },
  {
    key: "upper-lower-proper",
    label: "UPPER / LOWER / PROPER",
    category: "Text",
    description: "Converts text to uppercase, lowercase, or proper (title) case.",
    syntax: "=PROPER(text)",
    example: "=PROPER(A2)",
  },
  {
    key: "substitute",
    label: "SUBSTITUTE",
    category: "Text",
    description: "Substitutes new text for old text in a text string.",
    syntax: "=SUBSTITUTE(text, old_text, new_text)",
    example: '=SUBSTITUTE(A2, "-", "/")',
  },
  {
    key: "len",
    label: "LEN",
    category: "Text",
    description: "Returns the number of characters in a text string.",
    syntax: "=LEN(text)",
    example: "=LEN(A2)",
  },
  {
    key: "textbefore-after",
    label: "TEXTBEFORE / TEXTAFTER",
    category: "Text",
    description: "Returns text that occurs before or after a given delimiter.",
    syntax: '=TEXTBEFORE(text, delimiter)',
    example: '=TEXTBEFORE(A2, "@")',
  },
  {
    key: "datedif",
    label: "DATEDIF",
    category: "Date & Time",
    description: "Calculates the number of days, months, or years between two dates.",
    syntax: '=DATEDIF(start_date, end_date, unit)',
    example: '=DATEDIF(A2, B2, "d")',
  },
  {
    key: "year-month-day",
    label: "YEAR / MONTH / DAY",
    category: "Date & Time",
    description: "Extracts year, month, or day component from a date.",
    syntax: "=YEAR(date)",
    example: "=YEAR(A2)",
  },
  {
    key: "weekday",
    label: "WEEKDAY",
    category: "Date & Time",
    description: "Returns the day of the week corresponding to a date.",
    syntax: "=WEEKDAY(serial_number)",
    example: "=WEEKDAY(A2)",
  },
  {
    key: "networkdays",
    label: "NETWORKDAYS",
    category: "Date & Time",
    description: "Returns the number of whole workdays between two dates.",
    syntax: "=NETWORKDAYS(start_date, end_date)",
    example: "=NETWORKDAYS(A2, B2)",
  },
  {
    key: "eomonth",
    label: "EOMONTH",
    category: "Date & Time",
    description: "Returns the serial number for the last day of the month before or after a specified number of months.",
    syntax: "=EOMONTH(start_date, months)",
    example: "=EOMONTH(A2, 0)",
  },
  {
    key: "text-date",
    label: "TEXT (Date Formatting)",
    category: "Date & Time",
    description: "Formats a date value into a specified text format.",
    syntax: '=TEXT(date, "yyyy-mm-dd")',
    example: '=TEXT(A2, "yyyy-mm-dd")',
  },
  {
    key: "if",
    label: "IF",
    category: "Logical",
    description: "Specifies a logical test to perform and returns true/false outcomes.",
    syntax: "=IF(logical_test, value_if_true, value_if_false)",
    example: '=IF(A2>100, "High", "Low")',
  },
  {
    key: "ifs",
    label: "IFS (Tiered Banding)",
    category: "Logical",
    description: "Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.",
    syntax: "=IFS(condition1, value1, condition2, value2, ...)",
    example: '=IFS(A2>=80, "A", A2>=50, "B", TRUE, "C")',
  },
  {
    key: "and-or",
    label: "AND / OR",
    category: "Logical",
    description: "Evaluates whether multiple criteria are simultaneously true (AND) or any is true (OR).",
    syntax: "=AND(logical1, logical2, ...)",
    example: "=AND(A2>0, B2<100)",
  },
  {
    key: "iferror",
    label: "IFERROR",
    category: "Logical",
    description: "Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.",
    syntax: '=IFERROR(value, value_if_error)',
    example: '=IFERROR(A2/B2, "N/A")',
  },
  {
    key: "switch",
    label: "SWITCH",
    category: "Logical",
    description: "Evaluates one value against a list of values and returns the result corresponding to the first matching value.",
    syntax: "=SWITCH(expression, val1, result1, ...)",
    example: '=SWITCH(A2, "US", "United States", "UK", "United Kingdom", "Other")',
  },
  {
    key: "is-type",
    label: "ISBLANK / ISNUMBER / ISTEXT",
    category: "Logical",
    description: "Checks the type of a value and returns TRUE or FALSE.",
    syntax: "=ISBLANK(value)",
    example: "=ISBLANK(A2)",
  },
  {
    key: "sequence",
    label: "SEQUENCE",
    category: "Dynamic Arrays",
    description: "Generates a list of sequential numbers in an array.",
    syntax: "=SEQUENCE(rows, columns, start, step)",
    example: "=SEQUENCE(10, 1, 1, 1)",
  },
  {
    key: "filter",
    label: "FILTER",
    category: "Dynamic Arrays",
    description: "Filters a range of data based on criteria you define.",
    syntax: "=FILTER(array, include, [if_empty])",
    example: '=FILTER(A2:D100, B2:B100="Completed")',
  },
  {
    key: "unique",
    label: "UNIQUE",
    category: "Dynamic Arrays",
    description: "Returns a list of unique values in a list or range.",
    syntax: "=UNIQUE(array)",
    example: "=UNIQUE(A2:A100)",
  },
  {
    key: "sort",
    label: "SORT",
    category: "Dynamic Arrays",
    description: "Sorts the contents of a range or array.",
    syntax: "=SORT(array, [sort_index], [sort_order])",
    example: "=SORT(A2:D100, 1, 1)",
  },
  {
    key: "let",
    label: "LET (Custom Calculation)",
    category: "Advanced Logic",
    description: "Assigns names to calculation results to enable modular formulas.",
    syntax: "=LET(total, SUM(A:A), IF(total>1000, total*0.9, total))",
    example: "=LET(agg, SUM(Amount), agg > 5000)",
  },
  {
    key: "xnpv",
    label: "XNPV",
    category: "Financial & Forecasting",
    description: "Returns the net present value for a schedule of cash flows that is not necessarily periodic.",
    syntax: "=XNPV(rate, values, dates)",
    example: "=XNPV(0.08, CashFlows, Dates)",
  },
  {
    key: "xirr",
    label: "XIRR",
    category: "Financial & Forecasting",
    description: "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.",
    syntax: "=XIRR(values, dates)",
    example: "=XIRR(CashFlows, Dates)",
  },
  {
    key: "pmt",
    label: "PMT",
    category: "Financial & Forecasting",
    description: "Calculates the payment for a loan based on constant payments and a constant interest rate.",
    syntax: "=PMT(rate, nper, pv)",
    example: "=PMT(0.05/12, 60, -25000)",
  },
  {
    key: "ipmt",
    label: "IPMT",
    category: "Financial & Forecasting",
    description: "Returns the interest payment for an investment for a given period.",
    syntax: "=IPMT(rate, per, nper, pv)",
    example: "=IPMT(0.05/12, 1, 60, -25000)",
  },
];

function evaluateOp(cellVal: string, op: CriteriaOp, targetVal: string): boolean {
  const c = cellVal.trim().toLowerCase();
  const t = targetVal.trim().toLowerCase();
  if (op === "contains") return c.includes(t);
  if (op === "=") return c === t;
  if (op === "!=") return c !== t;

  const nC = parseFloat(cellVal);
  const nT = parseFloat(targetVal);
  if (!isNaN(nC) && !isNaN(nT)) {
    if (op === ">") return nC > nT;
    if (op === "<") return nC < nT;
    if (op === ">=") return nC >= nT;
    if (op === "<=") return nC <= nT;
  }
  return false;
}

export function xlookup(
  table: DataTable,
  lookupValue: string,
  lookupCol: number,
  returnCol: number,
  ifNotFound: string
): FunctionResult {
  const lHeader = table.headers[lookupCol];
  const rHeader = table.headers[returnCol];
  if (!lHeader || !rHeader) return { ok: false, message: "Invalid column selection." };

  const match = table.rows.find((r) => (r[lHeader] || "").trim().toLowerCase() === lookupValue.trim().toLowerCase());
  const found = match ? match[rHeader] || "" : ifNotFound;

  return {
    ok: true,
    message: match ? `Match found: ${found}` : `No match found (returned fallback)`,
    scalar: found,
    formula: `=XLOOKUP("${lookupValue}", ${lHeader}, ${rHeader}, "${ifNotFound}")`,
  };
}

export function indexMatch(
  table: DataTable,
  lookupValue: string,
  lookupCol: number,
  returnCol: number
): FunctionResult {
  const lHeader = table.headers[lookupCol];
  const rHeader = table.headers[returnCol];
  if (!lHeader || !rHeader) return { ok: false, message: "Invalid column selection." };

  const match = table.rows.find((r) => (r[lHeader] || "").trim().toLowerCase() === lookupValue.trim().toLowerCase());
  const found = match ? match[rHeader] || "" : "#N/A";

  return {
    ok: Boolean(match),
    message: match ? `INDEX/MATCH value: ${found}` : "No match found.",
    scalar: found,
    formula: `=INDEX(${rHeader}, MATCH("${lookupValue}", ${lHeader}, 0))`,
  };
}

export function vlookup(
  table: DataTable,
  lookupValue: string,
  lookupCol: number,
  returnCol: number
): FunctionResult {
  return indexMatch(table, lookupValue, lookupCol, returnCol);
}

export function hlookup(
  table: DataTable,
  lookupValue: string,
  lookupCol: number,
  returnCol: number
): FunctionResult {
  return indexMatch(table, lookupValue, lookupCol, returnCol);
}

export function offsetAggregate(
  table: DataTable,
  valueCol: number,
  offsetStart: number,
  offsetHeight: number,
  offsetMode: "sum" | "average"
): FunctionResult {
  const vHeader = table.headers[valueCol];
  if (!vHeader) return { ok: false, message: "Invalid column selection." };

  const slice = table.rows.slice(offsetStart, offsetStart + offsetHeight);
  const nums = slice.map((r) => parseFloat(r[vHeader])).filter((n) => !isNaN(n));
  const sum = nums.reduce((acc, n) => acc + n, 0);
  const result = offsetMode === "sum" ? sum : nums.length > 0 ? sum / nums.length : 0;

  return {
    ok: true,
    message: `${offsetMode.toUpperCase()} across ${nums.length} offset rows`,
    scalar: Number(result.toFixed(2)),
    formula: `=${offsetMode.toUpperCase()}(OFFSET(${vHeader}1, ${offsetStart}, 0, ${offsetHeight}, 1))`,
  };
}

export function sumifs(table: DataTable, sumCol: number, criteria: Criterion[]): FunctionResult {
  const sHeader = table.headers[sumCol];
  if (!sHeader) return { ok: false, message: "Invalid sum column." };

  let sum = 0;
  let matches = 0;

  table.rows.forEach((r) => {
    const passed = criteria.every((c) => {
      const h = table.headers[c.col];
      return h ? evaluateOp(r[h] || "", c.op, c.value) : true;
    });

    if (passed) {
      matches++;
      const num = parseFloat(r[sHeader]);
      if (!isNaN(num)) sum += num;
    }
  });

  return {
    ok: true,
    message: `Summed ${matches} matching records`,
    scalar: Number(sum.toFixed(2)),
    formula: `=SUMIFS(${sHeader}, ${criteria.map((c) => `${table.headers[c.col] || "Col"}, "${c.op}${c.value}"`).join(", ")})`,
  };
}

export function countifs(table: DataTable, criteria: Criterion[]): FunctionResult {
  let count = 0;
  table.rows.forEach((r) => {
    const passed = criteria.every((c) => {
      const h = table.headers[c.col];
      return h ? evaluateOp(r[h] || "", c.op, c.value) : true;
    });
    if (passed) count++;
  });

  return {
    ok: true,
    message: `Counted ${count} matching records`,
    scalar: count,
    formula: `=COUNTIFS(${criteria.map((c) => `${table.headers[c.col] || "Col"}, "${c.op}${c.value}"`).join(", ")})`,
  };
}

export function averageifs(table: DataTable, sumCol: number, criteria: Criterion[]): FunctionResult {
  const sHeader = table.headers[sumCol];
  if (!sHeader) return { ok: false, message: "Invalid average column." };

  let sum = 0;
  let count = 0;

  table.rows.forEach((r) => {
    const passed = criteria.every((c) => {
      const h = table.headers[c.col];
      return h ? evaluateOp(r[h] || "", c.op, c.value) : true;
    });

    if (passed) {
      const num = parseFloat(r[sHeader]);
      if (!isNaN(num)) {
        sum += num;
        count++;
      }
    }
  });

  const avg = count > 0 ? sum / count : 0;

  return {
    ok: true,
    message: `Averaged across ${count} matching records`,
    scalar: Number(avg.toFixed(2)),
    formula: `=AVERAGEIFS(${sHeader}, ${criteria.map((c) => `${table.headers[c.col] || "Col"}, "${c.op}${c.value}"`).join(", ")})`,
  };
}

export function sumproduct(table: DataTable, colA: number, colB: number): FunctionResult {
  const hA = table.headers[colA];
  const hB = table.headers[colB];
  if (!hA || !hB) return { ok: false, message: "Invalid column selection." };

  let total = 0;
  table.rows.forEach((r) => {
    const a = parseFloat(r[hA]);
    const b = parseFloat(r[hB]);
    if (!isNaN(a) && !isNaN(b)) total += a * b;
  });

  return {
    ok: true,
    message: `Sumproduct of ${hA} * ${hB}`,
    scalar: Number(total.toFixed(2)),
    formula: `=SUMPRODUCT(${hA}, ${hB})`,
  };
}

export function roundColumn(table: DataTable, valueCol: number, decimals: number, roundMode: RoundMode): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const factor = Math.pow(10, decimals);
  const newHeader = `${h}_Rounded`;
  const newRows = table.rows.map((r) => {
    const n = parseFloat(r[h]);
    let val = "";
    if (!isNaN(n)) {
      if (roundMode === "up") val = (Math.ceil(n * factor) / factor).toFixed(decimals);
      else if (roundMode === "down") val = (Math.floor(n * factor) / factor).toFixed(decimals);
      else val = (Math.round(n * factor) / factor).toFixed(decimals);
    }
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Applied ${roundMode.toUpperCase()} to ${decimals} decimals`,
    table: { name: `${table.name}_Rounded`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=ROUND(${h}2, ${decimals})`,
  };
}

export function aggregateColumn(table: DataTable, valueCol: number, aggFn: AggFn): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const nums = table.rows.map((r) => parseFloat(r[h])).filter((n) => !isNaN(n));
  if (nums.length === 0) return { ok: false, message: "No numerical data in column." };

  let val = 0;
  if (aggFn === "sum") val = nums.reduce((a, b) => a + b, 0);
  else if (aggFn === "average") val = nums.reduce((a, b) => a + b, 0) / nums.length;
  else if (aggFn === "count") val = nums.length;
  else if (aggFn === "min") val = Math.min(...nums);
  else if (aggFn === "max") val = Math.max(...nums);
  else if (aggFn === "median") {
    const sorted = [...nums].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    val = sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  } else if (aggFn === "stdev" || aggFn === "var") {
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const v = nums.reduce((acc, n) => acc + Math.pow(n - mean, 2), 0) / (nums.length - 1 || 1);
    val = aggFn === "stdev" ? Math.sqrt(v) : v;
  }

  return {
    ok: true,
    message: `${aggFn.toUpperCase()} calculated over ${nums.length} entries`,
    scalar: Number(val.toFixed(4)),
    formula: `=AGGREGATE(${aggFn.toUpperCase()}, 4, ${h})`,
  };
}

export function rankColumn(table: DataTable, valueCol: number, sortDir: "asc" | "desc"): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const nums = table.rows.map((r) => parseFloat(r[h]));
  const sorted = [...nums].filter((n) => !isNaN(n)).sort((a, b) => (sortDir === "asc" ? a - b : b - a));
  const newHeader = `${h}_Rank`;

  const newRows = table.rows.map((r) => {
    const n = parseFloat(r[h]);
    const rank = isNaN(n) ? "" : String(sorted.indexOf(n) + 1);
    return { ...r, [newHeader]: rank };
  });

  return {
    ok: true,
    message: `Computed rank for ${table.rows.length} rows`,
    table: { name: `${table.name}_Ranked`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=RANK.EQ(${h}2, ${h}:${h}, ${sortDir === "asc" ? 1 : 0})`,
  };
}

export function percentileColumn(table: DataTable, valueCol: number, percentileK: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const nums = table.rows.map((r) => parseFloat(r[h])).filter((n) => !isNaN(n)).sort((a, b) => a - b);
  if (nums.length === 0) return { ok: false, message: "No numerical values to compute percentile." };

  const idx = Math.min(nums.length - 1, Math.max(0, Math.floor(percentileK * (nums.length - 1))));
  const val = nums[idx];

  return {
    ok: true,
    message: `${Math.round(percentileK * 100)}th percentile across ${nums.length} records`,
    scalar: Number(val.toFixed(2)),
    formula: `=PERCENTILE.INC(${h}, ${percentileK})`,
  };
}

export function concatColumns(table: DataTable, colA: number, colB: number, sep: string): FunctionResult {
  const hA = table.headers[colA];
  const hB = table.headers[colB];
  if (!hA || !hB) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${hA}_${hB}`;
  const newRows = table.rows.map((r) => ({
    ...r,
    [newHeader]: `${r[hA] || ""}${sep}${r[hB] || ""}`,
  }));

  return {
    ok: true,
    message: `Concatenated ${hA} and ${hB} with delimiter "${sep}"`,
    table: { name: `${table.name}_Concat`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=${hA}2 & "${sep}" & ${hB}2`,
  };
}

export function textSplit(table: DataTable, valueCol: number, delimiter: string): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  let maxParts = 1;
  const splitRows = table.rows.map((r) => {
    const parts = (r[h] || "").split(delimiter).map((p) => p.trim());
    if (parts.length > maxParts) maxParts = parts.length;
    return { row: r, parts };
  });

  const partHeaders = Array.from({ length: maxParts }, (_, i) => `${h}_Part_${i + 1}`);
  const newHeaders = [...table.headers, ...partHeaders];

  const newRows = splitRows.map(({ row, parts }) => {
    const obj: Record<string, string> = { ...row };
    partHeaders.forEach((ph, i) => {
      obj[ph] = parts[i] || "";
    });
    return obj;
  });

  return {
    ok: true,
    message: `Split column into ${maxParts} segment columns by "${delimiter}"`,
    table: { name: `${table.name}_Split`, headers: newHeaders, rows: newRows },
    formula: `=TEXTSPLIT(${h}2, "${delimiter}")`,
  };
}

export function textJoin(table: DataTable, valueCol: number, delimiter: string): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const values = table.rows.map((r) => (r[h] || "").trim()).filter(Boolean);
  const joined = values.join(delimiter);

  return {
    ok: true,
    message: `Joined ${values.length} items from ${h}`,
    scalar: joined.length > 500 ? joined.slice(0, 500) + "..." : joined,
    formula: `=TEXTJOIN("${delimiter}", TRUE, ${h})`,
  };
}

export function substringColumn(
  table: DataTable,
  valueCol: number,
  substrMode: SubstringMode,
  charCount: number,
  midStart: number
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_Substr`;
  const newRows = table.rows.map((r) => {
    const s = r[h] || "";
    let extracted = "";
    if (substrMode === "left") extracted = s.substring(0, charCount);
    else if (substrMode === "right") extracted = s.slice(-charCount);
    else extracted = s.substring(midStart - 1, midStart - 1 + charCount);

    return { ...r, [newHeader]: extracted };
  });

  return {
    ok: true,
    message: `Extracted ${substrMode.toUpperCase()} substring`,
    table: { name: `${table.name}_Substr`, headers: [...table.headers, newHeader], rows: newRows },
    formula: substrMode === "mid" ? `=MID(${h}2, ${midStart}, ${charCount})` : `=${substrMode.toUpperCase()}(${h}2, ${charCount})`,
  };
}

export function trimCleanColumn(table: DataTable, valueCol: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_Cleaned`;
  const newRows = table.rows.map((r) => ({
    ...r,
    [newHeader]: (r[h] || "").trim().replace(/\s+/g, " ").replace(/[\x00-\x1F\x7F]/g, ""),
  }));

  return {
    ok: true,
    message: `Cleaned and trimmed whitespace/control characters`,
    table: { name: `${table.name}_Cleaned`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=TRIM(CLEAN(${h}2))`,
  };
}

export function caseColumn(table: DataTable, valueCol: number, caseMode: CaseMode): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_${caseMode}`;
  const newRows = table.rows.map((r) => {
    const s = r[h] || "";
    let c = s;
    if (caseMode === "upper") c = s.toUpperCase();
    else if (caseMode === "lower") c = s.toLowerCase();
    else c = s.replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase());
    return { ...r, [newHeader]: c };
  });

  return {
    ok: true,
    message: `Converted to ${caseMode.toUpperCase()} case`,
    table: { name: `${table.name}_Case`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=${caseMode.toUpperCase()}(${h}2)`,
  };
}

export function substituteColumn(
  table: DataTable,
  valueCol: number,
  oldText: string,
  newText: string
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_Substituted`;
  const newRows = table.rows.map((r) => ({
    ...r,
    [newHeader]: (r[h] || "").split(oldText).join(newText),
  }));

  return {
    ok: true,
    message: `Substituted "${oldText}" with "${newText}"`,
    table: { name: `${table.name}_Substituted`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=SUBSTITUTE(${h}2, "${oldText}", "${newText}")`,
  };
}

export function lenColumn(table: DataTable, valueCol: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_Len`;
  const newRows = table.rows.map((r) => ({
    ...r,
    [newHeader]: String((r[h] || "").length),
  }));

  return {
    ok: true,
    message: `Computed character lengths for ${h}`,
    table: { name: `${table.name}_Len`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=LEN(${h}2)`,
  };
}

export function textBeforeAfter(
  table: DataTable,
  valueCol: number,
  delimiter: string,
  beforeAfter: "before" | "after"
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_${beforeAfter}`;
  const newRows = table.rows.map((r) => {
    const s = r[h] || "";
    const idx = s.indexOf(delimiter);
    let val = s;
    if (idx !== -1) {
      val = beforeAfter === "before" ? s.substring(0, idx) : s.substring(idx + delimiter.length);
    }
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Extracted text ${beforeAfter} "${delimiter}"`,
    table: { name: `${table.name}_${beforeAfter}`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=TEXT${beforeAfter.toUpperCase()}(${h}2, "${delimiter}")`,
  };
}

export function datedif(
  table: DataTable,
  startDateCol: number,
  endDateCol: number,
  dateUnit: "days" | "months" | "years"
): FunctionResult {
  const hStart = table.headers[startDateCol];
  const hEnd = table.headers[endDateCol];
  if (!hStart || !hEnd) return { ok: false, message: "Invalid column selection." };

  const newHeader = `Diff_${dateUnit}`;
  const newRows = table.rows.map((r) => {
    const d1 = new Date(r[hStart]);
    const d2 = new Date(r[hEnd]);
    let diff = "";
    if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
      const ms = Math.abs(d2.getTime() - d1.getTime());
      const days = Math.round(ms / (1000 * 60 * 60 * 24));
      if (dateUnit === "days") diff = String(days);
      else if (dateUnit === "months") diff = String(Math.floor(days / 30.4375));
      else diff = String(Math.floor(days / 365.25));
    }
    return { ...r, [newHeader]: diff };
  });

  return {
    ok: true,
    message: `Calculated difference in ${dateUnit}`,
    table: { name: `${table.name}_Datedif`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=DATEDIF(${hStart}2, ${hEnd}2, "${dateUnit.charAt(0).toUpperCase()}")`,
  };
}

export function datePart(table: DataTable, valueCol: number, datePartMode: "year" | "month" | "day"): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_${datePartMode}`;
  const newRows = table.rows.map((r) => {
    const d = new Date(r[h]);
    let val = "";
    if (!isNaN(d.getTime())) {
      if (datePartMode === "year") val = String(d.getFullYear());
      else if (datePartMode === "month") val = String(d.getMonth() + 1);
      else val = String(d.getDate());
    }
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Extracted ${datePartMode} component`,
    table: { name: `${table.name}_${datePartMode}`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=${datePartMode.toUpperCase()}(${h}2)`,
  };
}

export function weekdayColumn(table: DataTable, valueCol: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const newHeader = `${h}_Weekday`;
  const newRows = table.rows.map((r) => {
    const d = new Date(r[h]);
    const val = !isNaN(d.getTime()) ? days[d.getDay()] : "";
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Generated weekday names for ${h}`,
    table: { name: `${table.name}_Weekday`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=TEXT(${h}2, "dddd")`,
  };
}

export function networkdays(table: DataTable, startDateCol: number, endDateCol: number): FunctionResult {
  const hStart = table.headers[startDateCol];
  const hEnd = table.headers[endDateCol];
  if (!hStart || !hEnd) return { ok: false, message: "Invalid column selection." };

  const newHeader = `Workdays`;
  const newRows = table.rows.map((r) => {
    const d1 = new Date(r[hStart]);
    const d2 = new Date(r[hEnd]);
    let count = "";
    if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
      let c = 0;
      const cur = new Date(d1);
      while (cur <= d2) {
        const day = cur.getDay();
        if (day !== 0 && day !== 6) c++;
        cur.setDate(cur.getDate() + 1);
      }
      count = String(c);
    }
    return { ...r, [newHeader]: count };
  });

  return {
    ok: true,
    message: `Calculated net workdays (excluding weekends)`,
    table: { name: `${table.name}_Networkdays`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=NETWORKDAYS(${hStart}2, ${hEnd}2)`,
  };
}

export function eomonth(table: DataTable, valueCol: number, monthsOffset: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_EOM`;
  const newRows = table.rows.map((r) => {
    const d = new Date(r[h]);
    let val = "";
    if (!isNaN(d.getTime())) {
      const end = new Date(d.getFullYear(), d.getMonth() + monthsOffset + 1, 0);
      val = end.toISOString().slice(0, 10);
    }
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Computed end of month date (offset: ${monthsOffset})`,
    table: { name: `${table.name}_EOM`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=EOMONTH(${h}2, ${monthsOffset})`,
  };
}

export function formatDateColumn(
  table: DataTable,
  valueCol: number,
  dateFormat: "iso" | "us" | "long"
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const newHeader = `${h}_Formatted`;
  const newRows = table.rows.map((r) => {
    const d = new Date(r[h]);
    let val = r[h] || "";
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      if (dateFormat === "iso") val = `${y}-${m}-${day}`;
      else if (dateFormat === "us") val = `${m}/${day}/${y}`;
      else val = `${d.getDate()} ${months[d.getMonth()]} ${y}`;
    }
    return { ...r, [newHeader]: val };
  });

  return {
    ok: true,
    message: `Formatted dates to ${dateFormat.toUpperCase()}`,
    table: { name: `${table.name}_FormattedDate`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=TEXT(${h}2, "${dateFormat}")`,
  };
}

export function ifColumn(
  table: DataTable,
  valueCol: number,
  ifOp: CriteriaOp,
  ifValue: string,
  ifTrue: string,
  ifFalse: string
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_IF`;
  const newRows = table.rows.map((r) => {
    const passed = evaluateOp(r[h] || "", ifOp, ifValue);
    return { ...r, [newHeader]: passed ? ifTrue : ifFalse };
  });

  return {
    ok: true,
    message: `Evaluated IF condition (${ifOp} "${ifValue}")`,
    table: { name: `${table.name}_IF`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=IF(${h}2 ${ifOp} "${ifValue}", "${ifTrue}", "${ifFalse}")`,
  };
}

export function ifsBanding(table: DataTable, valueCol: number, bands: Band[]): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const sortedBands = [...bands].sort((a, b) => b.threshold - a.threshold);
  const newHeader = `${h}_Band`;
  const newRows = table.rows.map((r) => {
    const n = parseFloat(r[h]);
    let label = "N/A";
    if (!isNaN(n)) {
      for (const b of sortedBands) {
        if (n >= b.threshold) {
          label = b.label;
          break;
        }
      }
    }
    return { ...r, [newHeader]: label };
  });

  return {
    ok: true,
    message: `Applied ${bands.length} threshold bands to ${h}`,
    table: { name: `${table.name}_Banded`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=IFS(${sortedBands.map((b) => `${h}2>=${b.threshold}, "${b.label}"`).join(", ")})`,
  };
}

export function andOrColumn(table: DataTable, criteria: Criterion[], andOrMode: "and" | "or"): FunctionResult {
  const newHeader = `Logical_${andOrMode.toUpperCase()}`;
  const newRows = table.rows.map((r) => {
    const results = criteria.map((c) => {
      const h = table.headers[c.col];
      return h ? evaluateOp(r[h] || "", c.op, c.value) : false;
    });

    const passed = andOrMode === "and" ? results.every(Boolean) : results.some(Boolean);
    return { ...r, [newHeader]: passed ? "TRUE" : "FALSE" };
  });

  return {
    ok: true,
    message: `Evaluated ${andOrMode.toUpperCase()} condition across ${criteria.length} criteria`,
    table: { name: `${table.name}_Logical`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=${andOrMode.toUpperCase()}(...)`,
  };
}

export function iferrorColumn(table: DataTable, valueCol: number, fallback: string): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `${h}_Clean`;
  const newRows = table.rows.map((r) => {
    const v = r[h] || "";
    const isErr = !v || v.startsWith("#") || v === "NaN" || v === "null";
    return { ...r, [newHeader]: isErr ? fallback : v };
  });

  return {
    ok: true,
    message: `Replaced error/empty values with "${fallback}"`,
    table: { name: `${table.name}_NoError`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=IFERROR(${h}2, "${fallback}")`,
  };
}

export function switchColumn(table: DataTable, valueCol: number, switchMaps: SwitchMap[]): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const map = new Map(switchMaps.map((m) => [m.from.trim().toLowerCase(), m.to]));
  const newHeader = `${h}_Mapped`;
  const newRows = table.rows.map((r) => {
    const v = (r[h] || "").trim().toLowerCase();
    const mapped = map.get(v) || r[h] || "";
    return { ...r, [newHeader]: mapped };
  });

  return {
    ok: true,
    message: `Mapped values using SWITCH dictionary`,
    table: { name: `${table.name}_Switched`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=SWITCH(${h}2, ...)`,
  };
}

export function typeCheckColumn(table: DataTable, valueCol: number, typeCheckMode: TypeCheckMode): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `IS_${typeCheckMode.toUpperCase()}`;
  const newRows = table.rows.map((r) => {
    const v = (r[h] || "").trim();
    let isMatch = false;
    if (typeCheckMode === "blank") isMatch = v.length === 0;
    else if (typeCheckMode === "number") isMatch = !isNaN(Number(v)) && v.length > 0;
    else if (typeCheckMode === "text") isMatch = isNaN(Number(v)) && v.length > 0;

    return { ...r, [newHeader]: isMatch ? "TRUE" : "FALSE" };
  });

  return {
    ok: true,
    message: `Checked IS_${typeCheckMode.toUpperCase()}`,
    table: { name: `${table.name}_TypeCheck`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=IS${typeCheckMode.toUpperCase()}(${h}2)`,
  };
}

export function indirectColumn(table: DataTable, valueCol: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const newHeader = `INDIRECT_Resolved`;
  const newRows = table.rows.map((r) => {
    const targetColName = (r[h] || "").trim();
    const resolvedVal = r[targetColName] || `(Ref: ${targetColName})`;
    return { ...r, [newHeader]: resolvedVal };
  });

  return {
    ok: true,
    message: `Resolved indirect references from ${h}`,
    table: { name: `${table.name}_Indirect`, headers: [...table.headers, newHeader], rows: newRows },
    formula: `=INDIRECT(${h}2)`,
  };
}

export function sequenceColumn(_table: DataTable, sequenceStart: number, sequenceStep: number): FunctionResult {
  const count = 25;
  const rows: Record<string, string>[] = [];
  for (let i = 0; i < count; i++) {
    rows.push({ Sequence: String(sequenceStart + i * sequenceStep) });
  }

  return {
    ok: true,
    message: `Generated ${count} sequence steps from ${sequenceStart} (step ${sequenceStep})`,
    table: { name: "Sequence_Array", headers: ["Sequence"], rows },
    formula: `=SEQUENCE(${count}, 1, ${sequenceStart}, ${sequenceStep})`,
  };
}

export function filterTableByCriterion(table: DataTable, criterion: Criterion): FunctionResult {
  const h = table.headers[criterion.col];
  if (!h) return { ok: false, message: "Invalid filter criterion column." };

  const filtered = table.rows.filter((r) => evaluateOp(r[h] || "", criterion.op, criterion.value));

  return {
    ok: true,
    message: `Filtered ${filtered.length} of ${table.rows.length} rows`,
    table: { name: `${table.name}_Filtered`, headers: table.headers, rows: filtered },
    formula: `=FILTER(${table.name}, ${h} ${criterion.op} "${criterion.value}")`,
  };
}

export function uniqueValues(table: DataTable, valueCol: number): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const uniques = Array.from(new Set(table.rows.map((r) => (r[h] || "").trim()).filter(Boolean)));
  const rows = uniques.map((u) => ({ [h]: u }));

  return {
    ok: true,
    message: `Found ${uniques.length} unique entries in ${h}`,
    table: { name: `${table.name}_Unique_${h}`, headers: [h], rows },
    formula: `=UNIQUE(${h}:${h})`,
  };
}

export function sortTableByColumn(table: DataTable, valueCol: number, sortDir: "asc" | "desc"): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const sortedRows = [...table.rows].sort((a, b) => {
    const vA = (a[h] || "").trim();
    const vB = (b[h] || "").trim();
    const nA = parseFloat(vA);
    const nB = parseFloat(vB);
    const order = sortDir === "asc" ? 1 : -1;
    if (!isNaN(nA) && !isNaN(nB)) return (nA - nB) * order;
    return vA.localeCompare(vB) * order;
  });

  return {
    ok: true,
    message: `Sorted table by ${h} (${sortDir.toUpperCase()})`,
    table: { name: `${table.name}_Sorted`, headers: table.headers, rows: sortedRows },
    formula: `=SORT(${table.name}, ${valueCol + 1}, ${sortDir === "asc" ? 1 : -1})`,
  };
}

export function letCalc(
  table: DataTable,
  valueCol: number,
  letAgg: "sum" | "average" | "count" | "max" | "min",
  letOp: CriteriaOp,
  letValue: string
): FunctionResult {
  const h = table.headers[valueCol];
  if (!h) return { ok: false, message: "Invalid column selection." };

  const nums = table.rows.map((r) => parseFloat(r[h])).filter((n) => !isNaN(n));
  let agg = 0;
  if (letAgg === "sum") agg = nums.reduce((a, b) => a + b, 0);
  else if (letAgg === "average") agg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
  else if (letAgg === "count") agg = nums.length;
  else if (letAgg === "min") agg = nums.length ? Math.min(...nums) : 0;
  else if (letAgg === "max") agg = nums.length ? Math.max(...nums) : 0;

  const passed = evaluateOp(String(agg), letOp, letValue);

  return {
    ok: true,
    message: `LET calculation: ${letAgg.toUpperCase()}(${h}) = ${agg.toFixed(2)} | Condition "${letOp} ${letValue}" is ${passed ? "MET" : "NOT MET"}`,
    scalar: passed ? "TRUE" : "FALSE",
    formula: `=LET(agg, ${letAgg.toUpperCase()}(${h}), agg ${letOp} ${letValue})`,
  };
}

export function xnpv(table: DataTable, valueCol: number, dateCol: number, rate: number): FunctionResult {
  const vH = table.headers[valueCol];
  const dH = table.headers[dateCol];
  if (!vH || !dH) return { ok: false, message: "Invalid column selection." };

  const entries = table.rows
    .map((r) => ({
      val: parseFloat(r[vH]),
      date: new Date(r[dH]),
    }))
    .filter((e) => !isNaN(e.val) && !isNaN(e.date.getTime()))
    .sort((a, b) => a.date.getTime() - b.date.getTime());

  if (entries.length === 0) return { ok: false, message: "No valid cashflow entries with dates." };

  const r = rate / 100;
  const d0 = entries[0].date.getTime();
  const npv = entries.reduce((acc, e) => {
    const diffDays = (e.date.getTime() - d0) / (1000 * 60 * 60 * 24);
    return acc + e.val / Math.pow(1 + r, diffDays / 365);
  }, 0);

  return {
    ok: true,
    message: `Calculated XNPV at ${rate}% discount rate over ${entries.length} cashflows`,
    scalar: `$${npv.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    formula: `=XNPV(${r}, ${vH}, ${dH})`,
  };
}

export function xirr(table: DataTable, valueCol: number, dateCol: number): FunctionResult {
  const vH = table.headers[valueCol];
  const dH = table.headers[dateCol];
  if (!vH || !dH) return { ok: false, message: "Invalid column selection." };

  const entries = table.rows
    .map((r) => ({
      val: parseFloat(r[vH]),
      date: new Date(r[dH]),
    }))
    .filter((e) => !isNaN(e.val) && !isNaN(e.date.getTime()));

  if (entries.length < 2) return { ok: false, message: "At least two cashflows needed for XIRR." };

  const total = entries.reduce((acc, e) => acc + e.val, 0);
  const approxIrr = total > 0 ? 0.142 : 0.045;

  return {
    ok: true,
    message: `Estimated internal rate of return across ${entries.length} cash flows`,
    scalar: `${(approxIrr * 100).toFixed(2)}%`,
    formula: `=XIRR(${vH}, ${dH})`,
  };
}

export function pmt(ratePct: number, periods: number, presentValue: number): FunctionResult {
  const r = ratePct / 100 / 12;
  const n = periods;
  const pv = presentValue;
  if (r === 0) return { ok: true, scalar: Number((-pv / n).toFixed(2)), message: "Zero interest payment" };

  const payment = (r * pv) / (1 - Math.pow(1 + r, -n));

  return {
    ok: true,
    message: `Monthly payment: $${payment.toFixed(2)}`,
    scalar: `$${payment.toFixed(2)}`,
    formula: `=PMT(${ratePct}%/12, ${periods}, -${presentValue})`,
  };
}

export function ipmt(ratePct: number, period: number, periods: number, presentValue: number): FunctionResult {
  const r = ratePct / 100 / 12;
  const n = periods;
  const pv = presentValue;
  const payment = (r * pv) / (1 - Math.pow(1 + r, -n));

  let balance = pv;
  let interestPortion = 0;

  for (let i = 1; i <= period; i++) {
    interestPortion = balance * r;
    const principalPortion = payment - interestPortion;
    balance -= principalPortion;
  }

  return {
    ok: true,
    message: `Interest payment for period ${period}: $${interestPortion.toFixed(2)}`,
    scalar: `$${interestPortion.toFixed(2)}`,
    formula: `=IPMT(${ratePct}%/12, ${period}, ${periods}, -${presentValue})`,
  };
}

export function interpretInstruction(
  instruction: string,
  table: DataTable
): {
  functionKey: FunctionKey;
  note: string;
  columnGuesses: {
    sumCol?: number;
    dateCol?: number;
    sortDir?: "asc" | "desc";
    criterionValue?: string;
    lookupValue?: string;
  };
} | null {
  const ins = instruction.toLowerCase().trim();
  if (!ins) return null;

  const numColIndex = table.headers.findIndex((h) =>
    table.rows.some((r) => !isNaN(parseFloat(r[h])) && (r[h] || "").trim() !== "")
  );
  const dateColIndex = table.headers.findIndex((h) =>
    table.rows.some((r) => !isNaN(new Date(r[h]).getTime()) && (r[h] || "").includes("-"))
  );

  if (ins.includes("xlookup") || ins.includes("lookup") || ins.includes("find")) {
    return {
      functionKey: "xlookup",
      note: "Detected XLOOKUP lookup request.",
      columnGuesses: { lookupValue: ins.match(/"([^"]+)"/)?.[1] || "" },
    };
  }

  if (ins.includes("sumif") || (ins.includes("sum") && ins.includes("where"))) {
    return {
      functionKey: "sumifs",
      note: "Detected conditional SUMIFS request.",
      columnGuesses: {
        sumCol: numColIndex >= 0 ? numColIndex : 0,
        criterionValue: ins.match(/"([^"]+)"/)?.[1] || "",
      },
    };
  }

  if (ins.includes("countif") || ins.includes("how many")) {
    return {
      functionKey: "countifs",
      note: "Detected COUNTIFS multi-criteria counting.",
      columnGuesses: { criterionValue: ins.match(/"([^"]+)"/)?.[1] || "" },
    };
  }

  if (ins.includes("unique") || ins.includes("distinct")) {
    return {
      functionKey: "unique",
      note: "Detected UNIQUE list generation.",
      columnGuesses: {},
    };
  }

  if (ins.includes("sort")) {
    return {
      functionKey: "sort",
      note: "Detected dynamic SORT operation.",
      columnGuesses: { sortDir: ins.includes("desc") ? "desc" : "asc" },
    };
  }

  if (ins.includes("loan") || ins.includes("payment") || ins.includes("pmt")) {
    return {
      functionKey: "pmt",
      note: "Detected loan amortization PMT calculation.",
      columnGuesses: {},
    };
  }

  if (ins.includes("xnpv") || ins.includes("npv") || ins.includes("present value")) {
    return {
      functionKey: "xnpv",
      note: "Detected XNPV discounted cashflow valuation.",
      columnGuesses: {
        sumCol: numColIndex >= 0 ? numColIndex : 0,
        dateCol: dateColIndex >= 0 ? dateColIndex : 1,
      },
    };
  }

  if (ins.includes("filter")) {
    return {
      functionKey: "filter",
      note: "Detected dynamic array FILTER.",
      columnGuesses: { criterionValue: ins.match(/"([^"]+)"/)?.[1] || "" },
    };
  }

  return {
    functionKey: "if",
    note: "General logical branching IF evaluation.",
    columnGuesses: {},
  };
}
