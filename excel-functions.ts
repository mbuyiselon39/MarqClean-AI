import type { DataTable } from "./engine";

// ---------------------------------------------------------------------------
// Advanced Excel function engine (client-side, no server, results in seconds)
// Operates on a loaded DataTable. Each action returns either a scalar result
// or a derived DataTable that can be previewed and exported.
// ---------------------------------------------------------------------------

export type FunctionCategory =
  | "Lookup & Reference"
  | "Conditional Mathematics"
  | "Statistical & Math"
  | "Text"
  | "Date & Time"
  | "Logical"
  | "Dynamic Arrays"
  | "Advanced Logic"
  | "Financial & Forecasting";

export type FunctionKey =
  | "xlookup"
  | "index-match"
  | "offset-sum"
  | "vlookup"
  | "hlookup"
  | "sumifs"
  | "countifs"
  | "averageifs"
  | "maxifs"
  | "minifs"
  | "correl"
  | "stdev-s"
  | "var-s"
  | "forecast-linear"
  | "forecast-ets"
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
  | "datevalue"
  | "timevalue"
  | "regexreplace"
  | "regextest"
  | "regexextract"
  | "formulatext"
  | "error-type"
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
  | "choosecols"
  | "vstack"
  | "tocol"
  | "sort"
  | "let"
  | "xnpv"
  | "npv"
  | "xirr"
  | "irr"
  | "pmt"
  | "ipmt";

export type FunctionDef = {
  key: FunctionKey;
  label: string;
  category: FunctionCategory;
  description: string;
  syntax: string;
};

export const EXCEL_FUNCTIONS: FunctionDef[] = [
  { key: "xlookup", label: "XLOOKUP", category: "Lookup & Reference", description: "Search a lookup column (left or right) and return a matching value, with an optional default if not found.", syntax: "XLOOKUP(lookup_value, lookup_column, return_column, [if_not_found])" },
  { key: "index-match", label: "INDEX & MATCH", category: "Lookup & Reference", description: "Look up a value anywhere in the table regardless of column position.", syntax: "INDEX(return_column, MATCH(lookup_value, lookup_column))" },
  { key: "offset-sum", label: "OFFSET + SUM/AVERAGE", category: "Lookup & Reference", description: "Sum or average a dynamic range offset by a number of rows from a starting row.", syntax: "SUM(OFFSET(start_row, rows, height) in column)" },
  { key: "vlookup", label: "VLOOKUP", category: "Lookup & Reference", description: "Look up a value in the first chosen column and return a value from a column to its right.", syntax: "VLOOKUP(lookup_value, lookup_column, return_column, FALSE)" },
  { key: "hlookup", label: "HLOOKUP (row match)", category: "Lookup & Reference", description: "Find the first row where a column equals a value and return another column from that row.", syntax: "HLOOKUP(lookup_value, lookup_column, return_column)" },
  { key: "sumifs", label: "SUMIFS", category: "Conditional Mathematics", description: "Sum a column where multiple criteria across other columns are met.", syntax: "SUMIFS(sum_column, criteria_col1, crit1, criteria_col2, crit2, ...)" },
  { key: "countifs", label: "COUNTIFS", category: "Conditional Mathematics", description: "Count rows where multiple criteria across columns are met.", syntax: "COUNTIFS(criteria_col1, crit1, criteria_col2, crit2, ...)" },
  { key: "averageifs", label: "AVERAGEIFS", category: "Conditional Mathematics", description: "Average a column where multiple criteria across columns are met.", syntax: "AVERAGEIFS(avg_column, criteria_col1, crit1, ...)" },
  { key: "maxifs", label: "MAXIFS", category: "Conditional Mathematics", description: "Return the largest value in a column among rows meeting all selected criteria.", syntax: "MAXIFS(max_column, criteria_col1, crit1, ...)" },
  { key: "minifs", label: "MINIFS", category: "Conditional Mathematics", description: "Return the smallest value in a column among rows meeting all selected criteria.", syntax: "MINIFS(min_column, criteria_col1, crit1, ...)" },
  { key: "correl", label: "CORREL", category: "Statistical & Math", description: "Return the Pearson correlation coefficient between two numeric columns.", syntax: "CORREL(array1, array2)" },
  { key: "stdev-s", label: "STDEV.S", category: "Statistical & Math", description: "Estimate sample standard deviation from a numeric column.", syntax: "STDEV.S(column)" },
  { key: "var-s", label: "VAR.S", category: "Statistical & Math", description: "Estimate sample variance from a numeric column.", syntax: "VAR.S(column)" },
  { key: "forecast-linear", label: "FORECAST.LINEAR", category: "Statistical & Math", description: "Project a future numeric value using linear regression.", syntax: "FORECAST.LINEAR(x, known_y, known_x)" },
  { key: "forecast-ets", label: "FORECAST.ETS", category: "Statistical & Math", description: "Forecast a future value with seasonal exponential smoothing.", syntax: "FORECAST.ETS(target, values, timeline, [seasonality])" },
  { key: "sumproduct", label: "SUMPRODUCT", category: "Conditional Mathematics", description: "Multiply two columns row by row and sum the products (weighted totals).", syntax: "SUMPRODUCT(column_a, column_b)" },
  { key: "round", label: "ROUND / ROUNDUP / ROUNDDOWN", category: "Statistical & Math", description: "Round every value in a numeric column to a set number of decimals.", syntax: "ROUND(column, decimals)" },
  { key: "aggregate", label: "SUM / AVERAGE / MIN / MAX / MEDIAN", category: "Statistical & Math", description: "Compute a single aggregate over a numeric column, including MEDIAN, STDEV, and VAR.", syntax: "AGGREGATE(function, column)" },
  { key: "rank", label: "RANK", category: "Statistical & Math", description: "Add a rank position for each row based on a numeric column.", syntax: "RANK(value, column, order)" },
  { key: "percentile", label: "PERCENTILE / QUARTILE", category: "Statistical & Math", description: "Return a percentile or quartile value from a numeric column.", syntax: "PERCENTILE(column, k)" },
  { key: "concat", label: "CONCAT / and", category: "Text", description: "Combine two columns into one new column, with an optional separator.", syntax: "CONCAT(column_a, separator, column_b)" },
  { key: "textsplit", label: "TEXTSPLIT", category: "Text", description: "Split a text column into multiple columns by a delimiter.", syntax: "TEXTSPLIT(column, delimiter)" },
  { key: "textjoin", label: "TEXTJOIN", category: "Text", description: "Join all values of a column into a single delimited string.", syntax: "TEXTJOIN(delimiter, column)" },
  { key: "left-right-mid", label: "LEFT / RIGHT / MID", category: "Text", description: "Extract characters from the start, end, or middle of a text column.", syntax: "LEFT(column, n) / RIGHT(column, n) / MID(column, start, n)" },
  { key: "trim-clean", label: "TRIM / CLEAN", category: "Text", description: "Remove extra spaces and non-printable characters from a text column.", syntax: "TRIM(CLEAN(column))" },
  { key: "upper-lower-proper", label: "UPPER / LOWER / PROPER", category: "Text", description: "Change the case of a text column.", syntax: "PROPER(column)" },
  { key: "substitute", label: "SUBSTITUTE / REPLACE", category: "Text", description: "Replace occurrences of text within a column.", syntax: "SUBSTITUTE(column, old, new)" },
  { key: "regexreplace", label: "REGEXREPLACE", category: "Text", description: "Replace text matching a regular expression.", syntax: "REGEXREPLACE(text, pattern, replacement)" },
  { key: "regextest", label: "REGEXTEST", category: "Text", description: "Test whether text matches a regular expression.", syntax: "REGEXTEST(text, pattern)" },
  { key: "regexextract", label: "REGEXEXTRACT", category: "Text", description: "Extract the first text match for a regular expression.", syntax: "REGEXEXTRACT(text, pattern)" },
  { key: "len", label: "LEN", category: "Text", description: "Return the character length of every value in a text column.", syntax: "LEN(column)" },
  { key: "textbefore-after", label: "TEXTBEFORE / TEXTAFTER", category: "Text", description: "Return the text before or after a delimiter in a column.", syntax: "TEXTBEFORE(column, delimiter) / TEXTAFTER(column, delimiter)" },
  { key: "datedif", label: "DATEDIF", category: "Date & Time", description: "Difference between two date columns in days, months, or years.", syntax: "DATEDIF(start_date, end_date, unit)" },
  { key: "year-month-day", label: "YEAR / MONTH / DAY", category: "Date & Time", description: "Extract the year, month, or day from a date column.", syntax: "YEAR(column) / MONTH(column) / DAY(column)" },
  { key: "weekday", label: "WEEKDAY / TEXT weekday", category: "Date & Time", description: "Return the day of week name for a date column.", syntax: "TEXT(column, \"dddd\")" },
  { key: "networkdays", label: "NETWORKDAYS", category: "Date & Time", description: "Count working days (excluding weekends) between two date columns.", syntax: "NETWORKDAYS(start_date, end_date)" },
  { key: "eomonth", label: "EOMONTH", category: "Date & Time", description: "Return the last day of the month a given number of months from each date.", syntax: "EOMONTH(column, months)" },
  { key: "text-date", label: "TEXT (format dates)", category: "Date & Time", description: "Format a date column into a chosen text pattern.", syntax: "TEXT(column, "yyyy-mm-dd")" },
  { key: "datevalue", label: "DATEVALUE", category: "Date & Time", description: "Convert a recognizable date string into an Excel-style serial date.", syntax: "DATEVALUE(date_text)" },
  { key: "timevalue", label: "TIMEVALUE", category: "Date & Time", description: "Convert a recognizable time string into an Excel-style fractional day.", syntax: "TIMEVALUE(time_text)" },
  { key: "if", label: "IF", category: "Logical", description: "Return one value when a condition is met and another when it is not, as a new column.", syntax: "IF(column op value, if_true, if_false)" },
  { key: "ifs", label: "IFS (banding)", category: "Logical", description: "Assign a label based on numeric thresholds, like grade or tier banding.", syntax: "IFS(value >= a, \"A\", value >= b, \"B\", ...)" },
  { key: "and-or", label: "AND / OR", category: "Logical", description: "Test whether all or any of several criteria are met, per row.", syntax: "AND(cond1, cond2) / OR(cond1, cond2)" },
  { key: "iferror", label: "IFERROR", category: "Logical", description: "Replace blank or error values in a column with a safe default.", syntax: "IFERROR(column, default)" },
  { key: "switch", label: "SWITCH", category: "Logical", description: "Map exact values in a column to new labels using a lookup list.", syntax: "SWITCH(column, val1, res1, val2, res2, ...)" },
  { key: "is-type", label: "ISBLANK / ISNUMBER / ISTEXT", category: "Logical", description: "Flag every row TRUE or FALSE based on whether a cell is blank, a number, or text, useful for validating a column before running other formulas.", syntax: "ISBLANK(column) / ISNUMBER(column) / ISTEXT(column)" },
  { key: "formulatext", label: "FORMULATEXT", category: "Logical", description: "Return the formula text stored in a cell when the imported value contains an Excel formula.", syntax: "FORMULATEXT(cell)" },
  { key: "error-type", label: "ERROR.TYPE", category: "Logical", description: "Return the Excel error code number for a recognized error value.", syntax: "ERROR.TYPE(value)" },
  { key: "indirect", label: "INDIRECT (dynamic column)", category: "Lookup & Reference", description: "For each row, use the text in a selector column as a column name and pull that row's value from the matching column, the same way INDIRECT resolves a text string into a live reference.", syntax: "INDIRECT(selector_column) -> value from the named column" },
  { key: "sequence", label: "SEQUENCE", category: "Dynamic Arrays", description: "Generate a column of sequential numbers, such as row IDs or a running count, starting at a chosen number with a chosen step.", syntax: "SEQUENCE(rows, [start], [step])" },
  { key: "filter", label: "FILTER", category: "Dynamic Arrays", description: "Return only the rows that meet a condition on a chosen column.", syntax: "FILTER(table, column operator value)" },
  { key: "unique", label: "UNIQUE", category: "Dynamic Arrays", description: "Return a list of unique values from a column (or unique rows).", syntax: "UNIQUE(column)" },
  { key: "choosecols", label: "CHOOSECOLS", category: "Dynamic Arrays", description: "Return only selected columns from the imported table.", syntax: "CHOOSECOLS(array, col_num1, col_num2, ...)" },
  { key: "vstack", label: "VSTACK", category: "Dynamic Arrays", description: "Append the rows of a second imported table below the first.", syntax: "VSTACK(array1, array2)" },
  { key: "tocol", label: "TOCOL", category: "Dynamic Arrays", description: "Flatten a table into one column.", syntax: "TOCOL(array, [ignore])" },
  { key: "sort", label: "SORT", category: "Dynamic Arrays", description: "Sort the table by a column, ascending or descending.", syntax: "SORT(table, by column, asc/desc)" },
  { key: "let", label: "LET (named calc)", category: "Advanced Logic", description: "Compute a named intermediate result and a final expression for readable calculations.", syntax: "LET(name = aggregate(column), name operator value)" },
  { key: "xnpv", label: "XNPV", category: "Financial & Forecasting", description: "Net present value for cash flows on irregular dates.", syntax: "XNPV(rate, values_column, dates_column)" },
  { key: "npv", label: "NPV", category: "Financial & Forecasting", description: "Net present value for periodic cash flows.", syntax: "NPV(rate, values_column)" },
  { key: "xirr", label: "XIRR", category: "Financial & Forecasting", description: "Internal rate of return for cash flows on irregular dates.", syntax: "XIRR(values_column, dates_column)" },
  { key: "irr", label: "IRR", category: "Financial & Forecasting", description: "Internal rate of return for periodic cash flows.", syntax: "IRR(values_column)" },
  { key: "pmt", label: "PMT", category: "Financial & Forecasting", description: "Total periodic loan payment (principal plus interest).", syntax: "PMT(annual_rate, periods, present_value)" },
  { key: "ipmt", label: "IPMT", category: "Financial & Forecasting", description: "Interest portion of a specific loan payment period.", syntax: "IPMT(annual_rate, period, periods, present_value)" },
];

export type CriteriaOp = "=" | "!=" | ">" | "<" | ">=" | "<=" | "contains";

export type FunctionResult = {
  ok: boolean;
  message: string;
  scalar?: string;
  table?: DataTable;
  formula?: string;
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function toNumber(value: string): number {
  const n = parseFloat(String(value ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function colValues(table: DataTable, col: number): string[] {
  return table.rows.map((row) => (row[col] ?? "").trim());
}

function matchCriteria(cellRaw: string, op: CriteriaOp, target: string): boolean {
  const cell = (cellRaw ?? "").trim();
  const isNumeric = /^-?[\d,.]+$/.test(cell) && /^-?[\d,.]+$/.test(target);
  if (isNumeric && (op === ">" || op === "<" || op === ">=" || op === "<=")) {
    const a = toNumber(cell);
    const b = toNumber(target);
    if (op === ">") return a > b;
    if (op === "<") return a < b;
    if (op === ">=") return a >= b;
    return a <= b;
  }
  const c = cell.toLowerCase();
  const t = target.trim().toLowerCase();
  switch (op) {
    case "=": return c === t;
    case "!=": return c !== t;
    case "contains": return c.includes(t);
    case ">": return c > t;
    case "<": return c < t;
    case ">=": return c >= t;
    case "<=": return c <= t;
    default: return false;
  }
}

export type Criterion = { col: number; op: CriteriaOp; value: string };

function rowMatchesAll(table: DataTable, rowIndex: number, criteria: Criterion[]): boolean {
  return criteria.every((c) => matchCriteria(table.rows[rowIndex][c.col] ?? "", c.op, c.value));
}

function singleColumnTable(header: string, values: string[], sourceName: string): DataTable {
  return { headers: [header], rows: values.map((v) => [v]), sourceName };
}

// ---------------------------------------------------------------------------
// Lookup & Reference
// ---------------------------------------------------------------------------

export function xlookup(table: DataTable, lookupValue: string, lookupCol: number, returnCol: number, ifNotFound = "#N/A"): FunctionResult {
  const target = lookupValue.trim().toLowerCase();
  const idx = table.rows.findIndex((row) => (row[lookupCol] ?? "").trim().toLowerCase() === target);
  const value = idx >= 0 ? (table.rows[idx][returnCol] ?? "") : ifNotFound;
  return {
    ok: true,
    message: idx >= 0 ? `Found in row ${idx + 2}.` : "No match found, returned default.",
    scalar: value,
    formula: `=XLOOKUP("${lookupValue}", ${table.headers[lookupCol]}, ${table.headers[returnCol]}, "${ifNotFound}")`,
  };
}

export function indexMatch(table: DataTable, lookupValue: string, lookupCol: number, returnCol: number): FunctionResult {
  const target = lookupValue.trim().toLowerCase();
  const idx = table.rows.findIndex((row) => (row[lookupCol] ?? "").trim().toLowerCase() === target);
  return {
    ok: true,
    message: idx >= 0 ? `Matched at position ${idx + 1}.` : "No match found.",
    scalar: idx >= 0 ? (table.rows[idx][returnCol] ?? "") : "#N/A",
    formula: `=INDEX(${table.headers[returnCol]}, MATCH("${lookupValue}", ${table.headers[lookupCol]}, 0))`,
  };
}

export function offsetAggregate(table: DataTable, col: number, startRow: number, height: number, mode: "sum" | "average"): FunctionResult {
  const start = Math.max(0, startRow);
  const slice = table.rows.slice(start, start + Math.max(1, height)).map((row) => toNumber(row[col] ?? ""));
  const sum = slice.reduce((a, b) => a + b, 0);
  const result = mode === "sum" ? sum : (slice.length ? sum / slice.length : 0);
  return {
    ok: true,
    message: `${mode === "sum" ? "Sum" : "Average"} of ${slice.length} rows starting at row ${start + 2}.`,
    scalar: formatNumber(result),
    formula: `=${mode === "sum" ? "SUM" : "AVERAGE"}(OFFSET(${table.headers[col]}, ${startRow}, 0, ${height}, 1))`,
  };
}

// ---------------------------------------------------------------------------
// Conditional Mathematics
// ---------------------------------------------------------------------------

export function sumifs(table: DataTable, sumCol: number, criteria: Criterion[]): FunctionResult {
  let total = 0;
  let count = 0;
  table.rows.forEach((_, i) => {
    if (rowMatchesAll(table, i, criteria)) { total += toNumber(table.rows[i][sumCol] ?? ""); count += 1; }
  });
  return { ok: true, message: `Summed ${count} matching rows.`, scalar: formatNumber(total), formula: buildConditionalFormula("SUMIFS", table, sumCol, criteria) };
}

export function countifs(table: DataTable, criteria: Criterion[]): FunctionResult {
  let count = 0;
  table.rows.forEach((_, i) => { if (rowMatchesAll(table, i, criteria)) count += 1; });
  return { ok: true, message: "Counted matching rows.", scalar: String(count), formula: buildConditionalFormula("COUNTIFS", table, -1, criteria) };
}

export function maxifs(table: DataTable, maxCol: number, criteria: Criterion[]): FunctionResult {
  const values = table.rows.filter((_, i) => rowMatchesAll(table, i, criteria)).map((row) => toNumber(row[maxCol] ?? ""));
  if (!values.length) return { ok: true, message: "No rows matched the criteria.", scalar: "0", formula: buildConditionalFormula("MAXIFS", table, maxCol, criteria) };
  return { ok: true, message: `Maximum across ${values.length} matching rows.`, scalar: formatNumber(Math.max(...values)), formula: buildConditionalFormula("MAXIFS", table, maxCol, criteria) };
}

export function minifs(table: DataTable, minCol: number, criteria: Criterion[]): FunctionResult {
  const values = table.rows.filter((_, i) => rowMatchesAll(table, i, criteria)).map((row) => toNumber(row[minCol] ?? ""));
  if (!values.length) return { ok: true, message: "No rows matched the criteria.", scalar: "0", formula: buildConditionalFormula("MINIFS", table, minCol, criteria) };
  return { ok: true, message: `Minimum across ${values.length} matching rows.`, scalar: formatNumber(Math.min(...values)), formula: buildConditionalFormula("MINIFS", table, minCol, criteria) };
}

export function correl(table: DataTable, colA: number, colB: number): FunctionResult {
  const pairs = table.rows.map(r => [toNumber(r[colA] ?? ""), toNumber(r[colB] ?? "")]).filter(([a,b]) => Number.isFinite(a) && Number.isFinite(b));
  if (pairs.length < 2) return { ok: false, message: "CORREL needs at least two numeric pairs." };
  const mx = pairs.reduce((a,p)=>a+p[0],0)/pairs.length, my = pairs.reduce((a,p)=>a+p[1],0)/pairs.length;
  const num = pairs.reduce((a,p)=>a+(p[0]-mx)*(p[1]-my),0);
  const den = Math.sqrt(pairs.reduce((a,p)=>a+(p[0]-mx)**2,0)*pairs.reduce((a,p)=>a+(p[1]-my)**2,0));
  if (!den) return { ok:false, message:"CORREL is undefined when one column has no variation." };
  return { ok:true, message:`Pearson correlation across ${pairs.length} pairs.`, scalar: num/den.toFixed ? String((num/den).toFixed(6)) : String(num/den), formula:`=CORREL(${table.headers[colA]}, ${table.headers[colB]})` };
}

export function stdevS(table: DataTable, col: number): FunctionResult {
  const nums = colValues(table,col).map(toNumber).filter(Number.isFinite);
  if (nums.length < 2) return { ok:false, message:"STDEV.S needs at least two numeric values." };
  const mean=nums.reduce((a,b)=>a+b,0)/nums.length;
  const value=Math.sqrt(nums.reduce((a,b)=>a+(b-mean)**2,0)/(nums.length-1));
  return { ok:true, message:`Sample standard deviation across ${nums.length} values.`, scalar:formatNumber(value), formula:`=STDEV.S(${table.headers[col]})` };
}

export function varS(table: DataTable, col: number): FunctionResult {
  const nums=colValues(table,col).map(toNumber).filter(Number.isFinite);
  if (nums.length < 2) return { ok:false, message:"VAR.S needs at least two numeric values." };
  const mean=nums.reduce((a,b)=>a+b,0)/nums.length;
  const value=nums.reduce((a,b)=>a+(b-mean)**2,0)/(nums.length-1);
  return { ok:true, message:`Sample variance across ${nums.length} values.`, scalar:formatNumber(value), formula:`=VAR.S(${table.headers[col]})` };
}

function linearRegression(xs:number[], ys:number) {
  const n=xs.length, mx=xs.reduce((a,b)=>a+b,0)/n, my=ys.reduce((a,b)=>a+b,0)/n;
  const denom=xs.reduce((a,x)=>a+(x-mx)**2,0);
  if(!denom) return null;
  const slope=xs.reduce((a,x,i)=>a+(x-mx)*(ys[i]-my),0)/denom;
  return { slope, intercept: my-slope*mx };
}

export function forecastLinear(table: DataTable, xCol:number, yCol:number, targetX:number): FunctionResult {
  const pairs=table.rows.map(r=>[toNumber(r[xCol]??""),toNumber(r[yCol]??"")]).filter(([x,y])=>Number.isFinite(x)&&Number.isFinite(y));
  if(pairs.length<2) return {ok:false,message:"FORECAST.LINEAR needs at least two numeric x/y pairs."};
  const fit=linearRegression(pairs.map(p=>p[0]),pairs.map(p=>p[1]));
  if(!fit) return {ok:false,message:"FORECAST.LINEAR needs variation in known_x."};
  const value=fit.slope*targetX+fit.intercept;
  return {ok:true,message:`Linear forecast for x=${targetX} from ${pairs.length} observations.`,scalar:formatNumber(value),formula:`=FORECAST.LINEAR(${targetX}, ${table.headers[yCol]}, ${table.headers[xCol]})`};
}

export function forecastETS(table: DataTable, timelineCol:number, valueCol:number, target:number, seasonality:number): FunctionResult {
  const points=table.rows.map(r=>({t:toNumber(r[timelineCol]??""),y:toNumber(r[valueCol]??"")})).filter(p=>Number.isFinite(p.t)&&Number.isFinite(p.y)).sort((a,b)=>a.t-b.t);
  if(points.length<4) return {ok:false,message:"FORECAST.ETS needs at least four numeric observations."};
  const period=Math.max(1,Math.floor(seasonality));
  const alpha=0.35, beta=0.1;
  let level=points[0].y, trend=points.length>1?points[1].y-points[0].y:0;
  const season=Array.from({length:period},(_,i)=>points[i]?.y-(level+i*trend));
  points.forEach((p,i)=>{ const s=season[i%period]??0; const prev=level; level=alpha*(p.y-s)+(1-alpha)*(level+trend); trend=beta*(level-prev)+(1-beta)*trend; season[i%period]=(1-alpha)*s+alpha*(p.y-level); });
  const steps=Math.max(1,Math.round(target-points[points.length-1].t));
  const idx=(points.length+steps-1)%period;
  const value=level+steps*trend+(season[idx]??0);
  return {ok:true,message:`ETS-style seasonal forecast for target ${target} using seasonality ${period}.`,scalar:formatNumber(value),formula:`=FORECAST.ETS(${target}, ${table.headers[valueCol]}, ${table.headers[timelineCol]}, ${period})`};
}

export function sumproduct(table: DataTable, colA: number, colB: number): FunctionResult {
  let total = 0;
  table.rows.forEach((row) => { total += toNumber(row[colA] ?? "") * toNumber(row[colB] ?? ""); });
  return { ok: true, message: "Multiplied both columns row by row and summed.", scalar: formatNumber(total), formula: `=SUMPRODUCT(${table.headers[colA]}, ${table.headers[colB]})` };
}

// ---------------------------------------------------------------------------
// Dynamic Arrays
// ---------------------------------------------------------------------------

export function filterTableByCriterion(table: DataTable, criterion: Criterion): FunctionResult {
  const rows = table.rows.filter((_, i) => rowMatchesAll(table, i, [criterion]));
  return {
    ok: true,
    message: `${rows.length} rows match.`,
    table: { headers: table.headers, rows, sourceName: "filtered" },
    formula: `=FILTER(table, ${table.headers[criterion.col]} ${criterion.op} "${criterion.value}")`,
  };
}

export function uniqueValues(table: DataTable, col: number): FunctionResult {
  const seen = new Set<string>();
  const values: string[] = [];
  colValues(table, col).forEach((v) => { const k = v.toLowerCase(); if (v && !seen.has(k)) { seen.add(k); values.push(v); } });
  return {
    ok: true,
    message: `${values.length} unique values.`,
    table: singleColumnTable(table.headers[col], values, "unique"),
    formula: `=UNIQUE(${table.headers[col]})`,
  };
}

export function chooseCols(table:DataTable,columns:number[]):FunctionResult {
  const valid=columns.map(n=>n-1).filter(n=>n>=0&&n<table.headers.length);
  if(!valid.length)return {ok:false,message:"Choose at least one valid column number."};
  const headers=valid.map(i=>table.headers[i]), rows=table.rows.map(r=>valid.map(i=>r[i]??""));
  return {ok:true,message:`Selected ${valid.length} columns.`,table:{headers,rows,sourceName:"choosecols"},formula:`=CHOOSECOLS(table, ${columns.join(", ")})`};
}
export function vstackTables(table:DataTable,other:DataTable|null):FunctionResult {
  if(!other)return {ok:false,message:"Import a second table to use VSTACK."};
  const width=Math.max(table.headers.length,other.headers.length);
  const headers=table.headers.length>=other.headers.length?table.headers:other.headers;
  const pad=(r:string[])=>Array.from({length:width},(_,i)=>r[i]??"");
  return {ok:true,message:`Stacked ${table.rows.length+other.rows.length} rows from two tables.`,table:{headers:[...headers],rows:[...table.rows.map(pad),...other.rows.map(pad)],sourceName:"vstack"},formula:`=VSTACK(${table.sourceName}, ${other.sourceName})`};
}
export function tocol(table:DataTable,col:number,ignoreBlanks=true):FunctionResult {
  const values=table.rows.flatMap(r=>r).filter(v=>!ignoreBlanks||String(v).trim()!=="");
  return {ok:true,message:`Flattened ${values.length} values into one column.`,table:singleColumnTable("TOCOL",values,"tocol"),formula:`=TOCOL(table, ${ignoreBlanks?1:0})`};
}
export function sortTableByColumn(table: DataTable, col: number, dir: "asc" | "desc"): FunctionResult {
  const numeric = table.rows.every((row) => !row[col] || /^-?[\d,.]+$/.test((row[col] ?? "").trim()));
  const rows = [...table.rows].sort((a, b) => {
    if (numeric) return toNumber(a[col] ?? "") - toNumber(b[col] ?? "");
    return (a[col] ?? "").localeCompare(b[col] ?? "", undefined, { sensitivity: "base" });
  });
  if (dir === "desc") rows.reverse();
  return {
    ok: true,
    message: `Sorted by ${table.headers[col]} (${dir}).`,
    table: { headers: table.headers, rows, sourceName: "sorted" },
    formula: `=SORT(table, ${col + 1}, ${dir === "asc" ? 1 : -1})`,
  };
}

// ---------------------------------------------------------------------------
// Advanced Logic (LET-style named aggregate then compare)
// ---------------------------------------------------------------------------

export function letCalc(table: DataTable, col: number, agg: "sum" | "average" | "count" | "max" | "min", op: CriteriaOp, compareValue: string): FunctionResult {
  const nums = colValues(table, col).map(toNumber);
  let named = 0;
  if (agg === "sum") named = nums.reduce((a, b) => a + b, 0);
  else if (agg === "average") named = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
  else if (agg === "count") named = nums.length;
  else if (agg === "max") named = Math.max(...nums, 0);
  else named = Math.min(...(nums.length ? nums : [0]));

  const result = matchCriteria(String(named), op, compareValue);
  return {
    ok: true,
    message: `LET named result = ${formatNumber(named)}. Condition (${op} ${compareValue}) is ${result ? "TRUE" : "FALSE"}.`,
    scalar: `${formatNumber(named)} -> ${result ? "TRUE" : "FALSE"}`,
    formula: `=LET(val, ${agg.toUpperCase()}(${table.headers[col]}), val ${op} ${compareValue})`,
  };
}

// ---------------------------------------------------------------------------
// Financial & Forecasting
// ---------------------------------------------------------------------------

function parseDateSafe(value: string): Date | null {
  const d = new Date(value.trim());
  return Number.isNaN(d.getTime()) ? null : d;
}

export function npv(table: DataTable, valuesCol: number, ratePct: number): FunctionResult {
  const rate = ratePct / 100;
  if (rate <= -1) return { ok: false, message: "Rate must be greater than -100%." };
  const values = table.rows.map((row) => toNumber(row[valuesCol] ?? ""));
  const value = values.reduce((sum, cashflow, i) => sum + cashflow / Math.pow(1 + rate, i + 1), 0);
  return { ok: true, message: `Discounted ${values.length} periodic cash flows.`, scalar: formatNumber(value), formula: `=NPV(${ratePct}%, ${table.headers[valuesCol]})` };
}

export function irr(table: DataTable, valuesCol: number): FunctionResult {
  const values = table.rows.map((row) => toNumber(row[valuesCol] ?? ""));
  if (values.length < 2 || !values.some((v) => v > 0) || !values.some((v) => v < 0)) return { ok: false, message: "IRR needs at least one positive and one negative cash flow." };
  let guess = 0.1;
  for (let i = 0; i < 100; i += 1) {
    let f = 0, df = 0;
    values.forEach((cashflow, periodIndex) => {
      const base = Math.pow(1 + guess, periodIndex);
      f += cashflow / base;
      if (periodIndex > 0) df -= periodIndex * cashflow / (base * (1 + guess));
    });
    if (Math.abs(f) < 1e-9 || !Number.isFinite(df) || Math.abs(df) < 1e-12) break;
    const next = guess - f / df;
    if (!Number.isFinite(next) || next <= -0.999999 || next > 1000) break;
    if (Math.abs(next - guess) < 1e-10) { guess = next; break; }
    guess = next;
  }
  const npvAt = (r: number) => values.reduce((sum, cashflow, i) => sum + cashflow / Math.pow(1 + r, i), 0);
  if (!Number.isFinite(guess) || Math.abs(npvAt(guess)) > 1e-5) return { ok: false, message: "Could not converge on an IRR for these cash flows." };
  return { ok: true, message: `Periodic internal rate of return across ${values.length} cash flows.`, scalar: `${(guess * 100).toFixed(2)}%`, formula: `=IRR(${table.headers[valuesCol]})` };
}

export function xnpv(table: DataTable, valuesCol: number, datesCol: number, ratePct: number): FunctionResult {
  const rate = ratePct / 100;
  const points: Array<{ value: number; date: Date }> = [];
  table.rows.forEach((row) => {
    const d = parseDateSafe(row[datesCol] ?? "");
    if (d) points.push({ value: toNumber(row[valuesCol] ?? ""), date: d });
  });
  if (points.length < 2) return { ok: false, message: "Need at least two dated cash flows with valid dates." };
  points.sort((a, b) => a.date.getTime() - b.date.getTime());
  const t0 = points[0].date.getTime();
  const npv = points.reduce((acc, p) => {
    const years = (p.date.getTime() - t0) / (365 * 24 * 3600 * 1000);
    return acc + p.value / Math.pow(1 + rate, years);
  }, 0);
  return { ok: true, message: `XNPV over ${points.length} cash flows at ${ratePct}%.`, scalar: formatNumber(npv), formula: `=XNPV(${ratePct / 100}, ${table.headers[valuesCol]}, ${table.headers[datesCol]})` };
}

export function xirr(table: DataTable, valuesCol: number, datesCol: number): FunctionResult {
  const points: Array<{ value: number; date: Date }> = [];
  table.rows.forEach((row) => {
    const d = parseDateSafe(row[datesCol] ?? "");
    if (d) points.push({ value: toNumber(row[valuesCol] ?? ""), date: d });
  });
  if (points.length < 2) return { ok: false, message: "Need at least two dated cash flows." };
  points.sort((a, b) => a.date.getTime() - b.date.getTime());
  const t0 = points[0].date.getTime();
  const npvAt = (rate: number) => points.reduce((acc, p) => {
    const years = (p.date.getTime() - t0) / (365 * 24 * 3600 * 1000);
    return acc + p.value / Math.pow(1 + rate, years);
  }, 0);

  // Bisection between -0.9999 and 10
  let low = -0.9999;
  let high = 10;
  let fLow = npvAt(low);
  let fHigh = npvAt(high);
  if (fLow * fHigh > 0) return { ok: false, message: "Could not solve XIRR (cash flows may not change sign)." };
  let mid = 0;
  for (let i = 0; i < 100; i += 1) {
    mid = (low + high) / 2;
    const fMid = npvAt(mid);
    if (Math.abs(fMid) < 1e-7) break;
    if (fLow * fMid < 0) { high = mid; fHigh = fMid; } else { low = mid; fLow = fMid; }
  }
  return { ok: true, message: `Internal rate of return across ${points.length} cash flows.`, scalar: `${(mid * 100).toFixed(2)}%`, formula: `=XIRR(${table.headers[valuesCol]}, ${table.headers[datesCol]})` };
}

export function pmt(annualRatePct: number, periods: number, presentValue: number): FunctionResult {
  const r = annualRatePct / 100 / 12;
  const n = periods;
  if (n <= 0) return { ok: false, message: "Number of periods must be positive." };
  const payment = r === 0 ? presentValue / n : (presentValue * r) / (1 - Math.pow(1 + r, -n));
  return { ok: true, message: `Monthly payment over ${n} periods.`, scalar: formatNumber(payment), formula: `=PMT(${annualRatePct}%/12, ${n}, ${presentValue})` };
}

export function ipmt(annualRatePct: number, period: number, periods: number, presentValue: number): FunctionResult {
  const r = annualRatePct / 100 / 12;
  const n = periods;
  if (period < 1 || period > n) return { ok: false, message: "Payment period is out of range." };
  const payment = r === 0 ? presentValue / n : (presentValue * r) / (1 - Math.pow(1 + r, -n));
  // Remaining balance before this period
  const balance = r === 0 ? presentValue - (period - 1) * payment : presentValue * Math.pow(1 + r, period - 1) - payment * ((Math.pow(1 + r, period - 1) - 1) / r);
  const interest = balance * r;
  return { ok: true, message: `Interest portion of payment ${period} of ${n}.`, scalar: formatNumber(interest), formula: `=IPMT(${annualRatePct}%/12, ${period}, ${n}, ${presentValue})` };
}

// ---------------------------------------------------------------------------
// Additional Lookup
// ---------------------------------------------------------------------------

export function vlookup(table: DataTable, lookupValue: string, lookupCol: number, returnCol: number): FunctionResult {
  return { ...indexMatch(table, lookupValue, lookupCol, returnCol), formula: `=VLOOKUP("${lookupValue}", ${table.headers[lookupCol]}:${table.headers[returnCol]}, ${returnCol - lookupCol + 1}, FALSE)` };
}

export function hlookup(table: DataTable, lookupValue: string, lookupCol: number, returnCol: number): FunctionResult {
  return { ...indexMatch(table, lookupValue, lookupCol, returnCol), formula: `=HLOOKUP("${lookupValue}", ${table.headers[lookupCol]}, ${table.headers[returnCol]})` };
}

// ---------------------------------------------------------------------------
// Conditional average
// ---------------------------------------------------------------------------

export function averageifs(table: DataTable, avgCol: number, criteria: Criterion[]): FunctionResult {
  let total = 0;
  let count = 0;
  table.rows.forEach((_, i) => {
    if (rowMatchesAll(table, i, criteria)) { total += toNumber(table.rows[i][avgCol] ?? ""); count += 1; }
  });
  const avg = count ? total / count : 0;
  return { ok: true, message: `Averaged ${count} matching rows.`, scalar: formatNumber(avg), formula: `=AVERAGEIFS(${table.headers[avgCol]}, ...)` };
}

// ---------------------------------------------------------------------------
// Statistical & Math
// ---------------------------------------------------------------------------

export type RoundMode = "round" | "up" | "down";

export function roundColumn(table: DataTable, col: number, decimals: number, mode: RoundMode): FunctionResult {
  const factor = Math.pow(10, decimals);
  const rows = table.rows.map((row) => {
    const copy = [...row];
    const n = toNumber(row[col] ?? "");
    let v: number;
    if (mode === "up") v = Math.ceil(n * factor) / factor;
    else if (mode === "down") v = Math.floor(n * factor) / factor;
    else v = Math.round(n * factor) / factor;
    copy[col] = String(v);
    return copy;
  });
  const fn = mode === "up" ? "ROUNDUP" : mode === "down" ? "ROUNDDOWN" : "ROUND";
  return { ok: true, message: `Rounded ${table.headers[col]} to ${decimals} decimals.`, table: { headers: table.headers, rows, sourceName: "rounded" }, formula: `=${fn}(${table.headers[col]}, ${decimals})` };
}

export type AggFn = "sum" | "average" | "min" | "max" | "median" | "count" | "stdev" | "var";

export function aggregateColumn(table: DataTable, col: number, fn: AggFn): FunctionResult {
  const nums = colValues(table, col).map(toNumber).filter((n) => Number.isFinite(n));
  if (!nums.length) return { ok: false, message: "No numeric values found in the column." };
  let result = 0;
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  if (fn === "sum") result = nums.reduce((a, b) => a + b, 0);
  else if (fn === "average") result = mean;
  else if (fn === "min") result = Math.min(...nums);
  else if (fn === "max") result = Math.max(...nums);
  else if (fn === "count") result = nums.length;
  else if (fn === "median") { const s = [...nums].sort((a, b) => a - b); result = s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2; }
  else if (fn === "var") result = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1 || 1);
  else result = Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1 || 1));
  return { ok: true, message: `${fn.toUpperCase()} of ${table.headers[col]} over ${nums.length} values.`, scalar: formatNumber(result), formula: `=${fn.toUpperCase()}(${table.headers[col]})` };
}

export function rankColumn(table: DataTable, col: number, order: "desc" | "asc"): FunctionResult {
  const nums = table.rows.map((row) => toNumber(row[col] ?? ""));
  const sorted = [...nums].sort((a, b) => (order === "desc" ? b - a : a - b));
  const rankOf = (v: number) => sorted.indexOf(v) + 1;
  const headers = [...table.headers, `Rank of ${table.headers[col]}`];
  const rows = table.rows.map((row, i) => [...row, String(rankOf(nums[i]))]);
  return { ok: true, message: `Added rank column (${order}).`, table: { headers, rows, sourceName: "ranked" }, formula: `=RANK(${table.headers[col]}, range, ${order === "asc" ? 1 : 0})` };
}

export function percentileColumn(table: DataTable, col: number, k: number): FunctionResult {
  const nums = colValues(table, col).map(toNumber).filter((n) => Number.isFinite(n)).sort((a, b) => a - b);
  if (!nums.length) return { ok: false, message: "No numeric values found." };
  const pos = (nums.length - 1) * k;
  const base = Math.floor(pos);
  const rest = pos - base;
  const value = nums[base] + (nums[base + 1] !== undefined ? rest * (nums[base + 1] - nums[base]) : 0);
  return { ok: true, message: `${Math.round(k * 100)}th percentile of ${table.headers[col]}.`, scalar: formatNumber(value), formula: `=PERCENTILE(${table.headers[col]}, ${k})` };
}

// ---------------------------------------------------------------------------
// Text functions (produce a new column)
// ---------------------------------------------------------------------------

function addColumn(table: DataTable, header: string, values: string[]): DataTable {
  return { headers: [...table.headers, header], rows: table.rows.map((row, i) => [...row, values[i] ?? ""]), sourceName: "result" };
}

export function concatColumns(table: DataTable, colA: number, colB: number, sep: string): FunctionResult {
  const values = table.rows.map((row) => `${row[colA] ?? ""}${sep}${row[colB] ?? ""}`.trim());
  return { ok: true, message: "Combined columns.", table: addColumn(table, `${table.headers[colA]} + ${table.headers[colB]}`, values), formula: `=${table.headers[colA]} & "${sep}" & ${table.headers[colB]}` };
}

export function textSplit(table: DataTable, col: number, delimiter: string): FunctionResult {
  const max = Math.max(1, ...table.rows.map((row) => (row[col] ?? "").split(delimiter).length));
  const headers = [...table.headers];
  for (let i = 0; i < max; i += 1) headers.push(`${table.headers[col]} ${i + 1}`);
  const rows = table.rows.map((row) => {
    const parts = (row[col] ?? "").split(delimiter);
    return [...row, ...Array.from({ length: max }, (_, i) => (parts[i] ?? "").trim())];
  });
  return { ok: true, message: `Split into ${max} columns.`, table: { headers, rows, sourceName: "split" }, formula: `=TEXTSPLIT(${table.headers[col]}, "${delimiter}")` };
}

export function textJoin(table: DataTable, col: number, delimiter: string): FunctionResult {
  const joined = colValues(table, col).filter(Boolean).join(delimiter);
  return { ok: true, message: `Joined ${table.rows.length} values.`, scalar: joined, formula: `=TEXTJOIN("${delimiter}", TRUE, ${table.headers[col]})` };
}

export type SubstringMode = "left" | "right" | "mid";

export function substringColumn(table: DataTable, col: number, mode: SubstringMode, n: number, start = 1): FunctionResult {
  const values = table.rows.map((row) => {
    const s = row[col] ?? "";
    if (mode === "left") return s.slice(0, n);
    if (mode === "right") return s.slice(Math.max(0, s.length - n));
    return s.slice(start - 1, start - 1 + n);
  });
  const label = mode === "left" ? `LEFT(${n})` : mode === "right" ? `RIGHT(${n})` : `MID(${start},${n})`;
  return { ok: true, message: "Extracted characters.", table: addColumn(table, `${table.headers[col]} ${label}`, values), formula: mode === "mid" ? `=MID(${table.headers[col]}, ${start}, ${n})` : `=${mode.toUpperCase()}(${table.headers[col]}, ${n})` };
}

export function trimCleanColumn(table: DataTable, col: number): FunctionResult {
  const rows = table.rows.map((row) => { const copy = [...row]; copy[col] = (row[col] ?? "").replace(/[\u0000-\u001F\u007F]/g, "").replace(/\s+/g, " ").trim(); return copy; });
  return { ok: true, message: "Trimmed spaces and removed non-printable characters.", table: { headers: table.headers, rows, sourceName: "trimmed" }, formula: `=TRIM(CLEAN(${table.headers[col]}))` };
}

export type CaseMode = "upper" | "lower" | "proper";

export function caseColumn(table: DataTable, col: number, mode: CaseMode): FunctionResult {
  const rows = table.rows.map((row) => {
    const copy = [...row];
    const s = row[col] ?? "";
    copy[col] = mode === "upper" ? s.toUpperCase() : mode === "lower" ? s.toLowerCase() : s.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
    return copy;
  });
  return { ok: true, message: `Applied ${mode} case.`, table: { headers: table.headers, rows, sourceName: "cased" }, formula: `=${mode.toUpperCase()}(${table.headers[col]})` };
}

export function regexReplaceColumn(table:DataTable,col:number,pattern:string,replacement:string):FunctionResult {
  let re:RegExp; try{re=new RegExp(pattern,"g");}catch{return {ok:false,message:"Invalid regular expression."};}
  const values=table.rows.map(r=>{try{return (r[col]??"").replace(re,replacement)}catch{return r[col]??""}});
  return {ok:true,message:`Applied REGEXREPLACE to ${values.length} rows.`,table:addColumn(table,`REGEXREPLACE(${table.headers[col]})`,values),formula:`=REGEXREPLACE(${table.headers[col]}, "${pattern}", "${replacement}")`};
}
export function regexTestColumn(table:DataTable,col:number,pattern:string):FunctionResult {
  let re:RegExp; try{re=new RegExp(pattern);}catch{return {ok:false,message:"Invalid regular expression."};}
  const values=table.rows.map(r=>String(re.test(r[col]??"")).toUpperCase());
  return {ok:true,message:`Tested ${values.length} rows with REGEXTEST.`,table:addColumn(table,`REGEXTEST(${table.headers[col]})`,values),formula:`=REGEXTEST(${table.headers[col]}, "${pattern}")`};
}
export function regexExtractColumn(table:DataTable,col:number,pattern:string):FunctionResult {
  let re:RegExp; try{re=new RegExp(pattern);}catch{return {ok:false,message:"Invalid regular expression."};}
  const values=table.rows.map(r=>{const m=String(r[col]??"").match(re); return m?.[0]??""});
  return {ok:true,message:`Extracted regex matches from ${values.length} rows.`,table:addColumn(table,`REGEXEXTRACT(${table.headers[col]})`,values),formula:`=REGEXEXTRACT(${table.headers[col]}, "${pattern}")`};
}
export function substituteColumn(table: DataTable, col: number, oldText: string, newText: string): FunctionResult {
  const rows = table.rows.map((row) => { const copy = [...row]; copy[col] = (row[col] ?? "").split(oldText).join(newText); return copy; });
  return { ok: true, message: `Replaced "${oldText}" with "${newText}".`, table: { headers: table.headers, rows, sourceName: "substituted" }, formula: `=SUBSTITUTE(${table.headers[col]}, "${oldText}", "${newText}")` };
}

export function lenColumn(table: DataTable, col: number): FunctionResult {
  const values = table.rows.map((row) => String((row[col] ?? "").length));
  return { ok: true, message: "Added length column.", table: addColumn(table, `Length of ${table.headers[col]}`, values), formula: `=LEN(${table.headers[col]})` };
}

export function textBeforeAfter(table: DataTable, col: number, delimiter: string, mode: "before" | "after"): FunctionResult {
  const values = table.rows.map((row) => {
    const s = row[col] ?? "";
    const idx = s.indexOf(delimiter);
    if (idx < 0) return mode === "before" ? s : "";
    return mode === "before" ? s.slice(0, idx) : s.slice(idx + delimiter.length);
  });
  return { ok: true, message: `Extracted text ${mode} "${delimiter}".`, table: addColumn(table, `${table.headers[col]} ${mode}`, values.map((v) => v.trim())), formula: `=TEXT${mode.toUpperCase()}(${table.headers[col]}, "${delimiter}")` };
}

// ---------------------------------------------------------------------------
// Date & Time
// ---------------------------------------------------------------------------

function toDate(value: string): Date | null {
  const d = new Date((value ?? "").trim());
  return Number.isNaN(d.getTime()) ? null : d;
}

const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

export function dateValue(table: DataTable, col:number): FunctionResult {
  const values=table.rows.map(r=>{const d=toDate(r[col]??""); return d?String(Math.floor(d.getTime()/86400000)+25569):"";});
  return {ok:true,message:"Converted date text to Excel-style serials.",table:addColumn(table,"DATEVALUE",values),formula:`=DATEVALUE(${table.headers[col]})`};
}
export function timeValue(table: DataTable, col:number): FunctionResult {
  const values=table.rows.map(r=>{const m=(r[col]??"").trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i); if(!m)return ""; let h=Number(m[1]),mi=Number(m[2]),sec=Number(m[3]||0); const ap=(m[4]||"").toUpperCase(); if(ap==="PM"&&h<12)h+=12;if(ap==="AM"&&h===12)h=0; return String((h*3600+mi*60+sec)/86400);});
  return {ok:true,message:"Converted time text to Excel-style fractional days.",table:addColumn(table,"TIMEVALUE",values),formula:`=TIMEVALUE(${table.headers[col]})`};
}
export function datedif(table: DataTable, startCol: number, endCol: number, unit: "days" | "months" | "years"): FunctionResult {
  const values = table.rows.map((row) => {
    const a = toDate(row[startCol] ?? "");
    const b = toDate(row[endCol] ?? "");
    if (!a || !b) return "";
    const ms = b.getTime() - a.getTime();
    if (unit === "days") return String(Math.round(ms / 86400000));
    if (unit === "years") return String(Math.floor((b.getFullYear() - a.getFullYear()) - ((b.getMonth() < a.getMonth() || (b.getMonth() === a.getMonth() && b.getDate() < a.getDate())) ? 1 : 0)));
    return String((b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth()));
  });
  return { ok: true, message: `Difference in ${unit}.`, table: addColumn(table, `Diff (${unit})`, values), formula: `=DATEDIF(${table.headers[startCol]}, ${table.headers[endCol]}, "${unit === "days" ? "d" : unit === "months" ? "m" : "y"}")` };
}

export function datePart(table: DataTable, col: number, part: "year" | "month" | "day"): FunctionResult {
  const values = table.rows.map((row) => {
    const d = toDate(row[col] ?? "");
    if (!d) return "";
    if (part === "year") return String(d.getFullYear());
    if (part === "month") return String(d.getMonth() + 1);
    return String(d.getDate());
  });
  return { ok: true, message: `Extracted ${part}.`, table: addColumn(table, `${part} of ${table.headers[col]}`, values), formula: `=${part.toUpperCase()}(${table.headers[col]})` };
}

export function weekdayColumn(table: DataTable, col: number): FunctionResult {
  const values = table.rows.map((row) => { const d = toDate(row[col] ?? ""); return d ? DAY_NAMES[d.getDay()] : ""; });
  return { ok: true, message: "Added weekday name.", table: addColumn(table, `Weekday of ${table.headers[col]}`, values), formula: `=TEXT(${table.headers[col]}, "dddd")` };
}

export function networkdays(table: DataTable, startCol: number, endCol: number): FunctionResult {
  const values = table.rows.map((row) => {
    const a = toDate(row[startCol] ?? "");
    const b = toDate(row[endCol] ?? "");
    if (!a || !b) return "";
    let count = 0;
    const cur = new Date(a);
    const end = new Date(b);
    const step = end >= a ? 1 : -1;
    while ((step > 0 && cur <= end) || (step < 0 && cur >= end)) {
      const day = cur.getDay();
      if (day !== 0 && day !== 6) count += 1;
      cur.setDate(cur.getDate() + step);
    }
    return String(step > 0 ? count : -count);
  });
  return { ok: true, message: "Counted working days.", table: addColumn(table, "Working days", values), formula: `=NETWORKDAYS(${table.headers[startCol]}, ${table.headers[endCol]})` };
}

export function eomonth(table: DataTable, col: number, months: number): FunctionResult {
  const values = table.rows.map((row) => {
    const d = toDate(row[col] ?? "");
    if (!d) return "";
    const target = new Date(d.getFullYear(), d.getMonth() + months + 1, 0);
    return target.toISOString().slice(0, 10);
  });
  return { ok: true, message: `End of month ${months >= 0 ? "+" : ""}${months}.`, table: addColumn(table, `EOMONTH ${months}`, values), formula: `=EOMONTH(${table.headers[col]}, ${months})` };
}

export function formatDateColumn(table: DataTable, col: number, pattern: "iso" | "us" | "long"): FunctionResult {
  const values = table.rows.map((row) => {
    const d = toDate(row[col] ?? "");
    if (!d) return "";
    if (pattern === "iso") return d.toISOString().slice(0, 10);
    if (pattern === "us") return `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}/${d.getFullYear()}`;
    return d.toLocaleDateString(undefined, { day: "numeric", month: "long", year: "numeric" });
  });
  return { ok: true, message: "Reformatted dates.", table: addColumn(table, `${table.headers[col]} formatted`, values), formula: `=TEXT(${table.headers[col]}, "${pattern === "iso" ? "yyyy-mm-dd" : pattern === "us" ? "mm/dd/yyyy" : "d mmmm yyyy"}")` };
}

// ---------------------------------------------------------------------------
// Logical
// ---------------------------------------------------------------------------

export function ifColumn(table: DataTable, col: number, op: CriteriaOp, value: string, ifTrue: string, ifFalse: string): FunctionResult {
  const values = table.rows.map((_, i) => (matchCriteria(table.rows[i][col] ?? "", op, value) ? ifTrue : ifFalse));
  return { ok: true, message: "Applied IF logic.", table: addColumn(table, `IF ${table.headers[col]}`, values), formula: `=IF(${table.headers[col]} ${op} "${value}", "${ifTrue}", "${ifFalse}")` };
}

export type Band = { threshold: number; label: string };

export function ifsBanding(table: DataTable, col: number, bands: Band[]): FunctionResult {
  const sorted = [...bands].sort((a, b) => b.threshold - a.threshold);
  const values = table.rows.map((row) => {
    const n = toNumber(row[col] ?? "");
    const band = sorted.find((b) => n >= b.threshold);
    return band ? band.label : "";
  });
  return { ok: true, message: "Applied banding.", table: addColumn(table, `Band of ${table.headers[col]}`, values), formula: `=IFS(${sorted.map((b) => `${table.headers[col]}>=${b.threshold}, "${b.label}"`).join(", ")})` };
}

export function andOrColumn(table: DataTable, criteria: Criterion[], mode: "and" | "or"): FunctionResult {
  const values = table.rows.map((_, i) => {
    const results = criteria.map((c) => matchCriteria(table.rows[i][c.col] ?? "", c.op, c.value));
    const pass = mode === "and" ? results.every(Boolean) : results.some(Boolean);
    return pass ? "TRUE" : "FALSE";
  });
  return { ok: true, message: `${mode.toUpperCase()} across criteria.`, table: addColumn(table, `${mode.toUpperCase()} result`, values), formula: `=${mode.toUpperCase()}(...)` };
}

export function iferrorColumn(table: DataTable, col: number, fallback: string): FunctionResult {
  const rows = table.rows.map((row) => { const copy = [...row]; const v = (row[col] ?? "").trim(); copy[col] = v && !/^#(N\/A|VALUE|REF|DIV\/0|NAME|NULL|NUM)/i.test(v) ? v : fallback; return copy; });
  return { ok: true, message: "Replaced blanks and errors.", table: { headers: table.headers, rows, sourceName: "iferror" }, formula: `=IFERROR(${table.headers[col]}, "${fallback}")` };
}

export type SwitchMap = { from: string; to: string };

export function switchColumn(table: DataTable, col: number, maps: SwitchMap[]): FunctionResult {
  const lookup = new Map(maps.map((m) => [m.from.trim().toLowerCase(), m.to]));
  const values = table.rows.map((row) => { const key = (row[col] ?? "").trim().toLowerCase(); return lookup.has(key) ? lookup.get(key)! : (row[col] ?? ""); });
  return { ok: true, message: "Mapped values.", table: addColumn(table, `${table.headers[col]} mapped`, values), formula: `=SWITCH(${table.headers[col]}, ${maps.map((m) => `"${m.from}", "${m.to}"`).join(", ")})` };
}

export type TypeCheckMode = "blank" | "number" | "text";

export function formulaTextColumn(table:DataTable,col:number):FunctionResult {
  const values=table.rows.map(r=>{const v=String(r[col]??""); return v.startsWith("=")?v:"#N/A";});
  return {ok:true,message:"Returned formula text where imported cells contain formulas.",table:addColumn(table,"FORMULATEXT",values),formula:`=FORMULATEXT(${table.headers[col]})`};
}
const ERROR_CODES:Record<string,number>={"#NULL!":1,"#DIV/0!":2,"#VALUE!":3,"#REF!":4,"#NAME?":5,"#NUM!":6,"#N/A":7,"#GETTING_DATA":8};
export function errorTypeColumn(table:DataTable,col:number):FunctionResult {
  const values=table.rows.map(r=>String(ERROR_CODES[String(r[col]??"").trim().toUpperCase()]??"#N/A"));
  return {ok:true,message:"Identified Excel-style error codes in the selected column.",table:addColumn(table,"ERROR.TYPE",values),formula:`=ERROR.TYPE(${table.headers[col]})`};
}
export function typeCheckColumn(table: DataTable, col: number, mode: TypeCheckMode): FunctionResult {
  const label = mode === "blank" ? "ISBLANK" : mode === "number" ? "ISNUMBER" : "ISTEXT";
  const values = table.rows.map((row) => {
    const raw = (row[col] ?? "").trim();
    if (mode === "blank") return raw === "" ? "TRUE" : "FALSE";
    const isNumeric = raw !== "" && /^-?[\d,.]+%?$/.test(raw);
    if (mode === "number") return isNumeric ? "TRUE" : "FALSE";
    return raw !== "" && !isNumeric ? "TRUE" : "FALSE";
  });
  return { ok: true, message: `Flagged rows with ${label}.`, table: addColumn(table, `${label} ${table.headers[col]}`, values), formula: `=${label}(${table.headers[col]})` };
}

export function indirectColumn(table: DataTable, selectorCol: number): FunctionResult {
  const headerIndex = new Map(table.headers.map((h, i) => [h.trim().toLowerCase(), i]));
  let unresolved = 0;
  const values = table.rows.map((row) => {
    const columnName = (row[selectorCol] ?? "").trim().toLowerCase();
    const targetIndex = headerIndex.get(columnName);
    if (targetIndex === undefined) {
      unresolved++;
      return "#REF!";
    }
    return row[targetIndex] ?? "";
  });
  const message = unresolved > 0
    ? `Resolved ${table.rows.length - unresolved} of ${table.rows.length} rows. ${unresolved} referenced a column name that does not exist (#REF!).`
    : "Resolved every row to its dynamically selected column.";
  return { ok: true, message, table: addColumn(table, `INDIRECT(${table.headers[selectorCol]})`, values), formula: `=INDIRECT(${table.headers[selectorCol]})` };
}

export function sequenceColumn(table: DataTable, start: number, step: number): FunctionResult {
  const values = table.rows.map((_, i) => String(start + i * step));
  return { ok: true, message: `Generated a sequence of ${table.rows.length} numbers.`, table: addColumn(table, "SEQUENCE", values), formula: `=SEQUENCE(${table.rows.length}, ${start}, ${step})` };
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

export function formatNumber(n: number): string {
  if (!Number.isFinite(n)) return "0";
  return Number.isInteger(n) ? n.toLocaleString() : n.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function buildConditionalFormula(fn: "SUMIFS" | "COUNTIFS" | "MAXIFS" | "MINIFS", table: DataTable, sumCol: number, criteria: Criterion[]): string {
  const crit = criteria.map((c) => `${table.headers[c.col]}, "${c.op}${c.value}"`).join(", ");
  if (fn === "COUNTIFS") return `=COUNTIFS(${crit})`;
  return `=${fn}(${table.headers[sumCol]}, ${crit})`;
}

// ---------------------------------------------------------------------------
// Natural language interpreter: map a typed instruction to a function + guess
// columns from header names and quoted/keyword values.
// ---------------------------------------------------------------------------

export type Interpretation = {
  functionKey: FunctionKey;
  note: string;
  columnGuesses: Partial<{ lookupCol: number; returnCol: number; sumCol: number; valueCol: number; dateCol: number; sortDir: "asc" | "desc"; criterionValue: string; lookupValue: string }>;
};

function findColumn(table: DataTable, keywords: string[]): number {
  const lowerHeaders = table.headers.map((h) => h.toLowerCase());
  for (const kw of keywords) {
    const idx = lowerHeaders.findIndex((h) => h.includes(kw));
    if (idx >= 0) return idx;
  }
  return -1;
}

export function interpretInstruction(text: string, table: DataTable | null): Interpretation | null {
  const t = text.toLowerCase();
  if (!t.trim()) return null;

  const guesses: Interpretation["columnGuesses"] = {};
  if (table) {
    // pick a numeric-looking column as default sum/value col
    const numericCol = table.headers.findIndex((_, i) =>
      table.rows.length && table.rows.every((row) => !row[i] || /^-?[\d,.$%]+$/.test((row[i] ?? "").trim()))
    );
    if (numericCol >= 0) { guesses.sumCol = numericCol; guesses.valueCol = numericCol; }
    const dateCol = findColumn(table, ["date", "day", "period"]);
    if (dateCol >= 0) guesses.dateCol = dateCol;
  }

  if (guesses.sortDir === undefined) guesses.sortDir = /desc|highest|largest|top/.test(t) ? "desc" : "asc";

  const quoted = text.match(/"([^"]+)"/);
  if (quoted) { guesses.criterionValue = quoted[1]; guesses.lookupValue = quoted[1]; }

  const pick = (key: FunctionKey, note: string): Interpretation => ({ functionKey: key, note, columnGuesses: guesses });

  if (/\bxlookup\b/.test(t)) return pick("xlookup", "Detected XLOOKUP. Confirm the lookup and return columns.");
  if (/\bvlookup\b/.test(t)) return pick("vlookup", "Detected VLOOKUP.");
  if (/\bhlookup\b|row match/.test(t)) return pick("hlookup", "Detected a row match (HLOOKUP style).");
  if (/index|match/.test(t)) return pick("index-match", "Detected an INDEX and MATCH lookup.");
  if (/offset|running|rolling|from row/.test(t)) return pick("offset-sum", "Detected a dynamic OFFSET aggregate.");
  if (/look ?up|search for|find the/.test(t)) return pick("xlookup", "Detected a lookup request. Confirm the columns.");

  if (/average(if|ifs)?|mean .* where/.test(t)) return pick("averageifs", "Detected a conditional average (AVERAGEIFS).");
  if (/sum(if|ifs)?|total .* where|sum all/.test(t)) return pick("sumifs", "Detected a conditional sum (SUMIFS). Set your criteria.");
  if (/count(if|ifs)?|how many|number of/.test(t)) return pick("countifs", "Detected a conditional count (COUNTIFS). Set your criteria.");
  if (/sumproduct|weighted|multiply .* sum/.test(t)) return pick("sumproduct", "Detected a SUMPRODUCT (weighted total).");

  if (/round ?up|roundup/.test(t)) return pick("round", "Detected ROUNDUP.");
  if (/round ?down|rounddown/.test(t)) return pick("round", "Detected ROUNDDOWN.");
  if (/\bround\b/.test(t)) return pick("round", "Detected ROUND.");
  if (/median|standard deviation|stdev|variance|\bmin\b|\bmax\b|average of|mean of/.test(t)) return pick("aggregate", "Detected a statistical aggregate.");
  if (/\brank\b|ranking|position/.test(t)) return pick("rank", "Detected RANK.");
  if (/percentile|quartile/.test(t)) return pick("percentile", "Detected PERCENTILE.");

  if (/concat|combine .* column|join .* column|merge .* text|full name/.test(t)) return pick("concat", "Detected CONCAT (combine columns).");
  if (/textsplit|split .* into columns|split by/.test(t)) return pick("textsplit", "Detected TEXTSPLIT.");
  if (/textjoin|join .* into|join all/.test(t)) return pick("textjoin", "Detected TEXTJOIN.");
  if (/\bleft\b|\bright\b|\bmid\b|first .* characters|last .* characters/.test(t)) return pick("left-right-mid", "Detected LEFT/RIGHT/MID.");
  if (/trim|clean|remove .* spaces|extra spaces/.test(t)) return pick("trim-clean", "Detected TRIM/CLEAN.");
  if (/upper ?case|lower ?case|proper ?case|capitali/.test(t)) return pick("upper-lower-proper", "Detected case change.");
  if (/substitute|replace/.test(t)) return pick("substitute", "Detected SUBSTITUTE/REPLACE.");
  if (/\blen\b|length of|character count/.test(t)) return pick("len", "Detected LEN.");
  if (/textbefore|text before|before the/.test(t)) return pick("textbefore-after", "Detected TEXTBEFORE.");
  if (/textafter|text after|after the/.test(t)) return pick("textbefore-after", "Detected TEXTAFTER.");

  if (/datedif|days between|months between|years between|age/.test(t)) return pick("datedif", "Detected DATEDIF.");
  if (/\byear\b|\bmonth\b|\bday\b/.test(t) && /date/.test(t)) return pick("year-month-day", "Detected a date part extraction.");
  if (/weekday|day of week|day name/.test(t)) return pick("weekday", "Detected WEEKDAY.");
  if (/networkdays|working days|business days/.test(t)) return pick("networkdays", "Detected NETWORKDAYS.");
  if (/eomonth|end of month|month end/.test(t)) return pick("eomonth", "Detected EOMONTH.");
  if (/format date|date format/.test(t)) return pick("text-date", "Detected date formatting.");

  if (/\bifs\b|grade|band|tier|category based/.test(t)) return pick("ifs", "Detected IFS banding.");
  if (/iferror|handle error|replace error/.test(t)) return pick("iferror", "Detected IFERROR.");
  if (/switch|map values|translate values/.test(t)) return pick("switch", "Detected SWITCH.");
  if (/isblank|is blank|isnumber|is a number|istext|is text/.test(t)) return pick("is-type", "Detected ISBLANK/ISNUMBER/ISTEXT. Choose which check.");
  if (/indirect|dynamic column|column by name/.test(t)) return pick("indirect", "Detected INDIRECT.");
  if (/sequence|row number|running count|numbered list/.test(t)) return pick("sequence", "Detected SEQUENCE.");
  if (/\band\b|\bor\b/.test(t) && /if|condition|criteria/.test(t)) return pick("and-or", "Detected AND/OR logic.");
  if (/\bif\b|flag|label rows where/.test(t)) return pick("if", "Detected IF logic.");

  if (/filter|only rows|extract rows|show rows where/.test(t)) return pick("filter", "Detected a FILTER request. Choose the condition.");
  if (/unique|distinct|de-?dupe list/.test(t)) return pick("unique", "Detected a UNIQUE values request.");
  if (/\bsort\b|order by|arrange/.test(t)) return pick("sort", "Detected a SORT request.");
  if (/\blet\b|named|readable formula/.test(t)) return pick("let", "Detected a LET named calculation.");
  if (/xnpv|net present value/.test(t)) return pick("xnpv", "Detected XNPV. Choose values and dates columns.");
  if (/xirr|internal rate/.test(t)) return pick("xirr", "Detected XIRR. Choose values and dates columns.");
  if (/\bipmt\b|interest portion|interest part/.test(t)) return pick("ipmt", "Detected IPMT (interest portion).");
  if (/\bpmt\b|loan payment|monthly payment|amorti/.test(t)) return pick("pmt", "Detected PMT (loan payment).");

  return null;
}
