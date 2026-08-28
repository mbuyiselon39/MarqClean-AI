export type DaxCategory =
  | "Aggregation"
  | "Filter & Context"
  | "Time Intelligence"
  | "Relationship"
  | "Financial & Math"
  | "Logical & Information";

export const DAX_CATEGORIES: DaxCategory[] = [
  "Aggregation",
  "Filter & Context",
  "Time Intelligence",
  "Relationship",
  "Financial & Math",
  "Logical & Information",
];

export interface DaxPattern {
  name: string;
  category: DaxCategory;
  daxSyntax: string;
  daxExample: string;
  equivalent: string;
}

export interface DaxSyntaxCheckResult {
  ok: boolean;
  issues: string[];
  matchedFunction?: DaxPattern;
}

export const DAX_PATTERNS: DaxPattern[] = [
  {
    name: "CALCULATE",
    category: "Filter & Context",
    daxSyntax: "CALCULATE(<expression>, <filter1>, <filter2>, ...)",
    daxExample: 'CALCULATE(SUM(Sales[Amount]), Sales[Region] = "East", Sales[Status] = "Completed")',
    equivalent: "Equivalent to SUMIFS or filtered aggregate with multiple column criteria.",
  },
  {
    name: "TOTALYTD",
    category: "Time Intelligence",
    daxSyntax: "TOTALYTD(<expression>, <dates>, [<filter>], [<year_end_date>])",
    daxExample: "TOTALYTD(SUM(Sales[Amount]), 'Calendar'[Date])",
    equivalent: "Cumulative sum of values where date is on or before current date within the current year.",
  },
  {
    name: "SUMX",
    category: "Aggregation",
    daxSyntax: "SUMX(<table>, <expression>)",
    daxExample: "SUMX(OrderDetails, OrderDetails[UnitPrice] * OrderDetails[Quantity])",
    equivalent: "Equivalent to SUMPRODUCT of UnitPrice and Quantity columns.",
  },
  {
    name: "AVERAGEX",
    category: "Aggregation",
    daxSyntax: "AVERAGEX(<table>, <expression>)",
    daxExample: "AVERAGEX(Sales, Sales[GrossRevenue] - Sales[Discount])",
    equivalent: "Calculates net row values and returns arithmetic average.",
  },
  {
    name: "ALLEXCEPT",
    category: "Filter & Context",
    daxSyntax: "ALLEXCEPT(<table>, <column1>, <column2>, ...)",
    daxExample: "CALCULATE(SUM(Sales[Amount]), ALLEXCEPT(Sales, Sales[Category]))",
    equivalent: "Removes all table filters except the specified grouping column.",
  },
  {
    name: "RELATED",
    category: "Relationship",
    daxSyntax: "RELATED(<column>)",
    daxExample: "RELATED(Customer[CustomerSegment])",
    equivalent: "Equivalent to XLOOKUP or VLOOKUP fetching a corresponding value from the 'one' side of a relationship.",
  },
  {
    name: "RELATEDTABLE",
    category: "Relationship",
    daxSyntax: "RELATEDTABLE(<table>)",
    daxExample: "COUNTROWS(RELATEDTABLE(Orders))",
    equivalent: "Aggregates across related records on the 'many' side of a relationship.",
  },
  {
    name: "SAMEPERIODLASTYEAR",
    category: "Time Intelligence",
    daxSyntax: "SAMEPERIODLASTYEAR(<dates>)",
    daxExample: "CALCULATE(SUM(Sales[Amount]), SAMEPERIODLASTYEAR('Calendar'[Date]))",
    equivalent: "Filters dates shifted by exactly -1 full year.",
  },
  {
    name: "DIVIDE",
    category: "Financial & Math",
    daxSyntax: "DIVIDE(<numerator>, <denominator>, [<alternate_result>])",
    daxExample: "DIVIDE(SUM(Sales[Profit]), SUM(Sales[Revenue]), 0)",
    equivalent: "Safe division with zero division handling (equivalent to IFERROR(A / B, 0)).",
  },
  {
    name: "SWITCH",
    category: "Logical & Information",
    daxSyntax: "SWITCH(<expression>, <value1>, <result1>, ..., [<else>])",
    daxExample: 'SWITCH(TRUE(), [Margin] >= 0.3, "High", [Margin] >= 0.15, "Medium", "Low")',
    equivalent: "Tiered conditional evaluation equivalent to IFS or nested IF.",
  },
];

export function checkDaxSyntaxShape(formula: string): DaxSyntaxCheckResult {
  const issues: string[] = [];
  const trimmed = formula.trim();

  if (!trimmed) {
    return { ok: false, issues: ["Formula cannot be empty."] };
  }

  const openParens = (trimmed.match(/\(/g) || []).length;
  const closeParens = (trimmed.match(/\)/g) || []).length;
  if (openParens !== closeParens) {
    issues.push(`Unbalanced parentheses: ${openParens} opening '(' vs ${closeParens} closing ')'`);
  }

  const openBrackets = (trimmed.match(/\[/g) || []).length;
  const closeBrackets = (trimmed.match(/\]/g) || []).length;
  if (openBrackets !== closeBrackets) {
    issues.push(`Unbalanced brackets for column names: ${openBrackets} '[' vs ${closeBrackets} ']'`);
  }

  const quotes = (trimmed.match(/"/g) || []).length;
  if (quotes % 2 !== 0) {
    issues.push(`Unterminated string literal (odd number of quotes: ${quotes})`);
  }

  const upper = trimmed.toUpperCase();
  const matched = DAX_PATTERNS.find((p) => upper.includes(p.name));

  return {
    ok: issues.length === 0,
    issues,
    matchedFunction: matched,
  };
}
