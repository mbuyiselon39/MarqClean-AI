/**
 * "Excel for Data Analysis" — a 10-lesson learning track.
 *
 * Every formula, worked example and calculated result in this file is
 * arithmetic that has been computed and checked, not copied from a slide.
 * Where a statement describes Excel *behaviour* (argument order, blank-cell
 * handling, version availability) it reflects documented Microsoft behaviour;
 * the version column in `docs/excel-learning-verification.md` records which
 * release each claim was checked against.
 *
 * Formula argument separator
 * --------------------------
 * Syntax and examples are authored with a semicolon, matching the existing
 * MarqClean AI function reference (`functions.ts`) and the default list
 * separator in many Windows regional settings, including a large share of
 * South African corporate installs. The reader can switch the whole site to
 * commas with the separator control on every lesson; `renderFormula` performs
 * that substitution at render time.
 */

export type XlExample = { label: string; formula: string; result: string };

export type XlFunctionDoc = {
  name: string;
  purpose: string;
  syntax: string;
  examples: XlExample[];
  keyPoints: string[];
  mistakes: string[];
  version?: string;
  category: CheatCategory;
};

export type XlTable = {
  caption?: string;
  headers: string[];
  rows: string[][];
  /** Indexes of columns that hold the "answer" and should be emphasised. */
  resultColumns?: number[];
};

export type XlDiagram = "etl" | "lookup" | "merge-append" | "chart-choice";

export type XlSection =
  | { kind: "prose"; id: string; heading: string; body: string[] }
  | { kind: "functions"; id: string; heading: string; intro?: string; functions: XlFunctionDoc[] }
  | { kind: "table"; id: string; heading: string; table: XlTable; note?: string }
  | { kind: "steps"; id: string; heading: string; steps: { title: string; detail: string }[] }
  | { kind: "diagram"; id: string; heading: string; diagram: XlDiagram; caption: string }
  | { kind: "callout"; id: string; heading: string; points: string[] }
  | { kind: "practice"; id: string; heading: string; task: string; answer: string };

export type XlQuizQ = { q: string; options: string[]; answer: number; explain: string };

export type XlWorkbookSheet = { name: string; headers: string[]; rows: (string | number)[][] };

export type XlLesson = {
  slug: string;
  n: number;
  title: string;
  metaTitle: string;
  metaDescription: string;
  summary: string;
  level: "Beginner" | "Intermediate" | "Advanced";
  minutes: number;
  objectives: string[];
  sections: XlSection[];
  quiz: XlQuizQ[];
  workbookSheets: XlWorkbookSheet[];
  tool: { label: string; path: string; note: string };
  faq: { q: string; a: string }[];
};

export type CheatCategory =
  | "Basics"
  | "Basic functions"
  | "Conditional functions"
  | "Conditional aggregation"
  | "Lookup functions"
  | "Text functions"
  | "Data-cleaning tools"
  | "PivotTables"
  | "Charts"
  | "Power Query";

export const CHEAT_CATEGORIES: CheatCategory[] = [
  "Basic functions",
  "Conditional functions",
  "Conditional aggregation",
  "Lookup functions",
  "Text functions",
  "Data-cleaning tools",
  "PivotTables",
  "Charts",
  "Power Query",
];

/** Swap the list separator for readers whose Excel expects commas. */
export function renderFormula(source: string, separator: ";" | ","): string {
  return separator === ";" ? source : source.split(";").join(",");
}

/**
 * The shared practice dataset. Ten sales rows, deliberately small enough to
 * check by hand and reused by every lesson so learners build one mental model
 * instead of ten.
 */
export const SALES_HEADERS = ["Region", "Product", "Rep", "Units", "Unit Price", "Sales"];

export const SALES_ROWS: (string | number)[][] = [
  ["North", "Widget", "Thabo", 10, 250, 2500],
  ["South", "Gadget", "Lerato", 6, 400, 2400],
  ["East", "Widget", "Sipho", 14, 250, 3500],
  ["West", "Gadget", "Naledi", 8, 400, 3200],
  ["North", "Gadget", "Thabo", 5, 400, 2000],
  ["South", "Widget", "Lerato", 20, 250, 5000],
  ["East", "Widget", "Sipho", 3, 250, 750],
  ["West", "Widget", "Naledi", 12, 250, 3000],
  ["North", "Gadget", "Kagiso", 9, 400, 3600],
  ["South", "Gadget", "Lerato", 7, 400, 2800],
];

/** Two blank amounts: the "missing values" example used in Lesson 10. */
export const PAYMENTS_HEADERS = ["Client", "Amount"];

export const PAYMENTS_ROWS: (string | number)[][] = [
  ["Alpha", 25000],
  ["Beta", ""],
  ["Gamma", 18000],
  ["Delta", ""],
];

export const PRICES_HEADERS = ["Product", "Code", "Unit Price"];

export const PRICES_ROWS: (string | number)[][] = [
  ["Widget", "W-100", 250],
  ["Gadget", "G-200", 400],
  ["Sensor", "S-300", 150],
];

const SALES_TABLE: XlTable = {
  caption: "Sheet 'Sales' — the practice dataset used across this course. Data occupies A1:F11.",
  headers: ["Row", "A · Region", "B · Product", "C · Rep", "D · Units", "E · Unit Price", "F · Sales"],
  rows: SALES_ROWS.map((row, i) => [String(i + 2), ...(row as (string | number)[]).map(String)]),
  resultColumns: [6],
};

const LESSON_1: XlLesson = {
  slug: "excel-basics",
  n: 1,
  title: "Excel Basics: Workbooks, Cells, Ranges and References",
  metaTitle: "Excel Basics for Data Analysis | Cells, Ranges & References",
  metaDescription:
    "Learn how Excel workbooks, worksheets, cells and ranges work, and how relative, absolute and mixed references change a formula when you copy it.",
  summary:
    "Before a single function, get the mental model right: what a workbook is, how a cell address is built, what a range means, and why the $ sign decides whether a formula survives being copied.",
  level: "Beginner",
  minutes: 12,
  objectives: [
    "Explain the difference between a workbook and a worksheet",
    "Read and write a cell address and a range reference",
    "Tell a formula apart from a function",
    "Predict what happens to a reference when you copy a formula",
  ],
  sections: [
    {
      kind: "prose",
      id: "what-excel-is",
      heading: "What Excel actually is",
      body: [
        "Excel is a spreadsheet application: a grid of cells that can each hold a value, a piece of text, or a formula that calculates from other cells. Its real power is not the grid — it is recalculation. Change one input cell and every formula that depends on it updates immediately.",
        "For data analysis that matters. You keep one copy of the source data, write your calculations once, and re-run the whole analysis by replacing the source. That is the difference between a spreadsheet and a calculator.",
        "Excel is also where most business data ends up. Finance packs, CRM exports, bank statements, payroll registers and marketing lists all arrive as .xlsx or .csv files. The skills in this course are aimed squarely at turning those files into something you can report on.",
      ],
    },
    {
      kind: "table",
      id: "workbook-vs-worksheet",
      heading: "Workbook versus worksheet",
      table: {
        caption: "The two levels people most often mix up.",
        headers: ["Term", "What it is", "How you see it", "Analogy"],
        rows: [
          [
            "Workbook",
            "The file itself — one .xlsx (or .xlsb, .xlsm, .csv) document.",
            "The file name in the title bar, e.g. Q1-Sales.xlsx",
            "The ring binder",
          ],
          [
            "Worksheet",
            "One tab inside the workbook, holding its own grid.",
            "The tabs along the bottom: Sales, Prices, Summary",
            "A divider tab in the binder",
          ],
          [
            "Cell",
            "One intersection of a column and a row. Holds one value.",
            "The box you click, e.g. F2",
            "One line on the page",
          ],
          [
            "Range",
            "A rectangular block of cells, written first:last.",
            "A highlighted block, e.g. F2:F11",
            "A paragraph of lines",
          ],
        ],
      },
      note: "A CSV file is a single worksheet with no formatting, no formulas and no other tabs. Opening a CSV in Excel and saving it as .xlsx is a common first step before any real analysis.",
    },
    {
      kind: "prose",
      id: "rows-columns",
      heading: "Rows, columns and cell addresses",
      body: [
        "Columns are lettered A, B, C … across the top. Rows are numbered 1, 2, 3 … down the left side. A cell address is the column letter followed by the row number, so the cell in column F on row 2 is F2.",
        "In the practice dataset for this course, row 1 holds the column headings and the data starts on row 2. That is the convention every function in this course assumes: F2:F11 means 'the Sales column, ten rows of data, headings excluded'.",
        "A range is written as the top-left cell, a colon, then the bottom-right cell. F2:F11 is a single column of ten cells. A2:F11 is the whole block — six columns by ten rows, sixty cells. A range used inside a function is often called an array or a reference.",
      ],
    },
    {
      kind: "table",
      id: "formula-vs-function",
      heading: "Formula versus function",
      table: {
        caption: "A function is a built-in building block. A formula is anything you type after the equals sign.",
        headers: ["", "Formula", "Function"],
        rows: [
          ["Definition", "Any expression that starts with = and returns a value.", "A named, built-in operation Excel already knows how to do."],
          ["Examples", "=F2*E2", "SUM, AVERAGE, IF, XLOOKUP"],
          ["Can contain the other?", "Yes — a formula usually calls one or more functions.", "No — a function on its own is not a formula until you type = and give it arguments."],
          ["Written as", "=F2*E2", "=SUM(F2:F11) — the function SUM inside a formula"],
        ],
      },
      note: "Every function takes arguments in a fixed order. Getting the order wrong is the single most common cause of a wrong-but-plausible answer, which is why Lesson 2 onwards shows the syntax before the example.",
    },
    {
      kind: "prose",
      id: "references",
      heading: "Cell references: relative, absolute and mixed",
      body: [
        "When you copy a formula, Excel rewrites the references inside it. The $ sign is how you tell it not to.",
        "A relative reference has no $ and moves with the formula. Copy =F2*E2 down one row and it becomes =F3*E3. That is exactly what you want for a calculated column: each row uses its own row's numbers.",
        "An absolute reference has $ on both parts, like $F$2. It never moves, no matter where you copy the formula. Use it when every row must point at the same cell — a VAT rate, a target, a single exchange rate.",
        "A mixed reference has $ on one part only. $F2 locks the column but lets the row move; F$2 locks the row but lets the column move. Mixed references are what make a multiplication table or a cross-tab work.",
      ],
    },
    {
      kind: "table",
      id: "reference-copied",
      heading: "What each reference becomes when copied one row down",
      table: {
        caption: "Start cell F2, formula copied to F3.",
        headers: ["Reference in F2", "Type", "Becomes in F3", "Use it when"],
        rows: [
          ["E2", "Relative", "E3", "Each row should use its own row."],
          ["$E$2", "Absolute", "$E$2", "Every row should use the same cell."],
          ["$E2", "Mixed (column locked)", "$E3", "You will copy across columns but not rows."],
          ["E$2", "Mixed (row locked)", "E$2", "You will copy down rows but not across columns."],
        ],
        resultColumns: [2],
      },
      note: "Shortcut: with the cursor on a reference in the formula bar, press F4 (Windows) or Cmd+T then F4 (Mac) to cycle through the four forms.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "A workbook is the file; a worksheet is a tab inside it. A CSV is one unformatted worksheet.",
        "A cell address is column letter then row number: F2. A range is first:last: F2:F11.",
        "Headings live in row 1, so data ranges normally start at row 2.",
        "A formula starts with =. A function is a built-in name used inside a formula.",
        "$ stops a reference from moving when you copy. No $ means it moves.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Including the heading row in a numeric range, so SUM silently ignores the text and COUNT is one short.",
        "Forgetting the $ before copying a formula that points at a single rate cell — every row ends up pointing somewhere different.",
        "Typing a range with a comma instead of a colon: F2,F11 is two separate cells, not a range.",
        "Assuming a CSV keeps formulas, colours or extra tabs. It does not — save as .xlsx first if you need them.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "Open the practice workbook for this lesson. On the Sales sheet, in cell G2 type a formula that multiplies Units by Unit Price without using the value already in column F. Then copy it down to G11 and confirm every value matches column F. Next, put 0.15 into I1, and in H2 write a formula that returns 15% of the row's Sales using an absolute reference to $I$1. Copy it down and check that no row has drifted onto the wrong rate cell.",
      answer: "G2: =D2*E2  (copies to =D3*E3 and so on). H2: =F2*$I$1  (copies to =F3*$I$1 — the rate cell stays fixed). If H3 had shown 0 instead of a number, you almost certainly wrote =F2*I1 without the $ signs.",
    },
  ],
  quiz: [
    {
      q: "A workbook contains five tabs. How many worksheets does it contain?",
      options: ["One", "Five", "Ten", "It depends on how many cells are filled"],
      answer: 1,
      explain: "Each tab is one worksheet, so five tabs means five worksheets inside the single workbook file.",
    },
    {
      q: "What does the range F2:F11 refer to?",
      options: ["Two cells, F2 and F11", "Eleven cells in column F", "Ten cells in column F, from row 2 to row 11", "The block from F2 to F11 including column E"],
      answer: 2,
      explain: "F2:F11 is every cell in column F from row 2 down to row 11 — ten cells. It does not include column E, and it is not just the two endpoints.",
    },
    {
      q: "You copy the formula =$F$2*E2 from row 2 to row 5. What does it become?",
      options: ["=$F$5*E5", "=$F$2*E5", "=F5*E5", "=$F$2*E2"],
      answer: 1,
      explain: "The $ signs lock F2, so it stays $F$2. E2 has no $ signs, so it moves to E5.",
    },
    {
      q: "Which of these is a function rather than a formula?",
      options: ["=F2*E2", "SUM", "=SUM(F2:F11)", "F2:F11"],
      answer: 1,
      explain: "SUM on its own is the function name. =SUM(F2:F11) is a formula that uses that function. =F2*E2 is a formula using no function at all.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
    { name: "Payments", headers: PAYMENTS_HEADERS, rows: PAYMENTS_ROWS },
  ],
  tool: {
    label: "Open the Quick Data & CSV Cleaner",
    path: "/",
    note: "Convert a CSV into a formatted Excel workbook before you start building formulas on it.",
  },
  faq: [
    {
      q: "Is a CSV file the same as an Excel file?",
      a: "No. A CSV is plain text with one value per line and a delimiter between columns. It has no formulas, no formatting, no multiple tabs and no data types. Excel opens it, but saving as .xlsx is what preserves anything you build on top of it.",
    },
    {
      q: "Why does my formula show ### instead of a number?",
      a: "The column is too narrow to display the result. Widen the column by dragging its border in the header row. The underlying value is correct — only the display is clipped.",
    },
    {
      q: "Do I need to know the difference between relative and absolute references to use this course?",
      a: "Yes. Almost every worked example in Lessons 2 to 6 is designed to be copied down a column, and a missing $ is the most common reason a copied formula returns the wrong number rather than an error.",
    },
  ],
};

const LESSON_2: XlLesson = {
  slug: "basic-functions",
  n: 2,
  title: "Basic Functions: SUM, AVERAGE, MIN, MAX, ROUND and the COUNT family",
  metaTitle: "Excel SUM, AVERAGE, COUNT & COUNTA Explained With Examples",
  metaDescription:
    "SUM, AVERAGE, MIN, MAX, ROUND, COUNT, COUNTA and COUNTBLANK with syntax, worked results and the blank-cell behaviour that changes your totals.",
  summary:
    "The eight functions behind almost every summary row. The important part is not the syntax — it is how each one treats blank cells, text and errors differently.",
  level: "Beginner",
  minutes: 14,
  objectives: [
    "Total, average, minimise and maximise a numeric column",
    "Round a result for presentation without changing the stored value",
    "Choose correctly between COUNT, COUNTA and COUNTBLANK",
    "Predict what AVERAGE does with a blank cell in the range",
  ],
  sections: [
    {
      kind: "table",
      id: "the-data",
      heading: "The practice data",
      table: SALES_TABLE,
      note: "Every result below is calculated from this exact table, so you can reproduce it in the practice workbook.",
    },
    {
      kind: "functions",
      id: "aggregation",
      heading: "Adding and summarising",
      intro: "These four ignore text and blank cells entirely. If a cell holds text in the middle of a numeric column, they skip it without warning — which is convenient and dangerous at the same time.",
      functions: [
        {
          name: "SUM",
          category: "Basic functions",
          purpose: "Adds every number in the arguments. The foundation of every total row.",
          syntax: "=SUM(number1;[number2];...)",
          examples: [
            { label: "Total sales", formula: "=SUM(F2:F11)", result: "28750" },
            { label: "Total units", formula: "=SUM(D2:D11)", result: "94" },
            { label: "Two ranges at once", formula: "=SUM(D2:D11;F2:F11)", result: "28844" },
          ],
          keyPoints: [
            "Blank cells and text are ignored, not treated as zero.",
            "Select a block of numbers and press Alt + = to insert SUM automatically.",
            "A range and an individual cell can be mixed in one call.",
          ],
          mistakes: [
            "Numbers stored as text are skipped, so the total is quietly too low. Check with COUNT versus COUNTA.",
            "Selecting F2:F10 instead of F2:F11 drops the last row — an off-by-one that is easy to miss on a 10,000-row file.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "AVERAGE",
          category: "Basic functions",
          purpose: "The arithmetic mean: the sum divided by how many numbers were found.",
          syntax: "=AVERAGE(number1;[number2];...)",
          examples: [
            { label: "Average sale", formula: "=AVERAGE(F2:F11)", result: "2875" },
            { label: "Average units", formula: "=AVERAGE(D2:D11)", result: "9.4" },
            { label: "Average of the two filled Payments amounts", formula: "=AVERAGE(B2:B5)", result: "21500" },
          ],
          keyPoints: [
            "The divisor is the count of numeric cells only. Blank cells are excluded from both the sum and the count.",
            "A cell containing 0 is a number and is included. That is the difference between 'no value' and 'zero'.",
            "AVERAGEA also counts text as 0 and TRUE as 1, which almost never gives the mean you want.",
          ],
          mistakes: [
            "Assuming blanks count as zero. =AVERAGE(B2:B5) on the Payments sheet returns 21500, not 10750.",
            "Averaging a column that contains a subtotal row, which double-counts part of the data.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "MIN",
          category: "Basic functions",
          purpose: "The smallest number in the range.",
          syntax: "=MIN(number1;[number2];...)",
          examples: [
            { label: "Smallest sale", formula: "=MIN(F2:F11)", result: "750" },
            { label: "Fewest units in one order", formula: "=MIN(D2:D11)", result: "3" },
          ],
          keyPoints: [
            "Ignores text and blanks, like SUM and AVERAGE.",
            "A blank cell does not make MIN return 0 — it is skipped.",
          ],
          mistakes: ["Reading a 0 result as 'no data'. Zero is a real value; use COUNT to check whether the range is empty."],
          version: "All versions of Excel.",
        },
        {
          name: "MAX",
          category: "Basic functions",
          purpose: "The largest number in the range.",
          syntax: "=MAX(number1;[number2];...)",
          examples: [
            { label: "Largest sale", formula: "=MAX(F2:F11)", result: "5000" },
            { label: "Most units in one order", formula: "=MAX(D2:D11)", result: "20" },
          ],
          keyPoints: ["The quickest way to spot an outlier before you chart anything.", "Pair MIN and MAX to get the range: =MAX(F2:F11)-MIN(F2:F11) returns 4250."],
          mistakes: ["Assuming MAX returns the most frequent value. That is MODE, not MAX."],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "functions",
      id: "rounding",
      heading: "Rounding for presentation",
      functions: [
        {
          name: "ROUND",
          category: "Basic functions",
          purpose: "Rounds a number to a set number of digits. Positive digits round decimals, negative digits round to tens, hundreds and beyond.",
          syntax: "=ROUND(number;num_digits)",
          examples: [
            { label: "Two decimals", formula: "=ROUND(28750/7;2)", result: "4107.14" },
            { label: "Whole number", formula: "=ROUND(28750/7;0)", result: "4107" },
            { label: "Nearest hundred", formula: "=ROUND(28750/7;-2)", result: "4100" },
          ],
          keyPoints: [
            "ROUND changes the stored value. Cell formatting only changes what is displayed.",
            "That difference is why a formatted column can appear not to add up: the visible numbers are rounded, the hidden values are not.",
            "ROUNDUP always rounds away from zero and ROUNDDOWN always towards it, regardless of the digit.",
          ],
          mistakes: [
            "Rounding each row and then summing the rounded rows, which can differ from rounding the true total. Decide once which you need and be consistent in the report.",
            "Using formatting to 'fix' a rounding problem, then exporting the unrounded values.",
          ],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "functions",
      id: "counting",
      heading: "Counting: COUNT, COUNTA and COUNTBLANK",
      intro: "Three functions that answer three different questions. Choosing the wrong one is the reason so many data-quality checks pass when they should not.",
      functions: [
        {
          name: "COUNT",
          category: "Basic functions",
          purpose: "Counts cells that contain a number. Nothing else.",
          syntax: "=COUNT(value1;[value2];...)",
          examples: [
            { label: "How many sales figures are present", formula: "=COUNT(F2:F11)", result: "10" },
            { label: "How many Payments amounts are filled", formula: "=COUNT(B2:B5)", result: "2" },
            { label: "How many region names are present", formula: "=COUNT(A2:A11)", result: "0" },
          ],
          keyPoints: [
            "Text, blanks and logical values are not counted.",
            "COUNT on a text column returns 0 — a fast way to confirm a numeric column has been imported as text.",
          ],
          mistakes: ["Using COUNT to check whether a column is complete, then concluding the column is empty because the values were imported as text."],
          version: "All versions of Excel.",
        },
        {
          name: "COUNTA",
          category: "Basic functions",
          purpose: "Counts cells that are not empty, whatever type they hold.",
          syntax: "=COUNTA(value1;[value2];...)",
          examples: [
            { label: "Filled rows on the Sales sheet", formula: "=COUNTA(C2:C11)", result: "10" },
            { label: "Filled client names", formula: "=COUNTA(A2:A5)", result: "4" },
          ],
          keyPoints: [
            "Counts numbers, text, errors and logical values alike.",
            "Also counts a cell whose formula returns an empty string. Those cells look blank but are not.",
          ],
          mistakes: [
            "Treating COUNTA as 'rows with data' after a formula fill. Use COUNTBLANK alongside it before trusting the number.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "COUNTBLANK",
          category: "Basic functions",
          purpose: "Counts cells that appear empty in a range.",
          syntax: "=COUNTBLANK(range)",
          examples: [
            { label: "Missing Payments amounts", formula: "=COUNTBLANK(B2:B5)", result: "2" },
            { label: "Missing sales figures", formula: "=COUNTBLANK(F2:F11)", result: "0" },
          ],
          keyPoints: [
            "Takes one range, not a list of values.",
            "Counts a cell containing a formula that returns an empty string, because it renders as blank.",
            "A cell containing 0 or a single space is not blank.",
          ],
          mistakes: ["Using it on a column cleaned with formulas and concluding the data is complete, when in fact every cell holds a formula returning nothing."],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "table",
      id: "count-comparison",
      heading: "The three counters side by side",
      table: {
        caption: "Applied to the Payments sheet: B2=25000, B3=blank, B4=18000, B5=blank.",
        headers: ["Function", "Counts", "Result on B2:B5", "Use it to answer"],
        rows: [
          ["COUNT", "Numbers only", "2", "How many values did I actually get?"],
          ["COUNTA", "Anything not empty", "4", "How many rows exist?"],
          ["COUNTBLANK", "Empty-looking cells", "2", "How many are missing?"],
          ["AVERAGE", "Sum of numbers ÷ count of numbers", "21500", "What is the typical value of what I have?"],
        ],
        resultColumns: [2],
      },
      note: "4 minus 2 equals 2. When COUNTA, COUNT and COUNTBLANK agree like that, the column is genuinely two-filled, two-missing — not two-filled, two-text.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "SUM, AVERAGE, MIN and MAX all skip blanks and text without telling you.",
        "AVERAGE divides by the number of numeric cells, not the size of the range.",
        "Zero is a value. A blank is not. They behave differently in every one of these functions.",
        "COUNT counts numbers, COUNTA counts anything non-empty, COUNTBLANK counts empties.",
        "ROUND changes the value. Cell formatting changes only the display.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "A total that is too low because a column was imported as text. Diagnose with COUNT versus COUNTA.",
        "An average that is too high or too low because blanks were assumed to be zero.",
        "A summary row sitting inside the range being summarised.",
        "Reporting rounded figures that no longer add up to the rounded total.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "On the Payments sheet of the practice workbook, build a small summary block: the number of clients, the number of amounts present, the number missing, the total of the amounts present and their average. Then answer in words what would happen to the average if Beta's blank were replaced with 0.",
      answer: "=COUNTA(A2:A5) → 4 clients. =COUNT(B2:B5) → 2 amounts. =COUNTBLANK(B2:B5) → 2 missing. =SUM(B2:B5) → 43000. =AVERAGE(B2:B5) → 21500. If Beta's blank became 0, the sum would still be 43000 but COUNT would become 3, so the average would fall to 14333.33 — a 7166.67 drop caused purely by how the missing value is represented.",
    },
  ],
  quiz: [
    {
      q: "=SUM(F2:F11) on the Sales sheet returns 28750. What does =COUNT(F2:F11) return?",
      options: ["0", "10", "11", "2875"],
      answer: 1,
      explain: "All ten cells in F2:F11 hold numbers, so COUNT returns 10. Eleven would include the heading row.",
    },
    {
      q: "On the Payments sheet B2=25000, B3 is blank, B4=18000, B5 is blank. What does =AVERAGE(B2:B5) return?",
      options: ["10750", "14333.33", "21500", "43000"],
      answer: 2,
      explain: "AVERAGE ignores the blanks entirely: (25000 + 18000) ÷ 2 = 21500. Dividing by 4 would give 10750, which is the mistake blanks invite.",
    },
    {
      q: "COUNT on a column returns 0 but COUNTA on the same column returns 500. What is the most likely cause?",
      options: ["The column has 500 blanks", "The values were imported as text, not numbers", "The column contains errors", "COUNT is broken in this version of Excel"],
      answer: 1,
      explain: "COUNT counts numbers only. A column full of text returns COUNT 0 and COUNTA 500. Fix it with Text to Columns or VALUE — see Lesson 6.",
    },
    {
      q: "=ROUND(4107.142857;-2) returns what?",
      options: ["4107.14", "4107", "4100", "4000"],
      answer: 2,
      explain: "A negative num_digits rounds to the left of the decimal point. -2 means the nearest hundred, so 4107.142857 becomes 4100.",
    },
    {
      q: "Which function counts a cell whose formula returns an empty string?",
      options: ["COUNTBLANK only", "COUNTA and COUNTBLANK both count it", "COUNTA only", "Neither"],
      answer: 1,
      explain: "A formula returning an empty string is not truly empty. COUNTA counts it as content and COUNTBLANK counts it as blank, which is exactly why the two can look contradictory.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
    { name: "Payments", headers: PAYMENTS_HEADERS, rows: PAYMENTS_ROWS },
  ],
  tool: {
    label: "Score my data quality",
    path: "/data-toolbox",
    note: "Count missing values, blanks and text-in-number columns across a whole file before you total anything.",
  },
  faq: [
    {
      q: "Why does AVERAGE ignore blank cells instead of treating them as zero?",
      a: "Because a blank means 'no observation', not 'an observation of zero'. Treating missing data as zero would pull every average down and hide the fact that data is absent. If zero is genuinely correct for your rows, type the 0.",
    },
    {
      q: "Should I use ROUND or number formatting?",
      a: "Use formatting when you want the full precision kept for further calculations, and ROUND when the rounded figure is the value you will report, invoice or reconcile. Never mix the two in one report.",
    },
    {
      q: "My SUM is smaller than the numbers I can see on screen. What happened?",
      a: "Some of the visible numbers are stored as text. Compare COUNT with COUNTA on the column: if COUNT is lower, that many values are text. Lesson 6 covers the Text to Columns fix.",
    },
  ],
};

const LESSON_3: XlLesson = {
  slug: "conditional-functions",
  n: 3,
  title: "Conditional Functions: IF, IFS, AND, OR and NOT",
  metaTitle: "Excel IF, IFS, AND, OR & NOT Explained With Worked Examples",
  metaDescription:
    "Use IF, IFS, AND, OR and NOT to flag exceptions and categorise rows. Syntax, nested versus flat logic, and the #N/A trap in IFS.",
  summary:
    "How Excel makes a decision. IF for two outcomes, IFS for several, and AND, OR and NOT to combine conditions without nesting five levels deep.",
  level: "Beginner",
  minutes: 14,
  objectives: [
    "Write a two-outcome test with IF",
    "Replace nested IF with a flat IFS",
    "Combine conditions with AND, OR and NOT",
    "Handle the case where no condition matches",
  ],
  sections: [
    {
      kind: "prose",
      id: "why",
      heading: "Why conditional logic matters in data analysis",
      body: [
        "Almost every analysis has a decision in it. Is this invoice above threshold? Is this record a duplicate? Does this region belong in the report? Conditional functions turn that judgement into a column you can filter, sort and total.",
        "The pattern is always the same: a test that is TRUE or FALSE, a value for TRUE, and a value for FALSE. Everything else is combining tests.",
      ],
    },
    {
      kind: "functions",
      id: "if-ifs",
      heading: "IF and IFS",
      functions: [
        {
          name: "IF",
          category: "Conditional functions",
          purpose: "Tests one condition and returns one value when it is true and another when it is false.",
          syntax: "=IF(logical_test;value_if_true;[value_if_false])",
          examples: [
            { label: "Flag sales above 3000", formula: '=IF(F2>3000;"Above target";"At or below")', result: '"At or below" for F2 = 2500' },
            { label: "Flag large orders", formula: '=IF(D2>10;"Large";"Standard")', result: '"Standard" for D2 = 10' },
            { label: "IF with a calculation", formula: "=IF(F2>3000;F2*0.05;0)", result: "0 for F2 = 2500" },
          ],
          keyPoints: [
            "Omitting value_if_false returns FALSE rather than a blank. Pass an empty string when you want nothing.",
            'Comparison text must be in double quotes. Numbers and cell references must not be.',
            "IF nests up to 64 levels, but beyond two or three the formula stops being readable. Use IFS instead.",
          ],
          mistakes: [
            "Using = when you mean comparison. Inside a test, = is comparison and there is no separate equality operator.",
            "Forgetting quotes around text results, which produces #NAME? because Excel reads the word as a defined name.",
            "Testing a floating-point result for exact equality. Compare with a tolerance instead.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "IFS",
          category: "Conditional functions",
          purpose: "Tests a list of conditions in order and returns the value paired with the first one that is TRUE. Replaces nested IF.",
          syntax: "=IFS(logical_test1;value_if_true1;[logical_test2;value_if_true2];...)",
          examples: [
            { label: "Three-tier grading", formula: '=IFS(F2>4000;"Tier A";F2>2500;"Tier B";TRUE;"Tier C")', result: '"Tier C" for F2 = 2500' },
            { label: "Same row, higher value", formula: '=IFS(F7>4000;"Tier A";F7>2500;"Tier B";TRUE;"Tier C")', result: '"Tier A" for F7 = 5000' },
          ],
          keyPoints: [
            "Conditions are evaluated left to right and it stops at the first TRUE.",
            "Accepts up to 127 test-and-value pairs.",
            "If no condition is TRUE and there is no catch-all, IFS returns #N/A. End the list with TRUE to give a default.",
          ],
          mistakes: [
            "Ordering conditions loosely. Because evaluation stops at the first TRUE, a broad test placed first swallows the specific ones.",
            "Leaving out the final TRUE pair and getting #N/A on rows that fall outside every band.",
          ],
          version: "Excel 2019, 2021, 2024 and Microsoft 365. Not in Excel 2016 or earlier — use nested IF there.",
        },
      ],
    },
    {
      kind: "table",
      id: "ifs-worked",
      heading: "Worked example: the tiering formula down the whole column",
      table: {
        caption: "Formula in G2 is =IFS(F2>4000;\"Tier A\";F2>2500;\"Tier B\";TRUE;\"Tier C\"), copied to G11.",
        headers: ["Row", "F · Sales", "G · Tier", "Which test fired"],
        rows: [
          ["2", "2500", "Tier C", "Neither — fell through to the TRUE catch-all"],
          ["3", "2400", "Tier C", "Catch-all"],
          ["4", "3500", "Tier B", "F4>2500"],
          ["5", "3200", "Tier B", "F5>2500"],
          ["6", "2000", "Tier C", "Catch-all"],
          ["7", "5000", "Tier A", "F7>4000"],
          ["8", "750", "Tier C", "Catch-all"],
          ["9", "3000", "Tier B", "F9>2500"],
          ["10", "3600", "Tier B", "F10>2500"],
          ["11", "2800", "Tier B", "F11>2500"],
        ],
        resultColumns: [2],
      },
      note: "Result: 1 Tier A, 5 Tier B, 4 Tier C — ten rows in total. Note that row 2 at exactly 2500 is Tier C, because 2500>2500 is FALSE. Boundary values are where tiering formulas go wrong.",
    },
    {
      kind: "functions",
      id: "combining",
      heading: "AND, OR and NOT",
      intro: "These three never return text or numbers. They return TRUE or FALSE, which you then feed into IF.",
      functions: [
        {
          name: "AND",
          category: "Conditional functions",
          purpose: "TRUE only when every condition is TRUE.",
          syntax: "=AND(logical1;[logical2];...)",
          examples: [
            { label: "North region selling Widgets", formula: '=IF(AND(A2="North";B2="Widget");"Priority";"")', result: '"Priority" — row 2 is the only match' },
            { label: "Used on row 5 (North, Gadget)", formula: '=IF(AND(A5="North";B5="Widget");"Priority";"")', result: '"" — product does not match' },
          ],
          keyPoints: ["One FALSE makes the whole thing FALSE.", "Takes up to 255 conditions."],
          mistakes: ["Nesting ANDs when one flat AND would do: =AND(a;b;c) is clearer than =AND(a;AND(b;c))."],
          version: "All versions of Excel.",
        },
        {
          name: "OR",
          category: "Conditional functions",
          purpose: "TRUE when at least one condition is TRUE.",
          syntax: "=OR(logical1;[logical2];...)",
          examples: [
            { label: "Flag the eastern and western regions", formula: '=IF(OR(A2="East";A2="West");"Review";"")', result: '"" for row 2 (North)' },
            { label: "Count across the column", formula: '=SUMPRODUCT(--(OR(A2:A11="East";A2:A11="West")))', result: "4 rows flagged" },
          ],
          keyPoints: ["One TRUE is enough.", "The direct way to test a value against a short list of alternatives."],
          mistakes: ["Writing =IF(A2=\"East\" OR A2=\"West\";...) — OR is a function with its own brackets, not an operator between tests."],
          version: "All versions of Excel.",
        },
        {
          name: "NOT",
          category: "Conditional functions",
          purpose: "Reverses a logical value. TRUE becomes FALSE and FALSE becomes TRUE.",
          syntax: "=NOT(logical)",
          examples: [
            { label: "Everything except North", formula: '=IF(NOT(A2="North");"Non-North";"North")', result: '"North" for row 2' },
            { label: "Across the column", formula: '=SUMPRODUCT(--NOT(A2:A11="North"))', result: "7 non-North rows" },
          ],
          keyPoints: ["Takes exactly one argument.", "Often clearer than writing the opposite comparison by hand."],
          mistakes: ["Applying NOT to a range and expecting a filtered list. NOT returns one value per call; use FILTER for lists."],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "IF answers two outcomes. IFS answers many, in order, stopping at the first TRUE.",
        "Always end an IFS with TRUE as the final test, or unmatched rows return #N/A.",
        "AND needs all conditions, OR needs one, NOT flips a single value.",
        "Text in a test or a result needs double quotes. Cell references and numbers do not.",
        "Order your IFS conditions from most specific to least specific.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Boundary values landing in the wrong tier because the test used > where >= was intended.",
        "An #N/A appearing on a handful of rows because the IFS has no catch-all.",
        "Using OR as an operator instead of a function.",
        "Nesting IF five deep when a flat IFS or a lookup table would be readable and maintainable.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "On the Sales sheet, add a column that reads 'Priority' when the region is North or South and the sale is above 2500, 'Watch' when the sale is 1000 or below, and 'Standard' otherwise. Count how many rows fall into each category.",
      answer: 'H2: =IF(AND(OR(A2="North";A2="South");F2>2500);"Priority";IF(F2<=1000;"Watch";"Standard"))  — or as a flat IFS: =IFS(AND(OR(A2="North";A2="South");F2>2500);"Priority";F2<=1000;"Watch";TRUE;"Standard"). Across the ten rows you should get 3 Priority, 1 Watch and 6 Standard.',
    },
  ],
  quiz: [
    {
      q: '=IFS(F2>4000;"Tier A";F2>2500;"Tier B") with F2 = 1000 returns what?',
      options: ['"Tier B"', "A blank cell", "#N/A", "FALSE"],
      answer: 2,
      explain: "No condition is TRUE and there is no catch-all, so IFS returns #N/A. Adding TRUE;\"Tier C\" as a final pair prevents it.",
    },
    {
      q: "Why does IFS evaluate its conditions in the order you write them?",
      options: ["It sorts them automatically", "It stops at the first TRUE and ignores the rest", "It evaluates all of them and picks the largest", "Order does not matter"],
      answer: 1,
      explain: "Evaluation is left to right and halts at the first TRUE. A broad condition written first will swallow the specific ones beneath it.",
    },
    {
      q: "=AND(TRUE;FALSE;TRUE) returns what?",
      options: ["TRUE", "FALSE", "2", "#VALUE!"],
      answer: 1,
      explain: "AND requires every argument to be TRUE. One FALSE makes the whole result FALSE.",
    },
    {
      q: 'What does =IF(A2="North";"Yes") return when A2 contains "South"?',
      options: ['"No"', "A blank cell", "FALSE", "0"],
      answer: 2,
      explain: "With value_if_false omitted, IF returns the logical value FALSE. Pass \"\" as the third argument if you want an empty-looking cell.",
    },
    {
      q: "IFS is not available in which Excel version?",
      options: ["Excel 2019", "Excel 2021", "Microsoft 365", "Excel 2016"],
      answer: 3,
      explain: "IFS arrived with Excel 2019 and Microsoft 365. Excel 2016 and earlier need nested IF instead.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
  ],
  tool: {
    label: "Validate my data against rules",
    path: "/data-toolbox",
    note: "Apply conditional rules across a whole file and export only the rows that fail.",
  },
  faq: [
    {
      q: "When should I use IFS instead of nested IF?",
      a: "Whenever you have three or more outcomes from a single value. IFS keeps the conditions on one level so a colleague can read them, and it avoids the bracket-counting errors that nested IF invites. If you must support Excel 2016 or earlier, keep the nested IF.",
    },
    {
      q: "Can I use IFS to look up a value from a table instead of hard-coding the bands?",
      a: "You can, but a lookup table with XLOOKUP or INDEX/MATCH is easier to maintain: you change the table, not the formula. Lesson 5 covers that. Hard-code bands with IFS only when the rules are fixed policy, such as tax brackets.",
    },
  ],
};

const LESSON_4: XlLesson = {
  slug: "sumif-countif-averageif",
  n: 4,
  title: "SUMIF, COUNTIF and AVERAGEIF: Conditional Totals",
  metaTitle: "SUMIF vs SUMIFS, COUNTIF & AVERAGEIF | Syntax and Examples",
  metaDescription:
    "SUMIF, COUNTIF, AVERAGEIF and their multi-condition S versions, with argument order, wildcards, comparison operators and verified results.",
  summary:
    "Totals that respect a condition. The three single-condition functions, their multi-condition S versions, and the argument-order trap that catches almost everyone once.",
  level: "Beginner",
  minutes: 15,
  objectives: [
    "Total, count and average rows that meet one condition",
    "Extend to several conditions with SUMIFS, COUNTIFS and AVERAGEIFS",
    "Use wildcards and comparison operators inside criteria",
    "Get the argument order right, and know when it changes",
  ],
  sections: [
    {
      kind: "callout",
      id: "argument-order",
      heading: "Read this first: the argument order changes",
      points: [
        "SUMIF puts the sum range last: =SUMIF(range;criteria;[sum_range]).",
        "SUMIFS puts the sum range first: =SUMIFS(sum_range;criteria_range1;criteria1;...).",
        "COUNTIF takes only range and criteria. COUNTIFS takes only pairs — there is nothing to sum.",
        "This inconsistency is not a typo in these notes. It is how the functions are defined.",
      ],
    },
    {
      kind: "functions",
      id: "single-condition",
      heading: "One condition",
      functions: [
        {
          name: "SUMIF",
          category: "Conditional aggregation",
          purpose: "Adds the cells in a sum range whose matching row meets one condition.",
          syntax: "=SUMIF(range;criteria;[sum_range])",
          examples: [
            { label: "Total sales for North", formula: '=SUMIF(A2:A11;"North";F2:F11)', result: "8100" },
            { label: "Total sales for South", formula: '=SUMIF(A2:A11;"South";F2:F11)', result: "10200" },
            { label: "Total of everything that is not Widget", formula: '=SUMIF(B2:B11;"<>Widget";F2:F11)', result: "14000" },
            { label: "Total sales above 2500", formula: '=SUMIF(F2:F11;">2500")', result: "21100" },
          ],
          keyPoints: [
            "When sum_range is omitted, Excel sums the criteria range itself — that is the last example.",
            "Criteria text and operators go in quotes. Cell references do not: =SUMIF(A2:A11;D1;F2:F11).",
            "Criteria are not case-sensitive: 'north' and 'North' match the same rows.",
            "range and sum_range may be different sizes; Excel aligns them on the top-left cell.",
          ],
          mistakes: [
            "Writing =SUMIF(F2:F11;\"North\";A2:A11) — the ranges are the wrong way round and the result is 0.",
            "Quoting a cell reference, so =SUMIF(A2:A11;\"D1\";F2:F11) looks for the literal text D1.",
            "Expecting a subtotal to respect filters. It does not — use SUBTOTAL(109;...) for that.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "COUNTIF",
          category: "Conditional aggregation",
          purpose: "Counts the cells in one range that meet a condition.",
          syntax: "=COUNTIF(range;criteria)",
          examples: [
            { label: "Rows for North", formula: '=COUNTIF(A2:A11;"North")', result: "3" },
            { label: "Widget rows", formula: '=COUNTIF(B2:B11;"Widget")', result: "5" },
            { label: "Orders over ten units", formula: '=COUNTIF(D2:D11;">10")', result: "3" },
            { label: "Regions ending in 'ast'", formula: '=COUNTIF(A2:A11;"*ast")', result: "2" },
          ],
          keyPoints: [
            "Takes exactly two arguments — there is no sum range.",
            "Works on text, numbers and dates in the same way.",
            "Counting a value against its own column is the standard duplicate check: =COUNTIF(A$2:A$11;A2) returns how many times that value appears.",
          ],
          mistakes: ["Supplying a third argument and getting an error, because COUNTIF only accepts two."],
          version: "All versions of Excel.",
        },
        {
          name: "AVERAGEIF",
          category: "Conditional aggregation",
          purpose: "Averages the cells in an average range whose matching row meets one condition.",
          syntax: "=AVERAGEIF(range;criteria;[average_range])",
          examples: [
            { label: "Average sale for North", formula: '=AVERAGEIF(A2:A11;"North";F2:F11)', result: "2700" },
            { label: "Average sale for East", formula: '=AVERAGEIF(A2:A11;"East";F2:F11)', result: "2125" },
            { label: "Average sale for West", formula: '=AVERAGEIF(A2:A11;"West";F2:F11)', result: "3100" },
          ],
          keyPoints: [
            "Blank cells inside average_range are skipped, exactly as AVERAGE skips them.",
            "If the criteria match rows but every matching average cell is blank, the result is #DIV/0!.",
            "Follows the same argument order as SUMIF, not SUMIFS.",
          ],
          mistakes: ["Assuming a #DIV/0! means the formula is broken. It means nothing matched, or every match was blank."],
          version: "Excel 2007 and later.",
        },
      ],
    },
    {
      kind: "functions",
      id: "multi-condition",
      heading: "Several conditions",
      intro: "Add an S and the pattern changes: every condition becomes a range-and-criteria pair, and for SUMIFS and AVERAGEIFS the value range moves to the front.",
      functions: [
        {
          name: "SUMIFS",
          category: "Conditional aggregation",
          purpose: "Adds cells that meet every condition in a list of range-and-criteria pairs.",
          syntax: "=SUMIFS(sum_range;criteria_range1;criteria1;[criteria_range2;criteria2];...)",
          examples: [
            { label: "North Widgets only", formula: '=SUMIFS(F2:F11;A2:A11;"North";B2:B11;"Widget")', result: "2500" },
            { label: "South Gadgets only", formula: '=SUMIFS(F2:F11;A2:A11;"South";B2:B11;"Gadget")', result: "5200" },
            { label: "Sales above 2000 with more than 5 units", formula: '=SUMIFS(F2:F11;F2:F11;">2000";D2:D11;">5")', result: "26000" },
          ],
          keyPoints: [
            "All ranges must cover the same number of rows and columns, or Excel returns #VALUE!.",
            "Conditions are combined with AND. There is no OR inside SUMIFS — sum two SUMIFS results for that.",
            "Up to 127 range-and-criteria pairs.",
          ],
          mistakes: [
            "Carrying SUMIF's order across: =SUMIFS(A2:A11;\"North\";F2:F11) is invalid.",
            "Ranges of different lengths, which produces #VALUE! rather than a wrong number.",
          ],
          version: "Excel 2007 and later.",
        },
        {
          name: "COUNTIFS",
          category: "Conditional aggregation",
          purpose: "Counts rows that meet every condition in a list of range-and-criteria pairs.",
          syntax: "=COUNTIFS(criteria_range1;criteria1;[criteria_range2;criteria2];...)",
          examples: [
            { label: "South Gadgets", formula: '=COUNTIFS(A2:A11;"South";B2:B11;"Gadget")', result: "2" },
            { label: "Lerato's rows", formula: '=COUNTIFS(C2:C11;"Lerato")', result: "3" },
            { label: "Widget orders above 10 units", formula: '=COUNTIFS(B2:B11;"Widget";D2:D11;">10")', result: "3" },
          ],
          keyPoints: ["No sum range at all — only pairs.", "Empty criteria match truly empty cells only, not zeros."],
          mistakes: ["Passing a sum range out of habit from SUMIFS and shifting every argument one place to the right."],
          version: "Excel 2007 and later.",
        },
        {
          name: "AVERAGEIFS",
          category: "Conditional aggregation",
          purpose: "Averages cells that meet every condition in a list of range-and-criteria pairs.",
          syntax: "=AVERAGEIFS(average_range;criteria_range1;criteria1;[criteria_range2;criteria2];...)",
          examples: [
            { label: "Average sale for Widgets", formula: '=AVERAGEIFS(F2:F11;B2:B11;"Widget")', result: "2950" },
            { label: "Average sale for Gadgets", formula: '=AVERAGEIFS(F2:F11;B2:B11;"Gadget")', result: "2800" },
            { label: "Average North Gadget sale", formula: '=AVERAGEIFS(F2:F11;A2:A11;"North";B2:B11;"Gadget")', result: "2800" },
          ],
          keyPoints: ["average_range comes first, like SUMIFS.", "Blank cells in average_range are excluded from the mean."],
          mistakes: ["Writing =AVERAGEIFS(B2:B11;\"Widget\";F2:F11), the single-condition order, and getting an error."],
          version: "Excel 2007 and later.",
        },
      ],
    },
    {
      kind: "table",
      id: "comparison",
      heading: "Quick comparison table",
      table: {
        caption: "The whole family on one page. Copy this into your own notes.",
        headers: ["Function", "Conditions", "First argument", "Result", "Verified example"],
        rows: [
          ["SUMIF", "One", "range (criteria range)", "Sum", '=SUMIF(A2:A11;"North";F2:F11) → 8100'],
          ["COUNTIF", "One", "range", "Count", '=COUNTIF(A2:A11;"North") → 3'],
          ["AVERAGEIF", "One", "range (criteria range)", "Mean", '=AVERAGEIF(A2:A11;"North";F2:F11) → 2700'],
          ["SUMIFS", "Many (AND)", "sum_range", "Sum", '=SUMIFS(F2:F11;A2:A11;"North";B2:B11;"Widget") → 2500'],
          ["COUNTIFS", "Many (AND)", "criteria_range1", "Count", '=COUNTIFS(A2:A11;"South";B2:B11;"Gadget") → 2'],
          ["AVERAGEIFS", "Many (AND)", "average_range", "Mean", '=AVERAGEIFS(F2:F11;B2:B11;"Widget") → 2950'],
        ],
        resultColumns: [4],
      },
      note: "The rule to remember: with one condition the value range comes last and is optional. With several it comes first and is required — except for COUNTIFS, which never has one.",
    },
    {
      kind: "prose",
      id: "criteria",
      heading: "Writing criteria",
      body: [
        "A criterion is a small expression Excel evaluates against each cell in the range. Text goes in quotes. Comparison operators go in quotes with the value: \">100\", \"<=50\", \"<>Widget\". A cell reference goes in without quotes and can be concatenated with an operator using &.",
        "Two wildcards work inside text criteria. An asterisk * stands for any number of characters, so \"W*\" matches Widget and anything else starting with W. A question mark ? stands for exactly one character, so \"?idget\" matches Widget. A tilde ~ escapes either one when you need the literal character.",
        "Wildcards apply to text criteria only. Using * in a numeric criterion does not mean 'any number'.",
      ],
    },
    {
      kind: "table",
      id: "wildcards",
      heading: "Wildcard and operator reference",
      table: {
        caption: "All criteria forms used in this lesson, each with its verified result on the Sales sheet.",
        headers: ["Criteria", "Matches", "Example", "Result"],
        rows: [
          ['"North"', "Exactly North, case-insensitive", '=COUNTIF(A2:A11;"North")', "3"],
          ['"*ast"', "Any region ending in 'ast'", '=COUNTIF(A2:A11;"*ast")', "2"],
          ['"W*"', "Any product starting with W", '=COUNTIF(B2:B11;"W*")', "5"],
          ['"<>Widget"', "Anything except Widget", '=SUMIF(B2:B11;"<>Widget";F2:F11)', "14000"],
          ['">10"', "Numbers greater than 10", '=COUNTIF(D2:D11;">10")', "3"],
          ['">="&D1', "Greater than or equal to the value in D1", '=COUNTIF(D2:D11;">="&D1)', "depends on D1"],
        ],
        resultColumns: [3],
      },
      note: "Note the concatenation in the last row: the operator is inside quotes and the cell reference sits outside, joined by &. Writing \">=D1\" searches for the literal text '> = D1' and matches nothing.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "SUMIF and AVERAGEIF put the value range last and make it optional. SUMIFS and AVERAGEIFS put it first and require it.",
        "COUNTIF takes two arguments. COUNTIFS takes only pairs.",
        "All criteria in the S versions are ANDed together.",
        "Criteria are not case-sensitive, and * and ? act as wildcards in text criteria.",
        "Concatenate a cell reference to an operator with &, never inside the quotes.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Swapping the argument order between the singular and plural forms — the most common error in this family.",
        "Different-length ranges in SUMIFS, producing #VALUE!.",
        "Putting a cell reference inside the quotes so it is read as literal text.",
        "Expecting OR behaviour from SUMIFS and getting a smaller total than expected.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "Build a small summary block on the Sales sheet that reports, for each of the four regions, the number of orders, the total sales and the average sale. Then add one line showing total sales for Widget orders above 1000. Check your totals add back to 28750.",
      answer: 'For North: =COUNTIF(A2:A11;"North") → 3, =SUMIF(A2:A11;"North";F2:F11) → 8100, =AVERAGEIF(A2:A11;"North";F2:F11) → 2700. The same pattern gives South 3 / 10200 / 3400, East 2 / 4250 / 2125 and West 2 / 6200 / 3100. The four totals 8100 + 10200 + 4250 + 6200 equal 28750, matching =SUM(F2:F11). The Widget line: =SUMIFS(F2:F11;B2:B11;"Widget";F2:F11;">1000") → 14000.',
    },
  ],
  quiz: [
    {
      q: "Which of these is the correct SUMIFS syntax?",
      options: [
        "=SUMIFS(criteria_range;criteria;sum_range)",
        "=SUMIFS(sum_range;criteria_range1;criteria1;...)",
        "=SUMIFS(criteria_range1;criteria1;sum_range)",
        "=SUMIFS(sum_range;criteria1;criteria_range1)",
      ],
      answer: 1,
      explain: "SUMIFS takes the sum range first, then alternating criteria ranges and criteria. SUMIF is the one that puts its sum range last.",
    },
    {
      q: '=SUMIF(A2:A11;"North";F2:F11) on the Sales sheet returns what?',
      options: ["3", "2700", "8100", "10200"],
      answer: 2,
      explain: "The three North rows are 2500, 2000 and 3600, which total 8100. Three is the COUNTIF answer and 2700 is the AVERAGEIF answer.",
    },
    {
      q: "How do you sum rows matching region North OR region South in one SUMIFS?",
      options: [
        'Use "North;South" as the criterion',
        "It cannot be done — SUMIFS only ANDs conditions, so add two SUMIF results",
        "Use OR() inside SUMIFS",
        "Put both regions in one cell",
      ],
      answer: 1,
      explain: "Conditions inside SUMIFS are always ANDed. For OR, either add two single-condition results together or use an array criterion with SUMPRODUCT.",
    },
    {
      q: '=COUNTIF(A2:A11;">="&D1) — what is the & doing?',
      options: [
        "It is a typo",
        "It joins the operator text to the value in D1",
        "It means 'and'",
        "It makes the reference absolute",
      ],
      answer: 1,
      explain: "The operator must be inside quotes and the cell reference outside, so & concatenates them into one criterion string at runtime. Writing \">=D1\" searches for literal text.",
    },
    {
      q: '=COUNTIF(A2:A11;"*ast") returns 2. Which regions matched?',
      options: ["North and South", "East twice", "East and West", "North and East"],
      answer: 1,
      explain: "The pattern ends in 'ast'. East appears on rows 4 and 8, so it matches twice. West ends in 'est' and does not match.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
  ],
  tool: {
    label: "Sort, subtotal and format my data",
    path: "/excel-automation",
    note: "Group by a column and generate real subtotals instead of writing a conditional total for every category.",
  },
  faq: [
    {
      q: "Why does SUMIF let me leave out the sum range but SUMIFS does not?",
      a: "Because SUMIF has only one range argument, Excel can assume you mean to sum that same range when you omit the third argument. SUMIFS has no such ambiguity to resolve, so the sum range is mandatory and comes first.",
    },
    {
      q: "Can I use SUMIFS with dates?",
      a: 'Yes. Put the date in quotes in your locale format, or better, reference a cell holding a real date: =SUMIFS(F2:F11;E2:E11;">="&DATE(2024;1;1)). Referencing a cell avoids locale formatting problems entirely.',
    },
    {
      q: "When should I use a PivotTable instead of these functions?",
      a: "When you want every category at once rather than one named category. SUMIF needs the category written into the formula; a PivotTable lists all of them and updates when the data does. Lesson 7 covers that.",
    },
  ],
};

const LESSON_5: XlLesson = {
  slug: "lookup-functions",
  n: 5,
  title: "Lookup Functions: VLOOKUP, XLOOKUP, INDEX and MATCH",
  metaTitle: "VLOOKUP vs XLOOKUP vs INDEX MATCH | Worked Examples",
  metaDescription:
    "VLOOKUP, XLOOKUP, INDEX, MATCH and INDEX+MATCH explained with syntax, exact-match rules, left-lookup limits and verified results.",
  summary:
    "Pulling a value from another table. Four tools, one job, and a set of rules about direction, exact matching and what to do when nothing is found.",
  level: "Intermediate",
  minutes: 16,
  objectives: [
    "Look up a value in a reference table with VLOOKUP",
    "Understand why the fourth argument must be FALSE",
    "Use XLOOKUP where it is available and know where it is not",
    "Combine INDEX and MATCH to look up in any direction",
  ],
  sections: [
    {
      kind: "table",
      id: "the-table",
      heading: "The reference table",
      table: {
        caption: "Sheet 'Prices' — occupies A1:C4. Note that Product is in column A, Code in column B and Unit Price in column C.",
        headers: ["Row", "A · Product", "B · Code", "C · Unit Price"],
        rows: [
          ["2", "Widget", "W-100", "250"],
          ["3", "Gadget", "G-200", "400"],
          ["4", "Sensor", "S-300", "150"],
        ],
      },
    },
    {
      kind: "functions",
      id: "the-functions",
      heading: "The four lookup tools",
      functions: [
        {
          name: "VLOOKUP",
          category: "Lookup functions",
          purpose: "Searches down the first column of a table and returns a value from a column you count to the right.",
          syntax: "=VLOOKUP(lookup_value;table_array;col_index_num;[range_lookup])",
          examples: [
            { label: "Unit price of Sensor", formula: '=VLOOKUP("Sensor";A2:C4;3;FALSE)', result: "150" },
            { label: "Code of Widget", formula: '=VLOOKUP("Widget";A2:C4;2;FALSE)', result: "W-100" },
            { label: "A product that does not exist", formula: '=VLOOKUP("Bracket";A2:C4;3;FALSE)', result: "#N/A" },
          ],
          keyPoints: [
            "The lookup value must be in the leftmost column of table_array. VLOOKUP cannot look to the left.",
            "col_index_num counts from the left of table_array, not from column A of the sheet.",
            "Pass FALSE (or 0) for an exact match. The default, TRUE, assumes the table is sorted ascending and returns an approximate match.",
            "HLOOKUP is the same function reading across a row instead of down a column.",
          ],
          mistakes: [
            "Omitting the fourth argument and silently getting an approximate match. This is the single most damaging default in Excel.",
            "Inserting a column into the table and leaving col_index_num unchanged, so the formula now returns the neighbouring column.",
            "Trailing spaces in the lookup value causing #N/A even though the text looks identical. Wrap both sides in TRIM.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "XLOOKUP",
          category: "Lookup functions",
          purpose: "Searches any column or row for a value and returns from any other, in any direction, with built-in not-found handling.",
          syntax: "=XLOOKUP(lookup_value;lookup_array;return_array;[if_not_found];[match_mode];[search_mode])",
          examples: [
            { label: "Unit price of Sensor", formula: '=XLOOKUP("Sensor";A2:A4;C2:C4)', result: "150" },
            { label: "Code for a product — looking left", formula: '=XLOOKUP("Gadget";A2:A4;B2:B4)', result: "G-200" },
            { label: "Product from a code — also looking left", formula: '=XLOOKUP("G-200";B2:B4;A2:A4)', result: "Gadget" },
            { label: "With a friendly not-found message", formula: '=XLOOKUP("Bracket";A2:A4;C2:C4;"Not in price list")', result: '"Not in price list"' },
          ],
          keyPoints: [
            "Exact match by default. There is no dangerous approximate-match default to forget.",
            "lookup_array and return_array are separate, so they can sit in any order on the sheet.",
            "The fourth argument replaces the IFERROR wrapper most VLOOKUP formulas need.",
            "Can return several columns at once by giving return_array multiple columns; the result spills.",
          ],
          mistakes: [
            "Using it in a workbook that colleagues open in Excel 2019 or 2016, where it shows #NAME?.",
            "Giving lookup_array and return_array different lengths, which returns #VALUE!.",
          ],
          version: "Microsoft 365, Excel 2021 and Excel 2024. Not available in Excel 2019, 2016 or earlier.",
        },
        {
          name: "INDEX",
          category: "Lookup functions",
          purpose: "Returns the value at a given position in a range. On its own it is positional, not a search.",
          syntax: "=INDEX(array;row_num;[col_num])",
          examples: [
            { label: "Third product in the list", formula: "=INDEX(A2:A4;3)", result: "Sensor" },
            { label: "Price in the second row of the price column", formula: "=INDEX(C2:C4;2)", result: "400" },
            { label: "Value at row 2, column 3 of the block", formula: "=INDEX(A2:C4;2;3)", result: "400" },
          ],
          keyPoints: ["row_num counts within the array you gave it, not the worksheet.", "INDEX does not search. It is the retrieval half of a lookup, and MATCH supplies the position."],
          mistakes: ["Expecting INDEX alone to find something. Without MATCH it only knows positions you already have."],
          version: "All versions of Excel.",
        },
        {
          name: "MATCH",
          category: "Lookup functions",
          purpose: "Returns the position of a value within a range. It finds where, not what.",
          syntax: "=MATCH(lookup_value;lookup_array;[match_type])",
          examples: [
            { label: "Position of Gadget", formula: '=MATCH("Gadget";A2:A4;0)', result: "2" },
            { label: "Position of a code", formula: '=MATCH("S-300";B2:B4;0)', result: "3" },
            { label: "Something not in the list", formula: '=MATCH("Bracket";A2:A4;0)', result: "#N/A" },
          ],
          keyPoints: [
            "match_type 0 means exact match. Always pass it.",
            "1 (the default) assumes ascending order and finds the largest value less than or equal to the target.",
            "Returns a number — a position — which is exactly what INDEX needs.",
          ],
          mistakes: ["Omitting match_type and inheriting the same approximate-match default that makes VLOOKUP unreliable."],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "prose",
      id: "index-match",
      heading: "INDEX + MATCH together",
      body: [
        "MATCH finds the position; INDEX retrieves the value at that position. Together they do what VLOOKUP does without its two big restrictions — they can look to the left, and they do not break when you insert a column.",
        "Read the formula inside out. =INDEX(C2:C4;MATCH(\"Gadget\";A2:A4;0)) means: find where Gadget sits in A2:A4 — position 2 — then return the value two rows down in C2:C4 — which is 400.",
        "Because the two ranges are independent, you can reverse them. =INDEX(B2:B4;MATCH(\"Gadget\";A2:A4;0)) returns G-200. That is a left lookup, and VLOOKUP cannot do it: it would need Code to be the leftmost column of the table.",
      ],
    },
    {
      kind: "table",
      id: "left-lookup",
      heading: "Worked example: the left lookup VLOOKUP cannot do",
      table: {
        caption: "Given a product name in column A, return its code from column B — which sits to the right, so this direction works for VLOOKUP. Now reverse it.",
        headers: ["Goal", "Formula", "Result", "Why"],
        rows: [
          ["Code from product (leftmost first)", '=VLOOKUP("Gadget";A2:C4;2;FALSE)', "G-200", "Code is to the right of Product — VLOOKUP is happy."],
          ["Product from code (rightmost first)", '=VLOOKUP("G-200";A2:C4;1;FALSE)', "#N/A", "Product sits to the left of Code. VLOOKUP cannot look left."],
          ["Product from code, INDEX+MATCH", '=INDEX(A2:A4;MATCH("G-200";B2:B4;0))', "Gadget", "MATCH searches column B and INDEX returns from column A."],
          ["Product from code, XLOOKUP", '=XLOOKUP("G-200";B2:B4;A2:A4)', "Gadget", "The arrays are named independently, so direction is irrelevant."],
        ],
        resultColumns: [2],
      },
      note: "The practical fix for VLOOKUP users is to reorder the source columns so the lookup key is leftmost. That works, but it means the formula depends on column order — the thing INDEX+MATCH and XLOOKUP free you from.",
    },
    {
      kind: "diagram",
      id: "workflow",
      heading: "The lookup workflow",
      diagram: "lookup",
      caption: "Every lookup, in any of the three functions, follows the same five steps. Most lookup errors happen at step 2 or step 5.",
    },
    {
      kind: "table",
      id: "comparison",
      heading: "Which one should I use?",
      table: {
        headers: ["", "VLOOKUP", "INDEX + MATCH", "XLOOKUP"],
        rows: [
          ["Available in", "Every Excel version", "Every Excel version", "Microsoft 365, Excel 2021 and 2024"],
          ["Lookup direction", "Right only", "Any direction", "Any direction"],
          ["Default match", "Approximate — you must pass FALSE", "Approximate — you must pass 0", "Exact"],
          ["Breaks when a column is inserted", "Yes — col_index_num is a count", "No — ranges are explicit", "No — ranges are explicit"],
          ["Built-in not-found value", "No — wrap in IFERROR", "No — wrap in IFERROR", "Yes — fourth argument"],
          ["Can return several columns at once", "No", "Not without extra work", "Yes"],
          ["Best for", "Simple right-hand lookups in shared files", "Any direction, any version", "New work on Microsoft 365 or Excel 2021+"],
        ],
      },
      note: "If the workbook will be opened by someone on Excel 2019 or earlier, do not use XLOOKUP. It will display #NAME? for them even though it works on your machine.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "VLOOKUP searches down the leftmost column of the table and counts columns to the right.",
        "Always pass FALSE to VLOOKUP and 0 to MATCH. The approximate-match default causes silent wrong answers.",
        "INDEX retrieves by position, MATCH finds the position. Together they look up in any direction.",
        "XLOOKUP defaults to an exact match and handles not-found values directly.",
        "XLOOKUP is unavailable in Excel 2019 and earlier — check your audience before using it.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Forgetting FALSE, so Excel returns a plausible but wrong approximate match.",
        "Trailing or double spaces in the lookup value producing #N/A against text that looks identical.",
        "A hard-coded col_index_num that silently points at the wrong column after someone inserts a column.",
        "Comparing a number stored as text against a real number. They never match, however similar they look.",
        "Using a whole-column reference in a very large workbook where a bounded range would calculate far faster.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "On the Sales sheet, add a column that pulls the correct Unit Price from the Prices sheet using the Product name in column B, and confirm it matches column E on every row. Then add a column that returns the product Code for the same row. Finally, change one Product name to include a trailing space and observe what happens.",
      answer: 'Price: =VLOOKUP(B2;Prices!A2:C4;3;FALSE) → matches column E on all ten rows. Code: =INDEX(Prices!B2:B4;MATCH(B2;Prices!A2:A4;0)) or =XLOOKUP(B2;Prices!A2:A4;Prices!B2:B4). With a trailing space the lookup returns #N/A, because Excel compares characters exactly. Fix it with =TRIM(B2) on both sides of the comparison.',
    },
  ],
  quiz: [
    {
      q: "What does the fourth argument of VLOOKUP do when you leave it out?",
      options: [
        "It defaults to an exact match",
        "It defaults to an approximate match and assumes the table is sorted ascending",
        "It returns #N/A",
        "It returns the first row of the table",
      ],
      answer: 1,
      explain: "The omitted fourth argument defaults to TRUE — approximate match. On an unsorted table that returns a wrong value with no error, which is why you should always pass FALSE.",
    },
    {
      q: '=INDEX(A2:A4;MATCH("G-200";B2:B4;0)) on the Prices sheet returns what?',
      options: ["G-200", "Gadget", "400", "#N/A"],
      answer: 1,
      explain: 'MATCH finds "G-200" at position 2 of B2:B4, then INDEX returns position 2 of A2:A4, which is "Gadget". This is a left lookup, which VLOOKUP cannot do.',
    },
    {
      q: "A colleague opens your workbook in Excel 2019 and every XLOOKUP shows #NAME?. What happened?",
      options: ["The file is corrupt", "XLOOKUP is not available in Excel 2019", "The formula is misspelled", "They need to enable macros"],
      answer: 1,
      explain: "XLOOKUP exists only in Microsoft 365, Excel 2021 and Excel 2024. In Excel 2019 the name is not recognised, so it displays #NAME?. Use INDEX+MATCH for files that must work everywhere.",
    },
    {
      q: '=VLOOKUP("Widget";A2:C4;2;FALSE) returns what?',
      options: ["Widget", "W-100", "250", "2"],
      answer: 1,
      explain: "Widget is found in the first row of the table and column index 2 of A2:C4 is the Code column, so the result is W-100.",
    },
    {
      q: "Your VLOOKUP returns #N/A but the value is visibly in the table. What is the most likely cause?",
      options: [
        "The table is too small",
        "A hidden space, or a number stored as text, in either the lookup value or the table",
        "VLOOKUP is case-sensitive",
        "You used FALSE",
      ],
      answer: 1,
      explain: "Excel compares exactly. A trailing space or a text-formatted number looks identical on screen but will not match. TRIM the text or convert the numbers, and note that VLOOKUP itself is not case-sensitive.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
    { name: "Prices", headers: PRICES_HEADERS, rows: PRICES_ROWS },
  ],
  tool: {
    label: "Fuzzy match my records",
    path: "/fuzzy-name-matching",
    note: "When the two tables do not share exact values, exact lookups fail. Fuzzy matching ranks near-matches for human review.",
  },
  faq: [
    {
      q: "Should I still learn VLOOKUP if XLOOKUP is better?",
      a: "Yes. Most workbooks in circulation were written before XLOOKUP existed, and a large share of business users are still on Excel 2019 or 2016 where it is the only option. You will spend more time reading VLOOKUP than writing it.",
    },
    {
      q: "Is INDEX+MATCH slower than VLOOKUP?",
      a: "On typical business-sized tables the difference is not noticeable. Both scan the lookup column. The reasons to prefer INDEX+MATCH are direction, resilience to inserted columns and explicit ranges — not speed.",
    },
    {
      q: "What should I do when the values do not match exactly?",
      a: "Clean first, then look up. TRIM removes stray spaces and converting text-numbers to real numbers resolves most failures. Where the names genuinely differ — 'Acme Ltd' against 'ACME Limited' — no formula will resolve it reliably. That is what the fuzzy matching tool is for.",
    },
  ],
};

/** Dirty rows used by the Data Cleaning lesson. */
export const DIRTY_HEADERS = ["Client Name", "Email", "Amount"];

export const DIRTY_ROWS: (string | number)[][] = [
  ["  Alpha Pty Ltd ", " INFO@ALPHA.CO.ZA ", 12500],
  ["Beta Holdings  ", "accounts@beta.co.za", 8200.5],
  ["Alpha Pty Ltd", "info@alpha.co.za", 12500],
  ["  Gamma CC", "", 3100],
  ["alpha pty ltd", "info@alpha.co.za", 12500],
];

/** Monthly revenue series used by the Charts lesson. */
export const MONTHLY_HEADERS = ["Month", "Revenue"];

export const MONTHLY_ROWS: (string | number)[][] = [
  ["Jan", 18500],
  ["Feb", 21200],
  ["Mar", 24800],
  ["Apr", 22100],
  ["May", 27400],
  ["Jun", 31000],
];

const LESSON_6: XlLesson = {
  slug: "data-cleaning",
  n: 6,
  title: "Data Cleaning in Excel: Duplicates, Blanks, TRIM and Validation",
  metaTitle: "Excel Data Cleaning: Duplicates, TRIM & Text to Columns",
  metaDescription:
    "Clean spreadsheet data in Excel: remove duplicates, handle missing values, TRIM stray spaces, Text to Columns, Flash Fill and Data Validation.",
  summary:
    "Real exports arrive dirty. This is the sequence that turns one into a table you can total, look up and pivot without the numbers lying to you.",
  level: "Intermediate",
  minutes: 18,
  objectives: [
    "Find and remove duplicate records safely",
    "Decide how to handle missing values instead of guessing",
    "Strip stray spaces that silently break lookups",
    "Split combined fields and prevent bad input with Data Validation",
  ],
  sections: [
    {
      kind: "prose",
      id: "why",
      heading: "Clean the data before you analyse it",
      body: [
        "Every analysis you have seen so far assumes the data is tidy: one record per row, one field per column, consistent spelling, no stray spaces, no blanks hiding in numeric columns. Real exports are none of those things.",
        "The order matters. Cleaning after you have built formulas means re-doing the formulas. Cleaning first means everything downstream is trustworthy. The sequence below is the one used in practice: standardise text, resolve duplicates, decide on missing values, split combined fields, then lock the result down with validation.",
      ],
    },
    {
      kind: "table",
      id: "the-mess",
      heading: "The dirty export",
      table: {
        caption: "Sheet 'Dirty' — five rows from a typical CRM export, occupying A1:C6. Note the leading and trailing spaces, the mixed capitalisation and the empty email.",
        headers: ["Row", "A · Client Name", "B · Email", "C · Amount"],
        rows: [
          ["2", "··Alpha Pty Ltd·", "·INFO@ALPHA.CO.ZA·", "12500"],
          ["3", "Beta Holdings··", "accounts@beta.co.za", "8200.5"],
          ["4", "Alpha Pty Ltd", "info@alpha.co.za", "12500"],
          ["5", "··Gamma CC", "", "3100"],
          ["6", "alpha pty ltd", "info@alpha.co.za", "12500"],
        ],
      },
      note: "Each dot represents one space character. Five rows, but only three distinct clients — and Excel sees five different names.",
    },
    {
      kind: "functions",
      id: "tools",
      heading: "The cleaning tools",
      functions: [
        {
          name: "Remove Duplicates",
          category: "Data-cleaning tools",
          purpose: "Deletes repeated rows based on the columns you select. Found on the Data tab.",
          syntax: "Data → Data Tools → Remove Duplicates → select the columns that define a duplicate",
          examples: [
            { label: "On the raw names", formula: "Remove Duplicates on column A alone", result: "Only 1 duplicate removed — 5 rows become 4. The space-padded variant survives." },
            { label: "On the normalised names", formula: "Add =TRIM(UPPER(A2)) in D2, then remove duplicates on D", result: "3 unique clients, correctly identified" },
          ],
          keyPoints: [
            "It deletes permanently. Work on a copy, or keep the original sheet untouched.",
            "It keeps the first occurrence and removes later ones. Sort first if 'first' matters.",
            "It is case-insensitive on text, but it is not space-insensitive. TRIM before you deduplicate.",
            "Selecting every column means a row is only a duplicate if all fields match — far stricter than most people expect.",
          ],
          mistakes: [
            "Deduplicating before trimming spaces, so 'Alpha Pty Ltd' and '·Alpha Pty Ltd·' both survive.",
            "Deduplicating on a single column when two genuinely different clients share that value.",
            "Running it on the only copy of the source data.",
          ],
          version: "Excel 2007 and later.",
        },
        {
          name: "Handling missing values",
          category: "Data-cleaning tools",
          purpose: "Deciding what a blank means before any function touches it. There is no universal right answer — only an answer you can justify.",
          syntax: "No single function. The decision is: leave blank, fill with a constant, or impute.",
          examples: [
            { label: "How many are missing", formula: "=COUNTBLANK(B2:B5)", result: "2 on the Payments sheet" },
            { label: "Mean of the values you do have", formula: "=AVERAGE(B2:B5)", result: "21500" },
            { label: "Fill blanks with that mean", formula: "F5 → Special → Blanks, then =AVERAGE(B$2:B$5) and Ctrl+Enter", result: "Beta and Delta both become 21500" },
          ],
          keyPoints: [
            "Leaving a blank means 'not observed'. AVERAGE, SUM, MIN and MAX all skip it.",
            "Filling with 0 means 'observed as zero'. That is a different claim and it changes every average.",
            "Filling with the mean preserves the total but understates the spread. It is a modelling choice, not a fix.",
            "Whatever you choose, record it. A reviewer needs to know which blanks were real and which were filled.",
          ],
          mistakes: [
            "Filling blanks with 0 without saying so, then reporting an average that has quietly fallen.",
            "Deleting rows with any missing field, which can bias the sample towards complete records.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "TRIM",
          category: "Text functions",
          purpose: "Removes leading and trailing spaces and collapses repeated internal spaces to one.",
          syntax: "=TRIM(text)",
          examples: [
            { label: "Clean a client name", formula: '=TRIM("  Alpha Pty Ltd ")', result: '"Alpha Pty Ltd"' },
            { label: "How many stray characters were there", formula: "=LEN(A2)-LEN(TRIM(A2))", result: "3 for row 2" },
            { label: "Full normalisation", formula: "=TRIM(UPPER(SUBSTITUTE(A2;CHAR(160);\" \")))", result: '"ALPHA PTY LTD" — also handles non-breaking spaces' },
          ],
          keyPoints: [
            "TRIM does not remove non-breaking spaces, character 160, which arrive from web pages and PDFs. Replace them with a normal space first.",
            "TRIM does not change capitalisation. Pair it with UPPER, LOWER or PROPER.",
            "TRIM returns a new value in a new cell. To fix the column in place, paste the result back as values.",
          ],
          mistakes: ["Trimming, then pasting over the original with formatting instead of values, so the formulas remain and the column recalculates."],
          version: "All versions of Excel.",
        },
        {
          name: "Text to Columns",
          category: "Data-cleaning tools",
          purpose: "Splits one column into several on a delimiter or a fixed width. Data tab, Data Tools group.",
          syntax: "Data → Text to Columns → Delimited or Fixed width → choose the separator → Finish",
          examples: [
            { label: "Split a comma-separated field", formula: "Delimited → Comma", result: '"Mokoena, Thabo" becomes two columns' },
            { label: "Convert text-numbers to real numbers", formula: "Delimited → Next → Next → Finish", result: "A column SUM refused to total now totals" },
            { label: "Split a date", formula: "Delimited → set the column format to Date → DMY", result: "Ambiguous dates resolve correctly" },
          ],
          keyPoints: [
            "The trick most people miss: running it with no delimiter selected converts a text-number column into real numbers in place.",
            "It writes into the columns to the right. Make sure they are empty or the data will be overwritten.",
            "The TEXTSPLIT function does the same thing with a formula, in Microsoft 365 only.",
          ],
          mistakes: ["Running it with occupied columns to the right and losing that data with no warning."],
          version: "All versions of Excel. TEXTSPLIT requires Microsoft 365.",
        },
        {
          name: "Flash Fill",
          category: "Data-cleaning tools",
          purpose: "Infers a pattern from examples you type and applies it down the column. Home tab, or Ctrl + E.",
          syntax: "Type the result you want in the first one or two rows, then press Ctrl + E",
          examples: [
            { label: "Extract a first name", formula: 'Type "Thabo" beside "Thabo Mokoena", then Ctrl + E', result: "First names filled down the column" },
            { label: "Reformat a phone number", formula: 'Type "011 887 4500" beside "0118874500", then Ctrl + E', result: "All numbers reformatted" },
            { label: "Build an email from two columns", formula: "Type one full address, then Ctrl + E", result: "Addresses generated for every row" },
          ],
          keyPoints: [
            "Flash Fill writes static values, not formulas. It will not update if the source changes.",
            "It works on patterns it can infer. With only one example it can guess wrong — give it two or three.",
            "Always check the last rows. Flash Fill fills quietly and a wrong inference on row 900 looks identical to a right one.",
          ],
          mistakes: ["Trusting Flash Fill on a large column without scrolling to the end and spot-checking."],
          version: "Excel 2013 and later.",
        },
        {
          name: "Find & Replace",
          category: "Data-cleaning tools",
          purpose: "Replaces text across a selection. Ctrl + H.",
          syntax: "Ctrl + H → Find what → Replace with → Replace All",
          examples: [
            { label: "Strip a currency prefix", formula: 'Find "R " replace with nothing', result: '"R 12,500.00" becomes "12,500.00"' },
            { label: "Standardise a suffix", formula: 'Find " Ltd" replace with " Limited"', result: "All company suffixes consistent" },
            { label: "Remove line breaks inside cells", formula: "Find Ctrl + J, replace with a space", result: "Multi-line cells become single-line" },
          ],
          keyPoints: [
            "Select the range first. With nothing selected it replaces across the entire worksheet.",
            "Options → Match entire cell avoids replacing part of a longer word.",
            "Ctrl + J in the Find box inserts a line break, which is how you remove them from inside cells.",
          ],
          mistakes: ["Running Replace All across a whole sheet and changing values inside formulas as well as data."],
          version: "All versions of Excel.",
        },
        {
          name: "Data Validation",
          category: "Data-cleaning tools",
          purpose: "Restricts what can be typed into a cell, so the data stays clean after you have cleaned it. Data tab.",
          syntax: "Data → Data Validation → Allow → choose a rule → set the input message and error alert",
          examples: [
            { label: "Amounts must be a positive number", formula: "Allow: Decimal, Data: greater than, Minimum: 0", result: "Negative and text entries are rejected" },
            { label: "Region must come from a list", formula: "Allow: List, Source: =Regions", result: "A dropdown of the four valid regions" },
            { label: "Block accidental duplicates", formula: 'Custom: =COUNTIF($A$2:$A$100;A2)=1', result: "A second identical client name is refused" },
          ],
          keyPoints: [
            "Validation only applies to newly typed data. It does not clean what is already there.",
            "Always write an input message. A rejection with no explanation just produces support tickets.",
            "Paste can bypass validation rules, so it is a guard-rail, not a guarantee.",
          ],
          mistakes: ["Assuming validation retroactively fixed existing bad data. It did not — it only prevents new bad data."],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "steps",
      id: "sequence",
      heading: "The cleaning sequence, in order",
      steps: [
        { title: "Copy the source and work on the copy", detail: "Never clean in place. Keep the raw export as evidence of where you started." },
        { title: "Standardise text", detail: "TRIM spaces, replace non-breaking spaces, then apply consistent capitalisation. Do this before anything that compares values." },
        { title: "Convert text-numbers to numbers", detail: "Text to Columns with no delimiter, or VALUE(). Confirm with COUNT versus COUNTA." },
        { title: "Resolve duplicates", detail: "Decide which columns define a duplicate, sort so the row you want to keep comes first, then remove." },
        { title: "Decide on missing values", detail: "Count them with COUNTBLANK, choose a treatment, and write the decision down." },
        { title: "Split combined fields", detail: "Text to Columns or Flash Fill. One field per column." },
        { title: "Lock it down", detail: "Apply Data Validation to the columns people will type into, then convert the range to a Table so new rows inherit the rules." },
      ],
    },
    {
      kind: "table",
      id: "before-after",
      heading: "Before and after",
      table: {
        caption: "The Dirty sheet after TRIM and UPPER, with duplicates resolved.",
        headers: ["", "Before", "After"],
        rows: [
          ["Rows", "5", "3"],
          ["Distinct names Excel can tell apart", "4 — case is ignored, spaces are not", "3"],
          ["Names with leading or trailing spaces", "4 of 5", "0"],
          ["Blank email addresses", "1", "1 — flagged, not silently dropped"],
          ["Safe to use as a lookup key", "No", "Yes"],
        ],
        resultColumns: [2],
      },
      note: "The blank email in row 5 is still blank. Cleaning does not mean inventing data — it means knowing exactly what is missing and why.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "Standardise text before you deduplicate or look up. Spaces and capitalisation break both.",
        "Remove Duplicates is permanent and keeps the first match. Sort first, and work on a copy.",
        "A blank is not a zero. Choose deliberately and record the choice.",
        "Text to Columns with no delimiter is the fastest text-to-number conversion in Excel.",
        "Flash Fill writes static values. Verify the bottom of the column.",
        "Data Validation prevents new bad data; it does not repair existing data.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Deduplicating on untrimmed text and concluding the data has no duplicates.",
        "Filling blanks with 0 and reporting a lower average without disclosing it.",
        "Running Find & Replace across a whole sheet and altering formulas.",
        "Trusting Flash Fill without checking the last rows.",
        "Cleaning in place with no copy of the original export.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "On the Dirty sheet, add a column D that holds a normalised client key: trimmed, single-spaced and upper case. Count the distinct clients. Then remove the duplicate rows using that key, confirm you are left with three rows, and report how much total Amount remains.",
      answer: 'D2: =TRIM(UPPER(SUBSTITUTE(A2;CHAR(160);" "))) → "ALPHA PTY LTD". The distinct count is 3. After removing duplicates on column D you keep rows for ALPHA PTY LTD (12500), BETA HOLDINGS (8200.5) and GAMMA CC (3100), a total of 23800.50. Before cleaning, summing column C gave 48800.50 — the duplicates were inflating the total by 25000, more than half.',
    },
  ],
  quiz: [
    {
      q: "You run Remove Duplicates on a name column and it reports no duplicates, but you can see the same client twice. What is the most likely cause?",
      options: [
        "Remove Duplicates is case-sensitive",
        "One of the names has a leading or trailing space",
        "The column is formatted as text",
        "You selected too many columns",
      ],
      answer: 1,
      explain: "Remove Duplicates is case-insensitive but not space-insensitive. TRIM the column first, then deduplicate.",
    },
    {
      q: "On the Payments sheet B2=25000, B3 blank, B4=18000, B5 blank. You fill both blanks with the mean of the existing values. What is the mean you used?",
      options: ["10750", "14333.33", "21500", "43000"],
      answer: 2,
      explain: "AVERAGE(B2:B5) ignores the blanks: (25000 + 18000) ÷ 2 = 21500. Note this is not 20833 — that figure would only be correct if a third value of 19500 were also present.",
    },
    {
      q: "What does TRIM not remove?",
      options: ["Leading spaces", "Trailing spaces", "Repeated internal spaces", "Non-breaking spaces (character 160)"],
      answer: 3,
      explain: "TRIM handles ordinary spaces only. Non-breaking spaces arrive from web pages and PDFs and must be replaced with SUBSTITUTE(text;CHAR(160);\" \") first.",
    },
    {
      q: "A numeric column will not total. What is the fastest in-place fix?",
      options: [
        "Re-type every value",
        "Text to Columns with no delimiter selected",
        "Apply currency formatting",
        "Use SUMPRODUCT instead of SUM",
      ],
      answer: 1,
      explain: "Running Text to Columns and pressing Finish with no delimiter converts the column from text to real numbers in place. Currency formatting changes only the display.",
    },
    {
      q: "What is the main limitation of Flash Fill?",
      options: [
        "It only works on numbers",
        "It writes static values that never update",
        "It requires Microsoft 365",
        "It cannot fill more than 100 rows",
      ],
      answer: 1,
      explain: "Flash Fill infers a pattern and writes plain values. If the source changes, the filled column does not follow — and a wrong inference is not flagged.",
    },
  ],
  workbookSheets: [
    { name: "Dirty", headers: DIRTY_HEADERS, rows: DIRTY_ROWS },
    { name: "Payments", headers: PAYMENTS_HEADERS, rows: PAYMENTS_ROWS },
  ],
  tool: {
    label: "Clean my file now",
    path: "/",
    note: "Run the same sequence — trim, standardise, deduplicate, flag blanks — across a whole file, and review every change before you export.",
  },
  faq: [
    {
      q: "Should I fill missing values with the mean, with zero, or leave them blank?",
      a: "It depends on what the blank means. If the value was never recorded, leave it blank and let the functions skip it. If zero is a true observation, type 0. Mean-filling is a modelling technique for statistical work and should be disclosed in the report. What you must not do is fill silently.",
    },
    {
      q: "How do I remove duplicates without deleting the rows?",
      a: "Add a helper column with =COUNTIF(A$2:A2;A2). A result of 1 marks the first occurrence, 2 and above mark repeats. Filter on that column to review or hide duplicates while keeping the data intact.",
    },
    {
      q: "Is Data Validation enough to keep a shared workbook clean?",
      a: "No. Paste operations can bypass it and it never touches data already in the sheet. It reduces new errors; it does not guarantee a clean file. Pair it with a periodic validation pass over the whole column.",
    },
  ],
};

const LESSON_7: XlLesson = {
  slug: "pivot-tables",
  n: 7,
  title: "PivotTables: Summarise Thousands of Rows in Five Steps",
  metaTitle: "Excel PivotTable Tutorial: Rows, Columns, Values & Filters",
  metaDescription:
    "Build an Excel PivotTable in five steps. Rows, columns, values and filters explained with a worked example and the mistakes that break totals.",
  summary:
    "The fastest way to summarise a large dataset without writing a single formula. Four areas, five steps, and one rule about the source data.",
  level: "Intermediate",
  minutes: 15,
  objectives: [
    "Explain what the four PivotTable areas do",
    "Build a summary from the Sales dataset in five steps",
    "Change how values are aggregated",
    "Avoid the source-data problems that produce wrong totals",
  ],
  sections: [
    {
      kind: "prose",
      id: "why",
      heading: "Why PivotTables instead of SUMIF",
      body: [
        "SUMIF answers a question about one category you already named. A PivotTable answers the question for every category at once, and updates when the data changes.",
        "A PivotTable is a cached summary, not a live formula. It reads a snapshot of the source and holds it in memory. That is why it is fast on large datasets and why you must refresh it after changing the source.",
      ],
    },
    {
      kind: "table",
      id: "components",
      heading: "The four areas",
      table: {
        headers: ["Area", "What goes there", "What it does", "In our example"],
        rows: [
          ["Filters", "A field you want to slice the whole report by", "Adds a dropdown above the table that filters everything", "Rep"],
          ["Columns", "A field whose values become column headings", "Splits each row across columns", "Product"],
          ["Rows", "A field whose values become row labels", "Lists the categories down the side", "Region"],
          ["Values", "The numeric field you are measuring", "Aggregated — summed, counted or averaged", "Sum of Sales"],
        ],
      },
      note: "Drag the same field into Rows twice and you get a hierarchy: Region on the outside, Rep nested inside it. That is how multi-level summaries are built.",
    },
    {
      kind: "steps",
      id: "five-steps",
      heading: "Five steps to create a PivotTable",
      steps: [
        { title: "Click any cell inside the data", detail: "Do not select the whole range. Excel detects the surrounding block automatically, and selecting manually is how stray blank columns get included." },
        { title: "Insert → PivotTable", detail: "Choose New Worksheet so the summary sits apart from the source. Keyboard: Alt, N, V, T." },
        { title: "Check the source range in the dialog", detail: "Confirm it covers every row and excludes any totals row already in the data. This is the step most often skipped and the step that causes wrong totals." },
        { title: "Drag fields into the four areas", detail: "Region to Rows, Product to Columns, Sales to Values. The table appears immediately." },
        { title: "Set the value aggregation and format", detail: "Click the field in Values → Value Field Settings → Sum, Count or Average. Then apply number formatting to the values, not to the cells, so it survives a refresh." },
      ],
    },
    {
      kind: "table",
      id: "result",
      heading: "The result",
      table: {
        caption: "Region in Rows, Product in Columns, Sum of Sales in Values. Row labels are sorted alphabetically by default.",
        headers: ["Row Labels", "Gadget", "Widget", "Grand Total"],
        rows: [
          ["East", "(blank)", "4250", "4250"],
          ["North", "5600", "2500", "8100"],
          ["South", "5200", "5000", "10200"],
          ["West", "3200", "3000", "6200"],
          ["Grand Total", "14000", "14750", "28750"],
        ],
        resultColumns: [3],
      },
      note: "The grand total of 28750 matches =SUM(F2:F11) exactly. That check — PivotTable grand total against a plain SUM — is the first thing to do with any new PivotTable. East sold no Gadgets, so that cell is empty by default rather than zero. Set PivotTable Options → 'For empty cells show' to 0 if you want an explicit zero; a blank and a zero read very differently to a finance reviewer.",
    },
    {
      kind: "prose",
      id: "aggregation",
      heading: "Changing how values are aggregated",
      body: [
        "The Values area defaults to Sum for numbers and Count for text. If a numeric column has been imported as text, you will get Count and wonder why your total is 10 instead of 28750 — which is also a useful diagnostic.",
        "Right-click any value in the table, or use Value Field Settings, to switch between Sum, Count, Average, Max, Min, Product and more. The same field can appear in Values twice with different settings, giving you total and average side by side.",
        "Show Values As changes the presentation rather than the calculation: percentage of grand total, percentage of column total, difference from, running total. For a regional share of revenue, use Percentage of Column Total.",
      ],
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "Rows list categories, Columns split them, Values measure, Filters slice the whole report.",
        "Click inside the data rather than selecting it, then verify the source range in the dialog.",
        "A PivotTable is a cache. Refresh it after changing the source — Ctrl + Alt + F5 refreshes everything.",
        "Format the value field, not the cells, so formatting survives a refresh.",
        "Always reconcile the grand total against a plain SUM of the source column.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "A totals row inside the source range, which gets included as a data row and inflates the grand total.",
        "Merged cells in the source, which produce blank row labels and unpredictable grouping.",
        "Blank column headings, which Excel names Field1 and refuse to drag cleanly.",
        "Editing the source and not refreshing, then reporting a stale number.",
        "A fixed source range such as A1:F11 that silently excludes rows added later. Convert the source to a Table so the range grows automatically.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "Build a PivotTable from the Sales sheet with Region in Rows and Sum of Sales in Values. Then add Product to Columns, add Rep to Filters, and change the Sales value field to show Average as well as Sum. Confirm the grand total still reads 28750.",
      answer: "You should reproduce the table above: East 4250, North 8100, South 10200, West 6200, Grand Total 28750. Filtering by Rep = Lerato should leave South 10200 and a grand total of 10200, because all three South rows belong to Lerato. If your grand total differs from =SUM(F2:F11), check the source range in the PivotTable Analyze tab first.",
    },
  ],
  quiz: [
    {
      q: "What is the grand total of Sum of Sales in the worked PivotTable?",
      options: ["14000", "14750", "28750", "10200"],
      answer: 2,
      explain: "The grand total is 28750, which matches =SUM(F2:F11). Gadget contributes 14000 and Widget 14750.",
    },
    {
      q: "Your PivotTable shows Count of Sales instead of Sum of Sales. What does that tell you?",
      options: [
        "You chose the wrong aggregation and nothing else",
        "The Sales column contains text, not numbers",
        "The PivotTable needs refreshing",
        "There are duplicate rows",
      ],
      answer: 1,
      explain: "Excel defaults to Sum for numeric fields and Count for text. Getting Count on a numeric column means the values were imported as text — fix the source with Text to Columns, then refresh.",
    },
    {
      q: "You add ten rows to the source data but the PivotTable total does not change. Why?",
      options: [
        "PivotTables never update",
        "A PivotTable is a cache and must be refreshed",
        "You need to rebuild the PivotTable",
        "The new rows must be formatted first",
      ],
      answer: 1,
      explain: "A PivotTable works from a cached snapshot. Refresh with Data → Refresh All or Ctrl + Alt + F5. If the source range is fixed, also check that it now covers the new rows.",
    },
    {
      q: "Which PivotTable area would you use to add a dropdown that filters the entire report by sales representative?",
      options: ["Rows", "Columns", "Values", "Filters"],
      answer: 3,
      explain: "The Filters area places a report-level dropdown above the PivotTable. Slicers do the same job with a more usable on-screen control.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
  ],
  tool: {
    label: "Explore and profile my dataset",
    path: "/data-toolbox",
    note: "Profile a large file and see the shape of every column before you decide what to pivot.",
  },
  faq: [
    {
      q: "How large a dataset can a PivotTable handle?",
      a: "PivotTables handle hundreds of thousands of rows comfortably in a normal workbook. Beyond that, add the data to the Data Model when creating the PivotTable, which uses a columnar store and copes with millions of rows. Browser-based tools hit memory limits far sooner than desktop Excel does.",
    },
    {
      q: "Should I convert my source data to a Table first?",
      a: "Yes. A Table gives the range a name, expands automatically when rows are added, and makes the source range self-maintaining. It removes the most common PivotTable failure — a fixed range that quietly excludes new data.",
    },
    {
      q: "Why does my PivotTable show blank row labels?",
      a: "Usually merged cells in the source, or a column with no heading. Unmerge the source, give every column a heading, then refresh. PivotTables assume one value per cell.",
    },
  ],
};

const LESSON_8: XlLesson = {
  slug: "charts",
  n: 8,
  title: "Excel Charts: Choosing Between Line, Bar, Histogram and Scatter",
  metaTitle: "Excel Chart Types Explained: Line, Bar, Histogram & Scatter",
  metaDescription:
    "When to use a line, bar, histogram or scatter chart in Excel, with worked data, bin calculations and the mistakes that mislead a reader.",
  summary:
    "Four chart types cover most business reporting. The skill is not building them — it is choosing the one that shows the shape of your data without distorting it.",
  level: "Intermediate",
  minutes: 14,
  objectives: [
    "Match a chart type to the question you are answering",
    "Build each of the four types from the practice data",
    "Set histogram bins deliberately",
    "Avoid the formatting choices that mislead a reader",
  ],
  sections: [
    {
      kind: "diagram",
      id: "chart-choice",
      heading: "Which chart, which question",
      diagram: "chart-choice",
      caption: "Start from the question, not from the chart gallery. Each of these four answers a different one.",
    },
    {
      kind: "functions",
      id: "types",
      heading: "The four chart types",
      functions: [
        {
          name: "Line chart",
          category: "Charts",
          purpose: "Shows how one measure changes over an ordered sequence, usually time.",
          syntax: "Select the data → Insert → Charts → Line",
          examples: [
            { label: "Monthly revenue", formula: "Select A1:B7 on the Monthly sheet → Insert → Line", result: "Six points rising from 18500 in Jan to 31000 in Jun" },
          ],
          keyPoints: [
            "The horizontal axis must be ordered. A line between unordered categories implies a sequence that does not exist.",
            "Gaps in the data should stay visible. Choose 'Show empty cells as gaps' rather than connecting across them.",
            "Two or three series is the useful maximum. Beyond that the lines cross into noise.",
          ],
          mistakes: [
            "Truncating the vertical axis to exaggerate a small change. A line chart on a partial scale makes 2% look like 40%.",
            "Plotting months alphabetically because Excel sorted them, so Apr precedes Jan.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "Bar and column chart",
          category: "Charts",
          purpose: "Compares a measure across discrete categories.",
          syntax: "Select the data → Insert → Charts → Clustered Column or Bar",
          examples: [
            { label: "Sales by region", formula: "PivotTable with Region in Rows, Sum of Sales in Values → Insert → Column", result: "South 10200, North 8100, West 6200, East 4250" },
          ],
          keyPoints: [
            "Use a column chart for few categories with short labels, and a horizontal bar chart when the labels are long or there are many of them.",
            "Sort the bars by value so the comparison is immediate. Alphabetical order hides the ranking.",
            "The vertical axis must start at zero. A bar encodes magnitude by length, and truncating the axis falsifies the ratio.",
          ],
          mistakes: [
            "A truncated axis on a bar chart. Unlike a line chart, there is no honest reason to do this.",
            "Three-dimensional bars, which distort the apparent length through perspective.",
            "A pie chart with more than about five slices, where the angles become impossible to compare.",
          ],
          version: "All versions of Excel.",
        },
        {
          name: "Histogram",
          category: "Charts",
          purpose: "Shows how frequently values fall into ranges — the distribution of one numeric column.",
          syntax: "Select one numeric column → Insert → Insert Statistic Chart → Histogram",
          examples: [
            { label: "Distribution of order sizes", formula: "Select D1:D11 on the Sales sheet → Histogram with bin width 5", result: "2 / 5 / 2 / 1 / 0 across the bins 0-5, 5-10, 10-15, 15-20 and 20-25" },
          ],
          keyPoints: [
            "A histogram plots one variable. A column chart plots categories. They look similar and are not interchangeable.",
            "Excel bins are lower-exclusive and upper-inclusive: a value of exactly 10 falls in the 5 to 10 bin, not the 10 to 15 bin.",
            "Bin width is a judgement. Too narrow and you see noise; too wide and you hide the shape. Try three widths before settling.",
          ],
          mistakes: [
            "Leaving the automatic bins when the default hides the pattern you are trying to show.",
            "Reading a gap between bars as meaningful. In a histogram the bars should touch, because the variable is continuous.",
          ],
          version: "Built-in chart type from Excel 2016. In Excel 2013 and earlier use the Analysis ToolPak or the FREQUENCY function with a column chart.",
        },
        {
          name: "Scatter chart",
          category: "Charts",
          purpose: "Shows the relationship between two numeric variables, one point per record.",
          syntax: "Select two numeric columns → Insert → Charts → Scatter",
          examples: [
            { label: "Units against Sales", formula: "Select D1:D11 and F1:F11 → Scatter", result: "Two straight lines — one for Widget at 250 per unit, one for Gadget at 400" },
          ],
          keyPoints: [
            "Both axes are numeric and both are scaled by value. That is what distinguishes it from a line chart.",
            "Use it to spot clusters, outliers and relationships before you model anything.",
            "A visible pattern is not proof of cause. Two lines here reflect two different unit prices, not two different sales behaviours.",
          ],
          mistakes: [
            "Using a line chart for two numeric variables, which orders the points by row position rather than by value.",
            "Adding a trendline across two distinct clusters and reporting one relationship where there are two.",
          ],
          version: "All versions of Excel.",
        },
      ],
    },
    {
      kind: "table",
      id: "monthly",
      heading: "Worked data: the monthly series",
      table: {
        caption: "Sheet 'Monthly' — the source for the line chart.",
        headers: ["Month", "Revenue", "Change from previous month"],
        rows: [
          ["Jan", "18500", "—"],
          ["Feb", "21200", "+2700"],
          ["Mar", "24800", "+3600"],
          ["Apr", "22100", "-2700"],
          ["May", "27400", "+5300"],
          ["Jun", "31000", "+3600"],
        ],
        resultColumns: [1, 2],
      },
      note: "Total 145000, average 24166.67. June is 12500 above January, a rise of 67.6%. Note the dip in April — a line chart makes that visible in a way a table of six numbers does not.",
    },
    {
      kind: "table",
      id: "bins",
      heading: "Worked data: histogram bins",
      table: {
        caption: "Order sizes from the Sales sheet, sorted: 3, 5, 6, 7, 8, 9, 10, 12, 14, 20. Bin width 5.",
        headers: ["Bin", "Values in bin", "Count", "Bar height"],
        rows: [
          ["0 to 5", "3, 5", "2", "██"],
          ["5 to 10", "6, 7, 8, 9, 10", "5", "█████"],
          ["10 to 15", "12, 14", "2", "██"],
          ["15 to 20", "20", "1", "█"],
          ["20 to 25", "none", "0", ""],
        ],
        resultColumns: [2],
      },
      note: "The ten counts sum to the ten source rows — always check that. Because bins are upper-inclusive, the value 20 lands in the 15 to 20 bin and the 20 to 25 bin is empty. That empty final bin is a signal to set the upper bound at 20 instead.",
    },
    {
      kind: "table",
      id: "when-to-use",
      heading: "When to use each chart type",
      table: {
        headers: ["Question you are answering", "Use", "Do not use"],
        rows: [
          ["How has this changed over time?", "Line", "Pie"],
          ["Which category is largest?", "Bar or column, sorted by value", "Scatter"],
          ["How are these values distributed?", "Histogram", "Line"],
          ["Do these two variables move together?", "Scatter", "Bar"],
          ["How is a total split into parts?", "Pie with five slices or fewer, or a stacked bar", "Line or scatter"],
          ["How does each part contribute over time?", "Stacked column or area", "Pie"],
        ],
      },
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "Line for change over time, bar for comparison between categories, histogram for distribution, scatter for relationships.",
        "A bar chart's axis starts at zero. A line chart's axis can be truncated only if you label it clearly.",
        "Sort categorical bars by value, not alphabetically.",
        "Histogram bins are upper-inclusive, and bin width is a choice you should defend.",
        "A scatter shows association, not cause.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "A truncated axis on a bar chart, which distorts every comparison in it.",
        "Months plotted in alphabetical order.",
        "A pie chart with a dozen slices.",
        "A 3D chart, where perspective changes apparent size.",
        "A chart with no axis titles, so the reader has to guess the units.",
        "Connecting across missing data instead of showing the gap.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "From the Monthly sheet build a line chart of revenue and add a second series showing a flat target line at 25000. Then build a histogram of the Units column from the Sales sheet, first with the default bins and then with a bin width of 5. Describe in one sentence what each version shows.",
      answer: "The line chart shows revenue staying below the 25000 target for four months — March peaks at 24800, just short — dipping in April, then clearing the target in May at 27400 and June at 31000. With default bins the histogram is uninformative; with a bin width of 5 it shows a clear concentration between 5 and 10 units, a thin tail to 14, and one outlier at 20. The single sentence: default bins hid a right-skewed distribution with one outlier.",
    },
  ],
  quiz: [
    {
      q: "You want to show how monthly revenue has changed over six months. Which chart?",
      options: ["Pie", "Line", "Histogram", "Scatter"],
      answer: 1,
      explain: "A line chart is the right choice for one measure across an ordered sequence. A histogram would show distribution and discard the order entirely.",
    },
    {
      q: "With a bin width of 5, which bin does an order of exactly 10 units fall into?",
      options: ["0 to 5", "5 to 10", "10 to 15", "It is excluded"],
      answer: 1,
      explain: "Excel histogram bins are lower-exclusive and upper-inclusive, so 10 falls in the 5 to 10 bin. The 10 to 15 bin starts above 10.",
    },
    {
      q: "Why must a bar chart's value axis start at zero?",
      options: [
        "Excel requires it",
        "Bars encode magnitude by length, so a truncated axis falsifies the ratios",
        "It makes the chart smaller",
        "It is only a convention",
      ],
      answer: 1,
      explain: "A reader compares bar lengths. Truncating the axis makes small differences look large. Line charts show position rather than length, so truncation is defensible there if labelled.",
    },
    {
      q: "A scatter of Units against Sales shows two distinct straight lines. What is the correct reading?",
      options: [
        "Sales cause units to increase",
        "The data has an error",
        "Two different unit prices are producing two relationships, not one",
        "A trendline should be added across both",
      ],
      answer: 2,
      explain: "Sales equals Units multiplied by Unit Price, and there are two prices in the data, so there are two lines. Fitting one trendline across both clusters would describe a relationship that does not exist.",
    },
    {
      q: "The built-in Histogram chart type is available from which Excel version?",
      options: ["Excel 2010", "Excel 2013", "Excel 2016", "Excel 2019"],
      answer: 2,
      explain: "Histogram became a built-in chart type in Excel 2016. In Excel 2013 and earlier you need the Analysis ToolPak or FREQUENCY with a column chart.",
    },
  ],
  workbookSheets: [
    { name: "Monthly", headers: MONTHLY_HEADERS, rows: MONTHLY_ROWS },
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
  ],
  tool: {
    label: "Chart my CSV",
    path: "/csv-statistics-and-charts",
    note: "Profile a file and generate the summary statistics behind these charts without opening Excel at all.",
  },
  faq: [
    {
      q: "When is a pie chart actually the right choice?",
      a: "When you have one total, five parts or fewer, and the reader needs the share of each part rather than a ranking. For anything else a sorted bar chart is easier to read, because people compare lengths more accurately than angles.",
    },
    {
      q: "How do I choose a histogram bin width?",
      a: "There is no single correct answer. Start with Excel's default, then try a narrower and a wider setting. Choose the width at which the shape you care about is visible without the chart becoming noise, and say what you chose in the caption.",
    },
    {
      q: "My line chart plots the months in the wrong order. Why?",
      a: "Excel is treating the month names as text and sorting them alphabetically, so Apr comes before Jan. Use real dates formatted as month names, or add a numeric month column and plot against that.",
    },
  ],
};

const LESSON_9: XlLesson = {
  slug: "power-query",
  n: 9,
  title: "Power Query: Extract, Transform, Load, and Merge versus Append",
  metaTitle: "Power Query in Excel: ETL, Merge vs Append Explained",
  metaDescription:
    "Power Query's Extract, Transform, Load workflow, and the difference between Merge and Append, with row and column counts you can check.",
  summary:
    "Record your cleaning once and replay it every month. Power Query's three stages, and the one distinction — Merge adds columns, Append adds rows — that trips up most new users.",
  level: "Advanced",
  minutes: 16,
  objectives: [
    "Describe the Extract, Transform, Load stages",
    "Explain why Power Query steps are repeatable",
    "Choose correctly between Merge and Append",
    "Predict the row and column count of each operation",
  ],
  sections: [
    {
      kind: "prose",
      id: "what",
      heading: "What Power Query is",
      body: [
        "Power Query is Excel's built-in tool for getting data in from somewhere else, reshaping it, and loading the result into a worksheet. On the Data tab it appears as the Get & Transform group.",
        "The property that makes it worth learning is that every step is recorded. You do not re-do the cleaning next month — you point the query at the new file and refresh, and the same steps run again in the same order. Your original source file is never modified.",
        "The steps are written in a language called M, visible in the formula bar of the Power Query Editor. You rarely need to write M by hand; the interface generates it.",
      ],
    },
    {
      kind: "diagram",
      id: "etl",
      heading: "The Extract, Transform, Load workflow",
      diagram: "etl",
      caption: "Three stages, one direction. Nothing reaches your worksheet until you close and load, which is why you can experiment freely in the editor.",
    },
    {
      kind: "steps",
      id: "stages",
      heading: "The three stages in detail",
      steps: [
        { title: "Extract — connect to the source", detail: "Data → Get Data → From File, From Database, From Web or From Table/Range. Power Query reads the source and shows a preview. Nothing is imported yet." },
        { title: "Transform — reshape in the editor", detail: "The Power Query Editor opens in a separate window. Promote headers, change types, remove columns, split text, filter rows, unpivot. Every action becomes a named, reorderable step on the right." },
        { title: "Load — put the result somewhere", detail: "Close & Load sends the result to a worksheet table. Close & Load To gives you the choice of a table, a PivotTable, or the Data Model only, which keeps large data out of the visible sheets." },
      ],
    },
    {
      kind: "prose",
      id: "steps-matter",
      heading: "Why the applied steps list matters",
      body: [
        "The Applied Steps pane on the right is the whole point. Each step has a name, an order and an M expression. You can delete a step, reorder them, or insert one in the middle.",
        "Order is not cosmetic. Changing a column type before splitting it behaves differently from splitting first. If a refresh fails, the error names the step that broke, which is far easier to diagnose than a chain of manual edits.",
        "Because the steps are recorded, the query is documentation. A colleague can read the list and understand exactly what was done to the data — something no amount of manual copy-paste achieves.",
      ],
    },
    {
      kind: "diagram",
      id: "merge-append",
      heading: "Merge versus Append",
      diagram: "merge-append",
      caption: "Append stacks tables that have the same shape. Merge joins tables that share a key. Getting this backwards is the most common Power Query mistake.",
    },
    {
      kind: "table",
      id: "comparison",
      heading: "Merge versus Append, side by side",
      table: {
        caption: "Using the Sales table (10 rows, 6 columns) and the Prices table (3 rows, 3 columns).",
        headers: ["", "Append", "Merge"],
        rows: [
          ["What it does", "Stacks one table underneath another", "Joins tables side by side on a key column"],
          ["Analogy", "Piling more pages onto the same report", "Looking up extra detail for each existing row"],
          ["Requirement", "Matching column names and compatible types", "A key column present in both tables"],
          ["Row count", "Adds up — 10 + 10 = 20", "Stays with the left table — 10 rows"],
          ["Column count", "Stays the same — 6 columns", "Adds the columns you expand — 6 becomes 7 or more"],
          ["Typical use", "Twelve monthly files into one year", "Adding a unit price or client name to a transaction list"],
          ["The Excel equivalent", "Copy and paste rows below each other", "VLOOKUP or XLOOKUP, but recorded and repeatable"],
          ["What goes wrong", "Column names differ, so you get null columns", "No match on the key, so the expanded columns are empty"],
        ],
      },
      note: "Append is vertical, Merge is horizontal. If you find yourself appending and getting a table twice as wide with null columns, you meant Merge. If you merge and get three rows where you expected ten, you meant Append.",
    },
    {
      kind: "steps",
      id: "append-how",
      heading: "Append: combining monthly files",
      steps: [
        { title: "Load each file as its own query", detail: "Or use Data → Get Data → From File → From Folder, which picks up every file in the folder and adds new ones automatically." },
        { title: "Check that the column names match", detail: "Append matches on name, not on position. A column called 'Amount' in one file and 'Amt' in another produces two columns, one of them mostly null." },
        { title: "Home → Append Queries", detail: "Choose two tables, or 'three or more' for a list." },
        { title: "Verify the row count", detail: "Ten rows from January plus ten from February must give twenty. If it does not, the column names did not line up." },
      ],
    },
    {
      kind: "steps",
      id: "merge-how",
      heading: "Merge: adding a price list to a transaction list",
      steps: [
        { title: "Load both tables", detail: "The Sales query and the Prices query." },
        { title: "Home → Merge Queries", detail: "Select Sales as the first table and Prices as the second." },
        { title: "Click the key column in each table", detail: "Product in both. The join kind defaults to Left Outer, which keeps every Sales row whether or not it matches." },
        { title: "Expand the new column", detail: "Click the expand icon on the merged column header and tick the fields you want — Code and Unit Price. Untick 'Use original column name as prefix' for cleaner headings." },
        { title: "Check for unmatched rows", detail: "Filter the expanded column for nulls. Every null is a Sales row with no matching product — a data problem worth knowing about before you report the numbers." },
      ],
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "Extract connects, Transform reshapes, Load publishes. Nothing touches your sheet until you load.",
        "Every transformation is a recorded, named, reorderable step.",
        "Append adds rows and needs matching column names. Merge adds columns and needs a shared key.",
        "Append keeps the column count; Merge keeps the row count of the left table.",
        "Always verify the row count after an Append and check for nulls after a Merge.",
      ],
    },
    {
      kind: "callout",
      id: "mistakes",
      heading: "Common mistakes",
      points: [
        "Appending tables whose column names differ slightly, and getting wide tables full of nulls.",
        "Merging on a key with trailing spaces, so nothing matches and every expanded column is empty.",
        "Assuming a Merge filters rows. A Left Outer join keeps all left rows; only an Inner join drops unmatched ones.",
        "Leaving the 'Table.' prefix on expanded column names.",
        "Not refreshing before reporting. The loaded table is a snapshot from the last refresh.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "Load the Sales and Prices sheets as two queries. First, Merge Sales to Prices on Product and expand the Code column — then count the rows and columns and check for nulls. Second, duplicate the Sales query, rename it Sales2, and Append Sales to Sales2 — then count the rows again.",
      answer: "After the Merge you have 10 rows and 7 columns, with no nulls because every product in Sales exists in Prices. After the Append you have 20 rows and 6 columns. If the Merge produced nulls in the Code column, the Product values differ by spacing or capitalisation between the two sheets — trim both before merging.",
    },
  ],
  quiz: [
    {
      q: "You have twelve monthly CSV files with the same columns. Which Power Query operation combines them?",
      options: ["Merge", "Append", "Unpivot", "Group By"],
      answer: 1,
      explain: "Append stacks tables vertically. Twelve files of ten rows each append to 120 rows with the same columns.",
    },
    {
      q: "You merge a 10-row Sales table to a 3-row Prices table on Product, using a Left Outer join. How many rows does the result have?",
      options: ["3", "10", "13", "30"],
      answer: 1,
      explain: "A Left Outer join keeps every row of the left table, so the result has the 10 Sales rows with the Prices columns added alongside. Merge never multiplies rows unless the key is duplicated on the right.",
    },
    {
      q: "After an Append you get 20 rows but 11 columns instead of 6. What happened?",
      options: [
        "Append always adds columns",
        "Column names differed between the two tables, so unmatched columns were added as separate fields",
        "You used Merge by mistake",
        "The files had different row counts",
      ],
      answer: 1,
      explain: "Append matches columns by name. If one file calls a column 'Amount' and the other 'Amt', both survive and each is half empty. Align the names before appending.",
    },
    {
      q: "What is the main advantage of Power Query over cleaning the data manually each month?",
      options: [
        "It is faster to type",
        "Every step is recorded and replays on refresh",
        "It does not need Excel",
        "It permanently edits the source file",
      ],
      answer: 1,
      explain: "The Applied Steps list is a recorded, repeatable recipe. Refreshing re-runs it against new source data. It also never modifies the original file.",
    },
    {
      q: "Where does Power Query live in Excel 2016 and later?",
      options: [
        "A separate application",
        "The Data tab, in the Get & Transform group",
        "The Review tab",
        "It requires a paid add-in",
      ],
      answer: 1,
      explain: "From Excel 2016 onwards Power Query is built in as Get & Transform on the Data tab. Excel 2010 and 2013 need the free Microsoft add-in.",
    },
  ],
  workbookSheets: [
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
    { name: "Prices", headers: PRICES_HEADERS, rows: PRICES_ROWS },
  ],
  tool: {
    label: "Merge and combine my files",
    path: "/merge-spreadsheets",
    note: "Combine several exports into one file in the browser when you do not have desktop Excel to hand.",
  },
  faq: [
    {
      q: "Is Power Query available on every version of Excel?",
      a: "It is built into Excel 2016 and later on Windows, under Get & Transform on the Data tab. Excel 2010 (Professional Plus) and 2013 need the free Microsoft add-in. Current Microsoft 365 builds for Mac include it; older Mac releases did not. Check File → Account if you are unsure.",
    },
    {
      q: "Does Power Query change my source file?",
      a: "No. Power Query reads the source and writes its result to a new location — a worksheet table, a PivotTable or the Data Model. The original file is untouched, which is one of the reasons it is safer than manual editing for repeatable reporting.",
    },
    {
      q: "When should I use Power Query instead of formulas?",
      a: "When the work is repeated, when the source is large, or when several files must be combined. A formula is better for a one-off calculation on a small range. Many analysts use Power Query to produce a clean table and formulas to analyse it.",
    },
  ],
};

const LESSON_10: XlLesson = {
  slug: "interview-questions",
  n: 10,
  title: "Excel Interview Questions for Data Analysis Roles",
  metaTitle: "Excel Interview Questions & Answers for Analyst Roles",
  metaDescription:
    "Six common Excel interview questions with model answers: VLOOKUP vs XLOOKUP, SUMIF vs SUMIFS, PivotTables, missing values, Merge vs Append.",
  summary:
    "Six questions that come up repeatedly in analyst, finance and operations interviews, with model answers that show judgement rather than recall.",
  level: "Intermediate",
  minutes: 14,
  objectives: [
    "Answer the six most common Excel interview questions with reasoning",
    "Show judgement about trade-offs, not just recall of syntax",
    "Recognise the follow-up question behind each one",
  ],
  sections: [
    {
      kind: "prose",
      id: "how",
      heading: "How to answer these well",
      body: [
        "Interviewers asking Excel questions are usually testing three things: whether you know the tool, whether you understand its limits, and whether you would catch your own mistakes. A good answer names the trade-off and the failure mode, not just the syntax.",
        "The answers below follow one shape: the direct answer, the reason, and the caveat. That shape works for almost every technical interview question.",
      ],
    },
    {
      kind: "table",
      id: "q1",
      heading: "Q1. What is the difference between VLOOKUP and XLOOKUP?",
      table: {
        headers: ["", "VLOOKUP", "XLOOKUP"],
        rows: [
          ["Direction", "Right only — the lookup value must be in the leftmost column of the range", "Any direction — the lookup and return arrays are independent"],
          ["Default match", "Approximate. You must pass FALSE for an exact match", "Exact"],
          ["Not-found handling", "Wrap in IFERROR or IFNA", "Built-in fourth argument"],
          ["Inserting a column", "Breaks silently, because col_index_num is a count", "Unaffected, because ranges are named explicitly"],
          ["Availability", "Every version of Excel", "Microsoft 365, Excel 2021 and Excel 2024"],
        ],
      },
      note: "Model answer: 'XLOOKUP removes the three things that make VLOOKUP fragile — the leftmost-column restriction, the approximate-match default and the positional column index. The one reason to keep using VLOOKUP is compatibility: XLOOKUP shows #NAME? in Excel 2019 and earlier, so for a workbook shared widely I would use INDEX and MATCH, which works everywhere and has none of VLOOKUP's weaknesses.'",
    },
    {
      kind: "table",
      id: "q2",
      heading: "Q2. What is the difference between SUMIF and SUMIFS?",
      table: {
        headers: ["", "SUMIF", "SUMIFS"],
        rows: [
          ["Conditions", "One", "Many, combined with AND"],
          ["Argument order", "range, criteria, then the optional sum_range last", "sum_range first, then range-and-criteria pairs"],
          ["Sum range", "Optional — defaults to the criteria range", "Required"],
          ["Maximum conditions", "One", "127 pairs"],
          ["OR logic", "Add two SUMIF results", "Not supported — sum two SUMIFS results or use SUMPRODUCT"],
        ],
      },
      note: "Model answer: 'SUMIF handles a single condition and puts its sum range last and optional. SUMIFS handles many conditions, requires the sum range, and puts it first. The argument order changes between them, which is the detail people get wrong. Neither supports OR — for that you add two results together.' The follow-up they usually ask is whether the ranges must be the same size in SUMIFS: yes, or you get #VALUE!.",
    },
    {
      kind: "table",
      id: "q3",
      heading: "Q3. What is a PivotTable and when would you use one?",
      table: {
        headers: ["Point", "Answer"],
        rows: [
          ["What it is", "An interactive summary that groups a large dataset and aggregates a numeric field, without formulas"],
          ["The four areas", "Rows for categories, Columns to split them, Values to measure, Filters to slice the whole report"],
          ["When to use it", "When you need every category summarised at once, rather than one named category as SUMIF would give"],
          ["The main caveat", "It is a cache, not a live formula. It must be refreshed after the source changes"],
          ["The source-data rules", "One header row, no merged cells, no totals row inside the range, no blank column headings"],
          ["Scale", "Hundreds of thousands of rows comfortably; add to the Data Model for millions"],
        ],
      },
      note: "Model answer: 'A PivotTable summarises a dataset by grouping and aggregating, and it beats SUMIF when you want every category rather than one. The two things I always check are the source range, because a totals row inside it inflates the grand total, and the refresh, because the PivotTable holds a snapshot. I reconcile the grand total against a plain SUM of the source column before I use it.'",
    },
    {
      kind: "table",
      id: "q4",
      heading: "Q4. How do you handle missing values in a dataset?",
      table: {
        caption: "The Payments sheet: Alpha 25000, Beta blank, Gamma 18000, Delta blank.",
        headers: ["Approach", "What it means", "Effect on this data", "When it is right"],
        rows: [
          ["Leave blank", "Not observed", "AVERAGE(B2:B5) = 21500 over the two known values", "Default, and honest about what you have"],
          ["Fill with zero", "Observed as zero", "Sum stays 43000 but COUNT becomes 3, so the average falls to 14333.33", "Only when zero is a true value, such as no transactions that month"],
          ["Fill with the mean", "Imputed", "Both blanks become 21500; the total rises to 86000 and the average stays 21500", "Statistical modelling, disclosed"],
          ["Fill with the median", "Imputed, outlier-resistant", "The median of 25000 and 18000 is also 21500 here", "Skewed data where the mean is pulled by outliers"],
          ["Delete the rows", "Excluded", "Only Alpha and Gamma remain; the total falls to 43000", "Rarely — it biases the sample towards complete records"],
        ],
        resultColumns: [2],
      },
      note: "Model answer: 'It depends on what the blank means, and I would ask before choosing. If it means not observed, I leave it blank, because AVERAGE, SUM and MIN skip blanks while treating them as zero would change every result. On this data the two known values are 25000 and 18000, so their mean is 21500 — not 20833, which is what you would get if a third value of 19500 were also present. Whatever I choose, I record it, because a reviewer needs to know which blanks were real.'",
    },
    {
      kind: "table",
      id: "q5",
      heading: "Q5. What is the difference between Merge and Append in Power Query?",
      table: {
        headers: ["", "Append", "Merge"],
        rows: [
          ["Direction", "Vertical — stacks rows", "Horizontal — joins columns"],
          ["Needs", "Matching column names", "A shared key column"],
          ["Rows", "Adds up", "Follows the left table"],
          ["Columns", "Unchanged", "Increases with the fields you expand"],
          ["Excel equivalent", "Pasting rows underneath", "VLOOKUP or XLOOKUP"],
          ["Failure mode", "Null columns from mismatched names", "Empty expanded columns from unmatched keys"],
        ],
      },
      note: "Model answer: 'Append stacks tables that have the same shape; Merge joins tables that share a key. Appending ten rows to ten rows gives twenty rows and the same columns. Merging a ten-row table to a three-row lookup on a Left Outer join gives ten rows with extra columns. I always verify the row count after an Append and filter for nulls after a Merge, because both fail quietly.'",
    },
    {
      kind: "table",
      id: "q6",
      heading: "Q6. Other questions worth preparing for",
      table: {
        headers: ["Question", "What they are really testing", "The strong answer"],
        rows: [
          ["How do you find duplicates in a large list?", "Whether you know more than Remove Duplicates", "=COUNTIF(A$2:A2;A2) in a helper column marks first occurrences with 1, so you can review before deleting anything."],
          ["What does #N/A mean in a lookup?", "Whether you debug rather than guess", "The lookup value was not found. Check trailing spaces with TRIM, check for a number stored as text, and confirm the exact-match argument is set."],
          ["How do you handle a file too large for Excel?", "Whether you know the limits", "Excel's limit is 1048576 rows per sheet. Power Query with the Data Model handles far more, or process it outside Excel entirely."],
          ["Absolute versus relative references?", "Whether your formulas survive being copied", "$ locks a reference. No $ means it moves when copied. F4 cycles the four forms."],
          ["How would you automate a monthly report?", "Whether you think in repeatable steps", "Power Query for the extraction and reshaping, a Table as the source, PivotTables for the summary, so the whole thing is a refresh rather than a rebuild."],
          ["How do you check your own work?", "Whether you can be trusted with the numbers", "Reconcile every total against an independent calculation — a PivotTable grand total against a plain SUM, a report total against the source file total."],
        ],
      },
      note: "The last one is the question that actually decides the interview. Every employer asking Excel questions wants to know whether you would catch your own error before they did.",
    },
    {
      kind: "callout",
      id: "key-points",
      heading: "Key points",
      points: [
        "Answer with the trade-off and the failure mode, not just the syntax.",
        "Name the compatibility limits — XLOOKUP in Excel 2019, IFS in Excel 2016, histograms before Excel 2016.",
        "For missing values, say it depends on what the blank means, and justify the choice.",
        "Always mention how you would verify the answer. That is the part they are scoring.",
      ],
    },
    {
      kind: "practice",
      id: "practice",
      heading: "Practice exercise",
      task: "Answer all six questions out loud in under ninety seconds each, without looking at the notes. Then take the Payments sheet and work through the missing-values answer with the real numbers, so you can quote 21500 and explain why it is not 20833.",
      answer: "If you can explain that 20833 is the mean of three values — 25000, 18000 and a third value of 19500 — while 21500 is the mean of the two values actually present, you have demonstrated exactly the attention to detail the question is testing.",
    },
  ],
  quiz: [
    {
      q: "On the Payments sheet the known values are 25000 and 18000. What is their mean?",
      options: ["20833", "21500", "43000", "14333.33"],
      answer: 1,
      explain: "(25000 + 18000) ÷ 2 = 21500. The figure 20833 is the mean of three values including 19500, and is a common carry-over error from a longer example.",
    },
    {
      q: "What is the single most important argument-order difference between SUMIF and SUMIFS?",
      options: [
        "SUMIFS takes no sum range",
        "SUMIF puts the sum range last and optional; SUMIFS puts it first and required",
        "SUMIFS has no criteria argument",
        "There is no difference",
      ],
      answer: 1,
      explain: "That inversion is the classic interview trap and the classic production error. COUNTIFS is the third case: it has no sum range at all.",
    },
    {
      q: "You merge a 10-row table to a 3-row lookup on a Left Outer join and get 10 rows. Is that correct?",
      options: ["No, it should be 13", "No, it should be 30", "Yes — a Left Outer join keeps every left-table row", "Only if the key is unique"],
      answer: 2,
      explain: "A Left Outer join returns one row per left-table row, with matched values filled in. It would only multiply rows if the key were duplicated in the right table.",
    },
    {
      q: "Which question are interviewers really asking when they say 'how do you check your work?'",
      options: [
        "Whether you know Excel shortcuts",
        "Whether you would catch your own error before they did",
        "Whether you can write VBA",
        "Whether you have used a specific tool",
      ],
      answer: 1,
      explain: "Technical skill is assumed. What is being assessed is whether you independently reconcile your output — a PivotTable grand total against a plain SUM, a report against its source.",
    },
  ],
  workbookSheets: [
    { name: "Payments", headers: PAYMENTS_HEADERS, rows: PAYMENTS_ROWS },
    { name: "Sales", headers: SALES_HEADERS, rows: SALES_ROWS },
    { name: "Prices", headers: PRICES_HEADERS, rows: PRICES_ROWS },
  ],
  tool: {
    label: "Book Excel training",
    path: "/contact",
    note: "Preparing for a specific role or assessment? A short session on your own data beats general revision.",
  },
  faq: [
    {
      q: "Do employers still test VLOOKUP if XLOOKUP exists?",
      a: "Yes, and more than they test XLOOKUP. Most existing workbooks and most interview panels were formed around VLOOKUP, and the approximate-match default is a genuine test of whether you understand the function rather than having memorised it.",
    },
    {
      q: "Is it better to memorise syntax or understand behaviour?",
      a: "Behaviour. Syntax is looked up constantly in real work and nobody minds. The approximate-match default, blank-cell handling and version availability are the things that cause wrong numbers in production, and they are what separates a strong answer from a recall answer.",
    },
  ],
};

export const XLSX_LESSONS: XlLesson[] = [
  LESSON_1,
  LESSON_2,
  LESSON_3,
  LESSON_4,
  LESSON_5,
  LESSON_6,
  LESSON_7,
  LESSON_8,
  LESSON_9,
  LESSON_10,
];

/** Flat index of every function documented in the lessons, for the cheat sheet. */
export const CHEAT_ENTRIES: XlFunctionDoc[] = XLSX_LESSONS.flatMap((lesson) =>
  lesson.sections.flatMap((section) =>
    section.kind === "functions" ? section.functions : [],
  ),
);

export function lessonBySlug(slug: string): XlLesson | undefined {
  return XLSX_LESSONS.find((lesson) => lesson.slug === slug);
}

export function lessonNeighbours(slug: string) {
  const index = XLSX_LESSONS.findIndex((lesson) => lesson.slug === slug);
  return {
    index,
    prev: index > 0 ? XLSX_LESSONS[index - 1] : null,
    next: index >= 0 && index < XLSX_LESSONS.length - 1 ? XLSX_LESSONS[index + 1] : null,
  };
}
