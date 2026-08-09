import type { Lesson } from "./types";

export const lookup: Lesson[] = [
  {
    slug: "vlookup",
    title: "VLOOKUP - pull data from another table",
    track: "lookup",
    level: "Beginner",
    summary:
      "VLOOKUP searches down the first column of a table and returns a value from a column you count across.",
    objectives: [
      "Write all four arguments correctly",
      "Always use FALSE for exact match",
      "Fix the three classic VLOOKUP failures",
    ],
    syntax: "=VLOOKUP(lookup_value; table_array; col_index_num; [range_lookup])",
    steps: [
      {
        title: "Choose what you are looking for",
        detail: "lookup_value is usually a code or ID in the current row, e.g. A2.",
      },
      {
        title: "Select the table - starting at the lookup column",
        detail:
          "This is the rule everyone forgets: the value you search for must sit in the FIRST column of table_array. Lock it with F4.",
        formula: "$H$2:$K$500",
      },
      {
        title: "Count the column to return",
        detail:
          "Counting starts at 1 on the first column of your table_array, not on column A of the sheet.",
      },
      {
        title: "Type FALSE (or 0) for exact match",
        detail:
          "Leaving the last argument out means TRUE = approximate match on unsorted data, which silently returns wrong answers.",
        formula: "=VLOOKUP(A2;$H$2:$K$500;3;FALSE)",
      },
      {
        title: "Handle not-found",
        detail: "Wrap in IFNA so missing codes read nicely.",
        formula: "=IFNA(VLOOKUP(A2;$H$2:$K$500;3;FALSE);\"Not found\")",
      },
      {
        title: "Make it survive inserted columns",
        detail: "Replace the hard-coded 3 with MATCH on the header.",
        formula: "=VLOOKUP(A2;$H$2:$K$500;MATCH($C$1;$H$1:$K$1;0);FALSE)",
      },
    ],
    example: {
      caption: "Price list on the right, order sheet on the left",
      headers: ["A (SKU)", "B (Price)", "", "H (SKU)", "I (Name)", "J (Price)"],
      rows: [
        ["S-102", "=VLOOKUP(A2;$H$2:$J$50;3;0) → 45", "", "S-101", "Mouse", "18"],
        ["S-101", "→ 18", "", "S-102", "Keyboard", "45"],
      ],
    },
    mistakes: [
      "#N/A: the value genuinely is not there, or one side is text and the other a number, or there are trailing spaces - wrap the lookup value in TRIM.",
      "Wrong value returned: you forgot FALSE.",
      "#REF!: your col_index_num is bigger than the table.",
      "Cannot look left: VLOOKUP only looks right. Use XLOOKUP or INDEX/MATCH.",
    ],
    practice: {
      task: "Employee IDs in A2:A200, staff table in Staff!A:D with department in column 4. Return the department, showing 'Unknown' if missing.",
      answer: "=IFNA(VLOOKUP(A2;Staff!$A$2:$D$999;4;FALSE);\"Unknown\")",
    },
    quiz: [
      {
        q: "Why does VLOOKUP return a wrong (not #N/A) result?",
        options: [
          "The table is too big",
          "The last argument was left out, so it did an approximate match",
          "The lookup value is text",
          "The columns are locked",
        ],
        answer: 1,
        explain: "Omitting range_lookup defaults to TRUE - approximate match on unsorted data.",
      },
      {
        q: "VLOOKUP can return a column that is…",
        options: ["To the left", "To the right", "Either", "On another sheet only"],
        answer: 1,
        explain: "Only to the right of the lookup column. Use XLOOKUP or INDEX/MATCH to look left.",
      },
    ],
  },
  {
    slug: "xlookup",
    title: "XLOOKUP - the modern replacement for VLOOKUP",
    track: "lookup",
    level: "Intermediate",
    summary:
      "XLOOKUP takes two ranges instead of a column number, looks in any direction, and has built-in not-found handling.",
    objectives: [
      "Replace VLOOKUP and INDEX/MATCH",
      "Use if_not_found and match_mode",
      "Return several columns at once",
    ],
    syntax:
      "=XLOOKUP(lookup_value; lookup_array; return_array; [if_not_found]; [match_mode]; [search_mode])",
    steps: [
      {
        title: "Point at the two columns",
        detail: "No counting, no locking the whole table.",
        formula: "=XLOOKUP(A2;Products[SKU];Products[Price])",
      },
      {
        title: "Add the not-found message",
        detail: "The fourth argument replaces the IFERROR wrapper.",
        formula: "=XLOOKUP(A2;Products[SKU];Products[Price];\"Not found\")",
      },
      {
        title: "Look left - it just works",
        detail: "The return array can sit anywhere, before or after the lookup array.",
      },
      {
        title: "Approximate match for bands",
        detail: "match_mode -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard.",
        formula: "=XLOOKUP(B2;Bands[Min];Bands[Discount];0;-1)",
      },
      {
        title: "Return the last match",
        detail: "search_mode -1 searches bottom-up - the newest price, the latest status.",
        formula: "=XLOOKUP(A2;Log[Item];Log[Status];\"\";0;-1)",
      },
      {
        title: "Return a whole row of columns",
        detail: "Give it a multi-column return array and the result spills across.",
        formula: "=XLOOKUP(A2;Products[SKU];Products[[Name]:[Price]])",
      },
      {
        title: "Two-way lookup",
        detail: "Nest XLOOKUP inside XLOOKUP for row and column.",
        formula: "=XLOOKUP($A2;$H$2:$H$50;XLOOKUP(B$1;$I$1:$M$1;$I$2:$M$50))",
      },
    ],
    tips: [
      "XLOOKUP needs Microsoft 365 or Excel 2021. In 2019 or earlier use INDEX/MATCH.",
      "Match on two criteria: =XLOOKUP(1;(A2:A100=F1)*(B2:B100=F2);C2:C100).",
    ],
    practice: {
      task: "Return the price for the SKU in A2 from a table where SKU is in column J and price in column G, showing 0 when missing.",
      answer: "=XLOOKUP(A2;$J$2:$J$500;$G$2:$G$500;0)",
    },
    quiz: [
      {
        q: "Which XLOOKUP argument handles a missing value?",
        options: ["2nd", "3rd", "4th", "5th"],
        answer: 2,
        explain: "if_not_found is the fourth argument.",
      },
    ],
  },
  {
    slug: "hlookup",
    title: "HLOOKUP - lookup across a row",
    track: "lookup",
    level: "Intermediate",
    summary:
      "HLOOKUP is VLOOKUP rotated 90°: it searches the top row and returns a value from a row you count down.",
    objectives: ["Use HLOOKUP on a horizontal table", "Know when to transpose instead"],
    syntax: "=HLOOKUP(lookup_value; table_array; row_index_num; [range_lookup])",
    steps: [
      {
        title: "Layout check",
        detail: "Your labels (months, years, sizes) must run across the FIRST row of the table.",
      },
      {
        title: "Write it with exact match",
        detail: "Row 1 holds the months, row 4 holds the value you want.",
        formula: "=HLOOKUP($B$1;$A$1:$M$20;4;FALSE)",
      },
      {
        title: "Two-way lookup with MATCH",
        detail: "Let MATCH find the row number so the formula survives new rows.",
        formula: "=HLOOKUP(B$1;$A$1:$M$20;MATCH($A2;$A$1:$A$20;0);FALSE)",
      },
      {
        title: "Modern alternative",
        detail: "XLOOKUP does rows and columns with the same syntax, so it replaces HLOOKUP entirely.",
        formula: "=XLOOKUP(B$1;$B$1:$M$1;$B$4:$M$4)",
      },
    ],
    practice: {
      task: "Months in B1:M1, 'Revenue' row is row 5. Return March revenue.",
      answer: "=HLOOKUP(\"Mar\";$B$1:$M$20;5;FALSE)  - or =XLOOKUP(\"Mar\";$B$1:$M$1;$B$5:$M$5)",
    },
    quiz: [
      {
        q: "HLOOKUP searches:",
        options: ["The first column", "The first row", "The last row", "Any row"],
        answer: 1,
        explain: "H = horizontal: it searches across the top row of the table.",
      },
    ],
  },
  {
    slug: "index",
    title: "INDEX - return the value at a position",
    track: "lookup",
    level: "Intermediate",
    summary:
      "INDEX returns the value at a given row (and optional column) inside a range. On its own it is simple; paired with MATCH it is the most flexible lookup in Excel.",
    objectives: ["Use INDEX on one and two dimensions", "Return a whole row or column"],
    syntax: "=INDEX(array; row_num; [column_num])",
    steps: [
      { title: "One-dimensional", detail: "The 5th name in the list.", formula: "=INDEX(A2:A100;5)" },
      { title: "Two-dimensional", detail: "Row 5, column 3 of the block.", formula: "=INDEX(A2:D100;5;3)" },
      {
        title: "Return an entire row or column",
        detail: "Use 0 for the dimension you want in full - it spills in 365.",
        formula: "=INDEX(A2:D100;5;0)",
      },
      {
        title: "Dynamic ranges",
        detail:
          "INDEX can return a reference, so A2:INDEX(A2:A1000;COUNTA(A2:A1000)) is a range that grows with the data.",
      },
    ],
    practice: {
      task: "Return the value in the 3rd row and 2nd column of the block B5:F20.",
      answer: "=INDEX(B5:F20;3;2)",
    },
    quiz: [
      {
        q: "=INDEX(A2:A10;0) with a 365 version returns:",
        options: ["#REF!", "The whole range spilled", "0", "The first cell"],
        answer: 1,
        explain: "0 means 'all rows', which spills the entire column.",
      },
    ],
  },
  {
    slug: "match",
    title: "MATCH - find the position of a value",
    track: "lookup",
    level: "Intermediate",
    summary:
      "MATCH returns where something sits in a list. It powers INDEX, makes VLOOKUP column numbers dynamic and validates data.",
    objectives: ["Use the three match types", "Combine MATCH with INDEX", "Test whether a value exists"],
    syntax: "=MATCH(lookup_value; lookup_array; [match_type])",
    steps: [
      {
        title: "Exact match with 0",
        detail: "Always pass 0 unless you deliberately want a banded match.",
        formula: "=MATCH(\"Keyboard\";A2:A100;0)",
      },
      {
        title: "Match types explained",
        detail:
          "0 = exact (any order). 1 = largest value ≤ lookup, list must be ascending. -1 = smallest value ≥ lookup, list must be descending.",
      },
      {
        title: "Find a column number for VLOOKUP",
        detail: "Now inserting a column will not break the lookup.",
        formula: "=MATCH($C$1;$H$1:$K$1;0)",
      },
      {
        title: "Does the value exist?",
        detail: "Great for data validation checks.",
        formula: "=ISNUMBER(MATCH(A2;List!$A:$A;0))",
      },
      {
        title: "Match on two criteria",
        detail: "Multiply the conditions to build a 1/0 array (365 spills automatically).",
        formula: "=MATCH(1;(A2:A100=F1)*(B2:B100=F2);0)",
      },
    ],
    practice: {
      task: "Which row of A2:A500 holds the code in F1? Return 'Missing' if not present.",
      answer: "=IFNA(MATCH(F1;A2:A500;0);\"Missing\")",
    },
    quiz: [
      {
        q: "Which match_type requires the list to be sorted ascending?",
        options: ["0", "1", "-1", "None"],
        answer: 1,
        explain: "1 (or omitted) does an approximate match and assumes an ascending list.",
      },
    ],
  },
  {
    slug: "index-match",
    title: "INDEX MATCH - the flexible lookup combination",
    track: "lookup",
    level: "Advanced",
    summary:
      "INDEX MATCH looks in any direction, never breaks when columns move, and works in every version of Excel.",
    objectives: [
      "Build the combination step by step",
      "Do a two-way lookup",
      "Do a two-criteria lookup",
    ],
    syntax: "=INDEX(return_column; MATCH(lookup_value; lookup_column; 0))",
    steps: [
      {
        title: "Step 1 - the answer column",
        detail: "Tell INDEX where the answer lives: =INDEX($C$2:$C$500;…)",
      },
      {
        title: "Step 2 - find the row",
        detail: "MATCH tells INDEX which row to take.",
        formula: "=MATCH($F2;$A$2:$A$500;0)",
      },
      {
        title: "Step 3 - put them together",
        detail: "The lookup column can be to the right of the answer column - no problem.",
        formula: "=INDEX($C$2:$C$500;MATCH($F2;$A$2:$A$500;0))",
      },
      {
        title: "Two-way lookup",
        detail: "A second MATCH finds the column, so you can read any cell of a matrix.",
        formula:
          "=INDEX($B$2:$M$50;MATCH($A2;$A$2:$A$50;0);MATCH(B$1;$B$1:$M$1;0))",
      },
      {
        title: "Two criteria",
        detail:
          "Multiply the two condition arrays. In 365 just press Enter; in older versions press Ctrl + Shift + Enter.",
        formula:
          "=INDEX($D$2:$D$500;MATCH(1;($A$2:$A$500=F1)*($B$2:$B$500=F2);0))",
      },
      {
        title: "Handle missing values",
        detail: "",
        formula: "=IFNA(INDEX($C$2:$C$500;MATCH($F2;$A$2:$A$500;0));\"-\")",
      },
    ],
    tips: [
      "INDEX MATCH only reads two columns, so it recalculates faster than VLOOKUP over a 20-column table.",
      "Ranges must be the same height - a common cause of an off-by-one wrong answer is starting one range at row 1 and the other at row 2.",
    ],
    practice: {
      task: "Names are in column D, IDs in column G. Return the name for the ID in F2.",
      answer: "=IFNA(INDEX($D$2:$D$500;MATCH(F2;$G$2:$G$500;0));\"Not found\")",
    },
    quiz: [
      {
        q: "The main advantage of INDEX MATCH over VLOOKUP is:",
        options: [
          "It is shorter",
          "It can return columns to the left and survives inserted columns",
          "It sorts the data",
          "It works only in 365",
        ],
        answer: 1,
        explain: "It references columns directly instead of counting across from the lookup column.",
      },
    ],
  },
];
