import type { Lesson } from "./types";

export const logic: Lesson[] = [
  {
    slug: "if",
    title: "IF - make Excel decide for you",
    track: "logic",
    level: "Beginner",
    summary:
      "IF tests a condition and returns one of two results. It is the single most-used decision function in Excel.",
    objectives: [
      "Write a logical test",
      "Return text, numbers or another formula from IF",
      "Avoid the classic =IF(A1=TRUE;TRUE;FALSE) beginner loop",
    ],
    syntax: "=IF(logical_test; value_if_true; value_if_false)",
    steps: [
      {
        title: "Write the test",
        detail:
          "A test is anything that returns TRUE or FALSE: B2>100, B2=\"Paid\", B2<>\"\", B2>=C2.",
      },
      {
        title: "Add the two outcomes",
        detail:
          "Text results must be in double quotes; numbers must not be.",
        formula: "=IF(B2>=1000;\"Bonus\";\"No bonus\")",
      },
      {
        title: "Return a calculation instead of a label",
        detail: "Either branch can be a formula.",
        formula: "=IF(B2>=1000;B2*5%;0)",
      },
      {
        title: "Show a blank when nothing applies",
        detail: "Two quotation marks with nothing between them produce an empty-looking cell.",
        formula: "=IF(B2=\"\";\"\";B2*C2)",
      },
      {
        title: "Never wrap a test in IF just to get TRUE/FALSE",
        detail: "=IF(B2>100;TRUE;FALSE) is simply =B2>100.",
      },
    ],
    example: {
      caption: "Pass / fail marking",
      headers: ["A (Student)", "B (Score)", "C (Result)"],
      rows: [
        ["Lena", "72", "=IF(B2>=50;\"Pass\";\"Fail\") → Pass"],
        ["Sam", "38", "→ Fail"],
      ],
    },
    tips: [
      "Excel evaluates text comparisons case-insensitively: \"paid\" = \"PAID\" is TRUE.",
      "Use SWITCH or IFS instead of stacking many IFs.",
    ],
    mistakes: [
      "Quotes around numbers: \"100\" is text and will not compare as you expect.",
      "Comparing a number to a text-number imported from a system - use VALUE first.",
    ],
    practice: {
      task: "Flag rows where stock in C2 is below the reorder point in D2 as \"Order\", otherwise blank.",
      answer: "=IF(C2<D2;\"Order\";\"\")",
    },
    quiz: [
      {
        q: "=IF(A1>10;\"Big\";\"Small\") where A1 = 10 returns:",
        options: ["Big", "Small", "TRUE", "#VALUE!"],
        answer: 1,
        explain: "10 is not greater than 10, so the false branch runs. Use >= to include 10.",
      },
    ],
  },
  {
    slug: "nested-if",
    title: "Nested IF - grading and banding with several outcomes",
    track: "logic",
    level: "Intermediate",
    summary:
      "Put an IF inside an IF to handle three or more outcomes - and learn the two modern replacements that are much easier to read.",
    objectives: [
      "Order the conditions correctly",
      "Convert a nested IF to IFS or a lookup table",
    ],
    syntax: "=IF(test1; result1; IF(test2; result2; IF(test3; result3; else)))",
    steps: [
      {
        title: "Sort your bands",
        detail:
          "Excel stops at the first TRUE, so always test from the most extreme value inwards (highest first for >, lowest first for <).",
      },
      {
        title: "Write the nest",
        detail: "Each false branch becomes the next IF.",
        formula:
          "=IF(B2>=90;\"A\";IF(B2>=75;\"B\";IF(B2>=60;\"C\";IF(B2>=50;\"D\";\"F\"))))",
      },
      {
        title: "Use IFS instead (2019/365)",
        detail: "Same logic, half the brackets.",
        formula: "=IFS(B2>=90;\"A\";B2>=75;\"B\";B2>=60;\"C\";B2>=50;\"D\";TRUE;\"F\")",
      },
      {
        title: "Best of all: a lookup table",
        detail:
          "Put the bands in two columns and use an approximate-match lookup. Now non-technical colleagues can change the bands without touching a formula.",
        formula: "=XLOOKUP(B2;Bands[Min];Bands[Grade];\"\";-1)",
      },
    ],
    example: {
      caption: "Band table for the lookup approach",
      headers: ["Min", "Grade"],
      rows: [
        ["0", "F"],
        ["50", "D"],
        ["60", "C"],
        ["75", "B"],
        ["90", "A"],
      ],
    },
    tips: [
      "Excel allows 64 nested IFs. If you need more than three, you almost certainly want a lookup table.",
      "Alt + Enter inside the formula bar puts each IF on its own line - a nested IF becomes readable instantly.",
    ],
    practice: {
      task: "Discount: 0% under 100 units, 5% from 100, 10% from 500, 15% from 1000.",
      answer:
        "=IFS(A2>=1000;15%;A2>=500;10%;A2>=100;5%;TRUE;0)  - or XLOOKUP against a band table with match mode -1.",
    },
    quiz: [
      {
        q: "Why must nested IFs be written from the highest band down?",
        options: [
          "It is faster",
          "Excel returns the first TRUE it meets",
          "IF cannot use >=",
          "Otherwise you get #VALUE!",
        ],
        answer: 1,
        explain:
          "If you tested >=50 first, a score of 95 would already stop there and return the wrong grade.",
      },
    ],
  },
  {
    slug: "and-or",
    title: "AND / OR / NOT - combining conditions",
    track: "logic",
    level: "Beginner",
    summary:
      "AND needs every condition to be true, OR needs at least one. Together with IF they express almost any business rule.",
    objectives: [
      "Combine two or more tests",
      "Use AND/OR inside IF and inside Conditional Formatting",
    ],
    syntax: "=AND(test1; test2; …)   =OR(test1; test2; …)   =NOT(test)",
    steps: [
      { title: "AND - everything must hold", detail: "Both conditions true.", formula: "=AND(B2>=1000;C2=\"North\")" },
      { title: "OR - any condition", detail: "At least one true.", formula: "=OR(C2=\"North\";C2=\"South\")" },
      {
        title: "Put it inside IF",
        detail: "AND/OR only return TRUE/FALSE, so IF turns them into a useful answer.",
        formula: "=IF(AND(B2>=1000;C2=\"North\");\"Bonus\";\"\")",
      },
      {
        title: "Mix them",
        detail: "Brackets decide the priority, exactly like in maths.",
        formula: "=IF(AND(OR(C2=\"North\";C2=\"South\");B2>500);\"Qualifies\";\"No\")",
      },
      {
        title: "Shortcut: * means AND, + means OR",
        detail:
          "(B2>500)*(C2=\"North\") returns 1 or 0 and is the trick that makes SUMPRODUCT work.",
      },
    ],
    example: {
      headers: ["B (Sales)", "C (Region)", "AND(B>=1000;C=\"North\")"],
      rows: [
        ["1 200", "North", "TRUE"],
        ["1 200", "South", "FALSE"],
        ["800", "North", "FALSE"],
      ],
    },
    practice: {
      task: "Mark an order 'Review' when the value is over 10 000 OR the customer is new.",
      answer: "=IF(OR(B2>10000;D2=\"New\");\"Review\";\"\")",
    },
    quiz: [
      {
        q: "=OR(FALSE;FALSE;TRUE) returns:",
        options: ["FALSE", "TRUE", "#VALUE!", "1 and 0"],
        answer: 1,
        explain: "OR needs only one TRUE.",
      },
    ],
  },
  {
    slug: "isblank",
    title: "ISBLANK - test whether a cell is truly empty",
    track: "logic",
    level: "Beginner",
    summary:
      "ISBLANK returns TRUE only for genuinely empty cells - which is why it sometimes 'fails' on cells that just look empty.",
    objectives: [
      "Test for empties",
      "Understand why ISBLANK is FALSE on a formula that returns \"\"",
      "Highlight missing data with Conditional Formatting",
    ],
    syntax: "=ISBLANK(value)",
    steps: [
      { title: "Simple test", detail: "TRUE when the cell has never been filled.", formula: "=ISBLANK(A2)" },
      {
        title: "Use it in IF",
        detail: "Warn about missing input.",
        formula: "=IF(ISBLANK(A2);\"Missing\";\"OK\")",
      },
      {
        title: "The \"\" trap",
        detail:
          "If A2 contains =IF(x;y;\"\") the cell holds an empty string, not nothing. ISBLANK returns FALSE. Use =A2=\"\" instead, which is TRUE for both cases.",
        formula: "=IF(A2=\"\";\"Missing\";\"OK\")",
      },
      {
        title: "Highlight blanks automatically",
        detail:
          "Select the range → Home → Conditional Formatting → New Rule → Use a formula → =ISBLANK(A2) → pick a red fill.",
      },
      { title: "Count them", detail: "COUNTBLANK counts both true blanks and \"\".", formula: "=COUNTBLANK(A2:A500)" },
    ],
    practice: {
      task: "Return \"Enter a date\" when C2 is empty, otherwise the number of days since C2.",
      answer: "=IF(C2=\"\";\"Enter a date\";TODAY()-C2)",
    },
    quiz: [
      {
        q: "A2 contains =IF(B2>0;B2;\"\"). What does =ISBLANK(A2) return?",
        options: ["TRUE", "FALSE", "#N/A", "0"],
        answer: 1,
        explain: "The cell contains a formula, so it is not blank - it only looks blank.",
      },
    ],
  },
  {
    slug: "sumif",
    title: "SUMIF - total the rows that match one condition",
    track: "logic",
    level: "Beginner",
    summary:
      "SUMIF adds only the values whose matching row satisfies your criterion - the fastest way to build a small summary table.",
    objectives: [
      "Write criteria for text, numbers and comparisons",
      "Use the optional sum_range",
      "Use a cell reference as the criterion",
    ],
    syntax: "=SUMIF(range; criteria; [sum_range])",
    steps: [
      {
        title: "Point at the column to test",
        detail: "range is where the condition is checked - for example the Region column.",
      },
      {
        title: "Write the criterion",
        detail:
          "Text goes in quotes, comparisons too: \"North\", \">1000\", \"<>\"&\"\", \"<\"&TODAY().",
      },
      {
        title: "Point at the column to add",
        detail: "sum_range is what actually gets totalled. Omit it to sum the tested range itself.",
        formula: "=SUMIF(C2:C500;\"North\";B2:B500)",
      },
      {
        title: "Refer to a cell so the summary is dynamic",
        detail: "Put the region name in E2 and copy the formula down the summary table.",
        formula: "=SUMIF($C$2:$C$500;E2;$B$2:$B$500)",
      },
      {
        title: "Wildcards",
        detail: "* is any number of characters, ? is one character.",
        formula: "=SUMIF(A2:A500;\"Inv-2024*\";B2:B500)",
      },
    ],
    example: {
      caption: "Summary built with SUMIF",
      headers: ["E (Region)", "F (Total)"],
      rows: [
        ["North", "=SUMIF($C$2:$C$500;E2;$B$2:$B$500)"],
        ["South", "=SUMIF($C$2:$C$500;E3;$B$2:$B$500)"],
      ],
    },
    tips: [
      "Lock the ranges with $ before copying the summary down.",
      "Need two or more conditions? Jump straight to SUMIFS - the argument order is different.",
    ],
    practice: {
      task: "Total all invoices over 5 000 in column B.",
      answer: "=SUMIF(B2:B500;\">5000\")",
    },
    quiz: [
      {
        q: "In =SUMIF(C:C;\"North\";B:B), which column is added?",
        options: ["C", "B", "Both", "Neither"],
        answer: 1,
        explain: "The third argument (sum_range) is what gets summed; C is only tested.",
      },
    ],
  },
  {
    slug: "countif",
    title: "COUNTIF - count the rows that match one condition",
    track: "logic",
    level: "Beginner",
    summary:
      "COUNTIF counts matches and is also the standard way to detect duplicates before they cause problems.",
    objectives: [
      "Count text, number and comparison matches",
      "Find duplicates",
      "Build a frequency table",
    ],
    syntax: "=COUNTIF(range; criteria)",
    steps: [
      { title: "Count a text match", detail: "Case-insensitive.", formula: "=COUNTIF(C2:C500;\"North\")" },
      { title: "Count with a comparison", detail: "Operators go inside the quotes.", formula: "=COUNTIF(B2:B500;\">1000\")" },
      {
        title: "Compare with a cell",
        detail: "Join the operator and the reference with &.",
        formula: "=COUNTIF(B2:B500;\">\"&E1)",
      },
      {
        title: "Find duplicates",
        detail: "Any result above 1 means the value appears more than once.",
        formula: "=COUNTIF($A$2:$A$500;A2)>1",
      },
      {
        title: "Highlight duplicates live",
        detail:
          "Conditional Formatting → New Rule → Use a formula → =COUNTIF($A$2:$A$500;A2)>1 → red fill.",
      },
      { title: "Contains, not equals", detail: "Wildcards make it a 'contains' search.", formula: "=COUNTIF(A2:A500;\"*ltd*\")" },
    ],
    practice: {
      task: "How many customers appear more than once in A2:A500?",
      answer: "=SUMPRODUCT((COUNTIF(A2:A500;A2:A500)>1)*1) counts the duplicated rows.",
    },
    quiz: [
      {
        q: "How do you count cells greater than the value in E1?",
        options: [
          "=COUNTIF(B:B;\">E1\")",
          "=COUNTIF(B:B;\">\"&E1)",
          "=COUNTIF(B:B;E1>)",
          "=COUNTIF(B:B;\">\";E1)",
        ],
        answer: 1,
        explain: "The operator is text and must be concatenated to the cell reference with &.",
      },
    ],
  },
  {
    slug: "averageif",
    title: "AVERAGEIF - conditional mean",
    track: "logic",
    level: "Beginner",
    summary:
      "AVERAGEIF averages only the rows that match a criterion, with the same argument pattern as SUMIF.",
    objectives: ["Average by category", "Exclude zeros and blanks properly"],
    syntax: "=AVERAGEIF(range; criteria; [average_range])",
    steps: [
      { title: "Average one category", detail: "Same shape as SUMIF.", formula: "=AVERAGEIF(C2:C500;\"North\";B2:B500)" },
      { title: "Exclude zeros", detail: "The criterion tests the value column itself.", formula: "=AVERAGEIF(B2:B500;\"<>0\")" },
      { title: "Only positive values", detail: "", formula: "=AVERAGEIF(B2:B500;\">0\")" },
      {
        title: "Empty result",
        detail: "If nothing matches you get #DIV/0!. Wrap it: =IFERROR(…;\"n/a\").",
      },
    ],
    practice: {
      task: "Average order value for customer in E2 only.",
      answer: "=AVERAGEIF($A$2:$A$500;E2;$B$2:$B$500)",
    },
    quiz: [
      {
        q: "AVERAGEIF returns #DIV/0! when:",
        options: [
          "The range contains text",
          "No rows match the criterion",
          "The range is locked",
          "The criterion is a number",
        ],
        answer: 1,
        explain: "With no matching rows there is nothing to divide by.",
      },
    ],
  },
  {
    slug: "sumifs",
    title: "SUMIFS - total with two or more conditions",
    track: "logic",
    level: "Intermediate",
    summary:
      "SUMIFS is SUMIF with unlimited criteria - and the sum range comes first. It handles date ranges beautifully.",
    objectives: [
      "Write multi-criteria totals",
      "Filter between two dates",
      "Build a cross-tab summary with mixed $ locking",
    ],
    syntax: "=SUMIFS(sum_range; criteria_range1; criteria1; [criteria_range2; criteria2]; …)",
    steps: [
      {
        title: "Start with what you are adding",
        detail: "Unlike SUMIF, the sum range is the FIRST argument. This trips everyone up once.",
      },
      {
        title: "Add pairs of range + criterion",
        detail: "All criteria must be true (it is an AND).",
        formula: "=SUMIFS(B2:B500;C2:C500;\"North\";D2:D500;\"2024\")",
      },
      {
        title: "Between two dates",
        detail: "Use the same date column twice with >= and <=.",
        formula:
          "=SUMIFS(B2:B500;A2:A500;\">=\"&$F$1;A2:A500;\"<=\"&$F$2)",
      },
      {
        title: "Cross-tab: regions down, months across",
        detail:
          "Lock the data ranges fully ($C$2:$C$500), lock the row header column ($E2) and the column header row (F$1). Then copy in both directions.",
        formula: "=SUMIFS($B$2:$B$500;$C$2:$C$500;$E2;$D$2:$D$500;F$1)",
      },
      { title: "Not equal to", detail: "", formula: "=SUMIFS(B2:B500;C2:C500;\"<>Cancelled\")" },
    ],
    example: {
      caption: "The cross-tab you can build in 30 seconds",
      headers: ["", "Jan", "Feb", "Mar"],
      rows: [
        ["North", "12 400", "9 800", "15 100"],
        ["South", "8 200", "11 050", "9 400"],
      ],
    },
    tips: [
      "Every criteria_range must be the same size as the sum_range or you get #VALUE!.",
      "SUMIFS ignores case, and treats blank criteria cells as 'zero', which silently returns 0 - validate your inputs.",
    ],
    practice: {
      task: "Total sales for region in E2, between the dates in F1 and F2.",
      answer:
        "=SUMIFS($B$2:$B$500;$C$2:$C$500;$E2;$A$2:$A$500;\">=\"&$F$1;$A$2:$A$500;\"<=\"&$F$2)",
    },
    quiz: [
      {
        q: "Which argument comes first in SUMIFS?",
        options: ["The criteria range", "The sum range", "The criterion", "The table name"],
        answer: 1,
        explain: "SUMIFS(sum_range; criteria_range1; criteria1; …) - the opposite of SUMIF.",
      },
    ],
  },
  {
    slug: "countifs",
    title: "COUNTIFS - count with several conditions",
    track: "logic",
    level: "Intermediate",
    summary:
      "COUNTIFS counts rows that satisfy every condition - ideal for status dashboards and date buckets.",
    objectives: ["Count with 2+ criteria", "Count within a date window", "Count blanks and non-blanks"],
    syntax: "=COUNTIFS(criteria_range1; criteria1; [criteria_range2; criteria2]; …)",
    steps: [
      { title: "Two conditions", detail: "Open orders in the North.", formula: "=COUNTIFS(C2:C500;\"North\";E2:E500;\"Open\")" },
      {
        title: "Date bucket",
        detail: "Same column twice.",
        formula: "=COUNTIFS(A2:A500;\">=\"&$F$1;A2:A500;\"<\"&$F$2)",
      },
      { title: "Count non-blanks with a condition", detail: "", formula: "=COUNTIFS(C2:C500;\"North\";B2:B500;\"<>\")" },
      {
        title: "Build a status dashboard",
        detail:
          "Put the statuses down column H and copy one formula: =COUNTIFS($E$2:$E$500;H2). Add a percentage column and you have a report.",
      },
    ],
    practice: {
      task: "Count overdue invoices: due date before today and status not 'Paid'.",
      answer: "=COUNTIFS(D2:D500;\"<\"&TODAY();E2:E500;\"<>Paid\")",
    },
    quiz: [
      {
        q: "COUNTIFS combines its criteria with:",
        options: ["OR", "AND", "XOR", "NOT"],
        answer: 1,
        explain: "Every criterion must be true for the row to be counted.",
      },
    ],
  },
  {
    slug: "maxifs",
    title: "MAXIFS - the largest value that meets conditions",
    track: "logic",
    level: "Intermediate",
    summary:
      "MAXIFS (Excel 2019 and 365) gives the highest value inside a filtered subset without any array gymnastics.",
    objectives: ["Conditional maximum", "Fallback formula for older Excel"],
    syntax: "=MAXIFS(max_range; criteria_range1; criteria1; …)",
    steps: [
      { title: "One condition", detail: "Biggest sale in the North.", formula: "=MAXIFS(B2:B500;C2:C500;\"North\")" },
      {
        title: "Several conditions",
        detail: "Add pairs, exactly like SUMIFS.",
        formula: "=MAXIFS(B2:B500;C2:C500;\"North\";A2:A500;\">=\"&DATE(2024;1;1))",
      },
      {
        title: "Latest date per customer",
        detail: "Dates are numbers, so MAXIFS finds the most recent order date.",
        formula: "=MAXIFS($A$2:$A$500;$D$2:$D$500;F2)",
      },
      {
        title: "Older Excel fallback",
        detail: "Array formula entered with Ctrl + Shift + Enter in Excel 2016 and earlier.",
        formula: "=MAX(IF(C2:C500=\"North\";B2:B500))",
      },
    ],
    practice: {
      task: "Return the most recent order date for the customer named in F2.",
      answer: "=MAXIFS($A$2:$A$500;$D$2:$D$500;F2) formatted as a date",
    },
    quiz: [
      {
        q: "MAXIFS returns 0 when:",
        options: [
          "The range has text",
          "No rows match the criteria",
          "There are more than 2 criteria",
          "The range is a Table",
        ],
        answer: 1,
        explain: "With no matches it returns 0 - wrap in IF(COUNTIFS(...)=0;\"n/a\";…) if that is misleading.",
      },
    ],
  },
  {
    slug: "minifs",
    title: "MINIFS - the smallest value that meets conditions",
    track: "logic",
    level: "Intermediate",
    summary:
      "MINIFS mirrors MAXIFS and is the cleanest way to find a first date, a cheapest price or a best quote.",
    objectives: ["Conditional minimum", "Ignore zeros while finding a minimum"],
    syntax: "=MINIFS(min_range; criteria_range1; criteria1; …)",
    steps: [
      { title: "Cheapest price for a product", detail: "", formula: "=MINIFS(C2:C500;A2:A500;F2)" },
      { title: "Ignore zeros", detail: "Test the value column itself.", formula: "=MINIFS(C2:C500;C2:C500;\">0\")" },
      { title: "First order date per customer", detail: "", formula: "=MINIFS($A$2:$A$500;$D$2:$D$500;F2)" },
      {
        title: "Return the supplier too",
        detail: "Wrap MINIFS in a lookup on two criteria.",
        formula: "=XLOOKUP(1;(A2:A500=F2)*(C2:C500=MINIFS(C2:C500;A2:A500;F2));B2:B500)",
      },
    ],
    practice: {
      task: "Cheapest non-zero quote for the item in F2 (items in A, prices in C).",
      answer: "=MINIFS(C2:C500;A2:A500;F2;C2:C500;\">0\")",
    },
    quiz: [
      {
        q: "Which argument comes first in MINIFS?",
        options: ["criteria_range", "min_range", "criteria", "range to ignore"],
        answer: 1,
        explain: "Like SUMIFS, the value range comes first.",
      },
    ],
  },
  {
    slug: "iferror",
    title: "IFERROR - clean, professional-looking results",
    track: "logic",
    level: "Beginner",
    summary:
      "IFERROR replaces any error with a value you choose - but used carelessly it also hides genuine mistakes.",
    objectives: [
      "Suppress #N/A, #DIV/0! and #VALUE!",
      "Use IFNA when only a missing lookup should be hidden",
      "Count the errors you are hiding",
    ],
    syntax: "=IFERROR(value; value_if_error)",
    steps: [
      { title: "Wrap the formula", detail: "The first argument is your normal formula.", formula: "=IFERROR(B2/C2;\"\")" },
      { title: "Give a helpful message", detail: "", formula: "=IFERROR(VLOOKUP(A2;List;2;0);\"Not registered\")" },
      {
        title: "Prefer IFNA for lookups",
        detail:
          "IFNA hides only #N/A, so a typo in your range still shows up as a real error instead of being swallowed.",
        formula: "=IFNA(XLOOKUP(A2;Ids;Names);\"Unknown\")",
      },
      {
        title: "Monitor what you hid",
        detail: "Put a counter in a corner cell so silent problems become visible.",
        formula: "=SUMPRODUCT(--ISERROR(D2:D500))",
      },
    ],
    tips: [
      "IFERROR recalculates the formula twice internally on old versions - on huge models prefer IFNA or restructure.",
      "Errors in a chart source are plotted as gaps; returning NA() on purpose is sometimes better than \"\".",
    ],
    practice: {
      task: "Divide B2 by C2, show a dash when C2 is empty.",
      answer: "=IFERROR(B2/C2;\"-\")",
    },
    quiz: [
      {
        q: "Which function hides ONLY the #N/A error?",
        options: ["IFERROR", "IFNA", "ISERROR", "ERROR.TYPE"],
        answer: 1,
        explain: "IFNA is the surgical version of IFERROR.",
      },
    ],
  },
  {
    slug: "sumproduct",
    title: "SUMPRODUCT - multiply, condition and total in one cell",
    track: "logic",
    level: "Advanced",
    summary:
      "SUMPRODUCT multiplies arrays element by element and adds the results. It does weighted averages and conditional maths in every Excel version.",
    objectives: [
      "Compute qty × price totals",
      "Use conditions as 1/0 multipliers",
      "Build a weighted average",
    ],
    syntax: "=SUMPRODUCT(array1; [array2]; …)",
    steps: [
      {
        title: "The classic use",
        detail: "Multiplies each quantity by each price and totals the lot.",
        formula: "=SUMPRODUCT(B2:B100;C2:C100)",
      },
      {
        title: "Conditions become 1 and 0",
        detail:
          "(C2:C100=\"North\") gives an array of TRUE/FALSE. Multiplying arrays converts them to 1/0, so only matching rows survive.",
        formula: "=SUMPRODUCT((C2:C100=\"North\")*(B2:B100))",
      },
      {
        title: "Two conditions = AND",
        detail: "Multiply the conditions together.",
        formula: "=SUMPRODUCT((C2:C100=\"North\")*(D2:D100=\"2024\")*(B2:B100))",
      },
      {
        title: "OR with +",
        detail: "Add conditions, then clamp with >0 to avoid double counting.",
        formula: "=SUMPRODUCT(((C2:C100=\"North\")+(C2:C100=\"South\")>0)*(B2:B100))",
      },
      {
        title: "Weighted average",
        detail: "Value × weight, divided by total weight.",
        formula: "=SUMPRODUCT(B2:B100;C2:C100)/SUM(C2:C100)",
      },
      {
        title: "Count with a calculation inside",
        detail: "Count orders where the margin is below 10%. -- converts TRUE/FALSE to 1/0.",
        formula: "=SUMPRODUCT(--((B2:B100-C2:C100)/B2:B100<0.1))",
      },
    ],
    tips: [
      "All arrays must be the same size - full-column references make SUMPRODUCT slow, so use B2:B100 not B:B.",
      "SUMPRODUCT works in every version without Ctrl + Shift + Enter, which is why it survives in shared workbooks.",
    ],
    practice: {
      task: "Total the value of North orders where quantity is above 10 (qty in B, price in C, region in D).",
      answer: "=SUMPRODUCT((D2:D100=\"North\")*(B2:B100>10)*B2:B100*C2:C100)",
    },
    quiz: [
      {
        q: "What does the double minus (--) do in SUMPRODUCT?",
        options: [
          "Makes numbers negative",
          "Converts TRUE/FALSE into 1/0",
          "Rounds the result",
          "Nothing",
        ],
        answer: 1,
        explain: "Two negations coerce booleans into numbers so they can be summed.",
      },
    ],
  },
];
