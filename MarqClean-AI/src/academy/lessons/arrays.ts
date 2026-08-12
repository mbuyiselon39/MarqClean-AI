import type { Lesson } from "./types";

export const arrays: Lesson[] = [
  {
    slug: "unique",
    title: "UNIQUE - a live list of distinct values",
    track: "arrays",
    level: "Intermediate",
    summary:
      "One formula replaces Remove Duplicates and updates itself whenever new rows arrive.",
    objectives: ["Extract unique values", "Combine with SORT and FILTER", "Return values that appear only once"],
    syntax: "=UNIQUE(array; [by_col]; [exactly_once])",
    steps: [
      { title: "The whole trick", detail: "The result spills down as far as it needs.", formula: "=UNIQUE(A2:A500)" },
      { title: "Sort it", detail: "", formula: "=SORT(UNIQUE(A2:A500))" },
      {
        title: "Remove the zeros from empty cells",
        detail: "A range bigger than the data produces 0s; FILTER them away first.",
        formula: "=SORT(UNIQUE(FILTER(A2:A500;A2:A500<>\"\")))",
      },
      {
        title: "Unique combinations of two columns",
        detail: "Give it two columns and it returns unique ROWS.",
        formula: "=UNIQUE(A2:B500)",
      },
      {
        title: "Values that appear exactly once",
        detail: "The third argument flips the logic - great for finding a one-off invoice.",
        formula: "=UNIQUE(A2:A500;FALSE;TRUE)",
      },
      {
        title: "Feed a drop-down",
        detail:
          "Data Validation → List → Source: =$H$2# (the # refers to the whole spilled range) gives a self-maintaining drop-down.",
      },
    ],
    tips: [
      "#SPILL! means something is in the way of the result - clear the cells below.",
      "Reference an entire spill range with the hash: H2#.",
    ],
    practice: {
      task: "Build a sorted, blank-free list of departments from B2:B900.",
      answer: "=SORT(UNIQUE(FILTER(B2:B900;B2:B900<>\"\")))",
    },
    quiz: [
      {
        q: "How do you reference a whole spilled range starting at H2?",
        options: ["H2:H100", "H2#", "H2*", "@H2"],
        answer: 1,
        explain: "The # spill operator always covers the current size of the result.",
      },
    ],
  },
  {
    slug: "sort-function",
    title: "SORT - sort with a formula",
    track: "arrays",
    level: "Intermediate",
    summary:
      "SORT returns a sorted copy that refreshes automatically and never disturbs the source data.",
    objectives: ["Sort ascending and descending", "Sort by any column", "Sort left to right"],
    syntax: "=SORT(array; [sort_index]; [sort_order]; [by_col])",
    steps: [
      { title: "Simple sort", detail: "", formula: "=SORT(A2:A100)" },
      { title: "Sort a table by column 3, descending", detail: "-1 means descending.", formula: "=SORT(A2:D100;3;-1)" },
      { title: "Top 5 only", detail: "Slice the sorted result.", formula: "=TAKE(SORT(A2:D100;3;-1);5)" },
      { title: "Sort left to right", detail: "The fourth argument set to TRUE sorts columns.", formula: "=SORT(B1:M2;2;-1;TRUE)" },
      {
        title: "Sort a filtered list",
        detail: "Chain the functions - filter first, then sort.",
        formula: "=SORT(FILTER(A2:D100;C2:C100=\"North\");4;-1)",
      },
    ],
    practice: {
      task: "Show the 10 biggest orders (amount in column D) as a live list.",
      answer: "=TAKE(SORT(A2:D500;4;-1);10)",
    },
    quiz: [
      {
        q: "What does sort_order -1 mean?",
        options: ["Ascending", "Descending", "By column", "Random"],
        answer: 1,
        explain: "1 = ascending, -1 = descending.",
      },
    ],
  },
  {
    slug: "sortby",
    title: "SORTBY - sort one range by another",
    track: "arrays",
    level: "Advanced",
    summary:
      "SORTBY sorts a list using columns that do not even have to be part of the output - including a custom priority order.",
    objectives: ["Sort by an external column", "Sort by several keys", "Build a custom order"],
    syntax: "=SORTBY(array; by_array1; [sort_order1]; [by_array2]; [sort_order2]; …)",
    steps: [
      { title: "Sort names by their score", detail: "The score column is not shown in the result.", formula: "=SORTBY(A2:A100;B2:B100;-1)" },
      {
        title: "Two sort keys",
        detail: "Region ascending, then amount descending.",
        formula: "=SORTBY(A2:D100;C2:C100;1;D2:D100;-1)",
      },
      {
        title: "Custom priority order",
        detail:
          "MATCH turns your own sequence into sortable numbers - High, Medium, Low instead of alphabetical.",
        formula: "=SORTBY(A2:D100;MATCH(E2:E100;{\"High\";\"Medium\";\"Low\"};0))",
      },
      {
        title: "Shuffle a list",
        detail: "Sort by random numbers for a raffle or a random sample.",
        formula: "=SORTBY(A2:A100;RANDARRAY(99))",
      },
    ],
    practice: {
      task: "Sort a task list by priority High → Medium → Low, then by due date.",
      answer:
        "=SORTBY(A2:D100;MATCH(C2:C100;{\"High\";\"Medium\";\"Low\"};0);1;D2:D100;1)",
    },
    quiz: [
      {
        q: "The difference between SORT and SORTBY is:",
        options: [
          "SORTBY can sort by columns outside the output",
          "SORTBY is faster",
          "SORT cannot descend",
          "None",
        ],
        answer: 0,
        explain: "SORT uses a column index inside the array; SORTBY takes separate by_arrays.",
      },
    ],
  },
  {
    slug: "filter-function",
    title: "FILTER - return only the rows that match",
    track: "arrays",
    level: "Intermediate",
    summary:
      "FILTER is a live AutoFilter that spills its results anywhere you like - the heart of modern Excel dashboards.",
    objectives: ["Filter on one and several conditions", "Handle the empty result", "Combine with SORT and UNIQUE"],
    syntax: "=FILTER(array; include; [if_empty])",
    steps: [
      { title: "One condition", detail: "include is an array of TRUE/FALSE the same height as the array.", formula: "=FILTER(A2:D500;C2:C500=\"North\")" },
      { title: "AND - multiply the conditions", detail: "", formula: "=FILTER(A2:D500;(C2:C500=\"North\")*(D2:D500>5000))" },
      { title: "OR - add them", detail: "", formula: "=FILTER(A2:D500;(C2:C500=\"North\")+(C2:C500=\"South\"))" },
      {
        title: "Always give if_empty",
        detail: "Otherwise no matches produces #CALC!.",
        formula: "=FILTER(A2:D500;C2:C500=F1;\"No matches\")",
      },
      {
        title: "Return selected columns only",
        detail: "Use CHOOSECOLS to pick the columns you want from the filtered result.",
        formula: "=CHOOSECOLS(FILTER(A2:H500;C2:C500=F1);1;3;8)",
      },
      {
        title: "A search-as-you-type report",
        detail:
          "Put a search box in F1 and filter on a contains test - the report updates as you type.",
        formula: "=FILTER(A2:D500;ISNUMBER(SEARCH($F$1;A2:A500));\"-\")",
      },
    ],
    practice: {
      task: "List every open order over 1 000 for the customer typed in F1, sorted newest first.",
      answer:
        "=SORT(FILTER(A2:E500;(A2:A500=$F$1)*(D2:D500>1000)*(E2:E500=\"Open\");\"None\");2;-1)",
    },
    quiz: [
      {
        q: "How do you express AND between two FILTER conditions?",
        options: ["+", "*", "AND()", "&"],
        answer: 1,
        explain: "Multiplying the TRUE/FALSE arrays gives 1 only where both are true.",
      },
    ],
  },
  {
    slug: "randarray",
    title: "RANDARRAY - generate random numbers in bulk",
    track: "arrays",
    level: "Intermediate",
    summary:
      "RANDARRAY builds a whole block of random values in one formula - for test data, sampling and simple simulations.",
    objectives: ["Generate decimals and integers", "Create sample data", "Freeze the results"],
    syntax: "=RANDARRAY([rows]; [columns]; [min]; [max]; [whole_number])",
    steps: [
      { title: "10 rows of decimals between 0 and 1", detail: "", formula: "=RANDARRAY(10)" },
      { title: "Whole numbers in a range", detail: "20 rows × 3 columns of integers from 1 to 100.", formula: "=RANDARRAY(20;3;1;100;TRUE)" },
      {
        title: "Random dates",
        detail: "Generate serial numbers between two dates and format the result as a date.",
        formula: "=RANDARRAY(50;1;DATE(2024;1;1);DATE(2024;12;31);TRUE)",
      },
      {
        title: "Pick random names from a list",
        detail: "Sort the list randomly and take the first n.",
        formula: "=TAKE(SORTBY(A2:A200;RANDARRAY(199));10)",
      },
      {
        title: "Freeze them",
        detail:
          "RANDARRAY recalculates on every change. Copy the result and Paste Special → Values to lock it.",
      },
    ],
    practice: {
      task: "Create 100 practice sales figures between 500 and 5 000.",
      answer: "=RANDARRAY(100;1;500;5000;TRUE)",
    },
    quiz: [
      {
        q: "Which argument makes RANDARRAY return whole numbers?",
        options: ["2nd", "3rd", "4th", "5th"],
        answer: 3,
        explain: "whole_number is the fifth argument.",
      },
    ],
  },
  {
    slug: "sequence",
    title: "SEQUENCE - generate a list of numbers or dates",
    track: "arrays",
    level: "Intermediate",
    summary:
      "SEQUENCE spills a counted series in any shape, which powers date calendars, numbering and amortisation schedules.",
    objectives: ["Create number and date series", "Shape the output", "Use SEQUENCE inside other functions"],
    syntax: "=SEQUENCE(rows; [columns]; [start]; [step])",
    steps: [
      { title: "1 to 10 down a column", detail: "", formula: "=SEQUENCE(10)" },
      { title: "A grid", detail: "5 rows × 3 columns starting at 100, stepping 5.", formula: "=SEQUENCE(5;3;100;5)" },
      { title: "Every day of this month", detail: "Format as dates.", formula: "=SEQUENCE(DAY(EOMONTH(TODAY();0));1;DATE(YEAR(TODAY());MONTH(TODAY());1))" },
      { title: "Twelve month-ends", detail: "", formula: "=EOMONTH(TODAY();SEQUENCE(12;1;0))" },
      {
        title: "Payment numbers for a loan schedule",
        detail: "Feeds directly into PPMT/IPMT.",
        formula: "=SEQUENCE(360)",
      },
    ],
    practice: {
      task: "Generate the 12 month start dates of 2025.",
      answer: "=DATE(2025;SEQUENCE(12);1)",
    },
    quiz: [
      {
        q: "=SEQUENCE(3;2;10;10) produces:",
        options: [
          "10,20 / 30,40 / 50,60",
          "10,10 / 20,20 / 30,30",
          "1..6",
          "An error",
        ],
        answer: 0,
        explain: "3 rows, 2 columns, starting at 10 with a step of 10, filling across then down.",
      },
    ],
  },
  {
    slug: "let",
    title: "LET - name parts of a formula and speed it up",
    track: "arrays",
    level: "Advanced",
    summary:
      "LET stores intermediate results in names, so a long formula becomes readable and each calculation runs once instead of many times.",
    objectives: ["Declare names inside a formula", "Avoid repeating an expensive lookup", "Structure a complex calculation"],
    syntax: "=LET(name1; value1; [name2; value2]; …; calculation)",
    steps: [
      {
        title: "The pattern",
        detail: "Pairs of name and value, then one final calculation that uses them.",
        formula: "=LET(x;A2*B2;x+x*0.2)",
      },
      {
        title: "Do the lookup once",
        detail:
          "Without LET this XLOOKUP would run twice - once for the test and once for the result.",
        formula:
          "=LET(res;XLOOKUP(A2;Ids;Names);IF(ISNA(res);\"Not found\";res))",
      },
      {
        title: "Build up in readable layers",
        detail: "Each name can use the ones declared before it.",
        formula:
          "=LET(data;FILTER(Sales;Region=\"North\");total;SUM(data);avg;AVERAGE(data);total&\" / \"&avg)",
      },
      {
        title: "Formatting rule",
        detail:
          "Press Alt + Enter after each pair inside the formula bar. A 6-line LET is far easier to audit than a one-line monster.",
      },
    ],
    tips: [
      "Names must start with a letter and cannot look like a cell reference (A1 is illegal, amt is fine).",
      "The last argument is always the calculation - a common error is leaving a dangling name pair.",
    ],
    practice: {
      task: "Return the margin % of a row, showing '-' when revenue is zero, using LET.",
      answer:
        "=LET(rev;B2;cost;C2;IF(rev=0;\"-\";(rev-cost)/rev))",
    },
    quiz: [
      {
        q: "The final argument of LET must be:",
        options: ["A name", "A value", "The calculation", "TRUE"],
        answer: 2,
        explain: "Names come in pairs; the last lone argument is the result.",
      },
    ],
  },
  {
    slug: "vstack",
    title: "VSTACK - append tables on top of each other",
    track: "arrays",
    level: "Intermediate",
    summary:
      "VSTACK combines ranges from different sheets into one continuous, live table - a formula alternative to Power Query appending.",
    objectives: ["Stack ranges vertically", "Combine 12 monthly sheets", "Clean the stacked result"],
    syntax: "=VSTACK(array1; [array2]; …)",
    steps: [
      { title: "Stack two ranges", detail: "They should have the same number of columns.", formula: "=VSTACK(Jan!A2:D100;Feb!A2:D100)" },
      { title: "Remove the empty rows", detail: "Stacking fixed ranges leaves zeros; filter them out.", formula: "=FILTER(VSTACK(Jan!A2:D100;Feb!A2:D100);VSTACK(Jan!A2:A100;Feb!A2:A100)<>\"\")" },
      { title: "Stack, then sort", detail: "", formula: "=SORT(VSTACK(Jan!A2:D100;Feb!A2:D100);1)" },
      {
        title: "Add a header row",
        detail: "Stack your header labels on top of the data.",
        formula: "=VSTACK({\"Date\"\\\"Item\"\\\"Qty\"\\\"Value\"};data)",
      },
    ],
    practice: {
      task: "Combine Q1, Q2, Q3 and Q4 sheets into one list.",
      answer: "=VSTACK(Q1!A2:E200;Q2!A2:E200;Q3!A2:E200;Q4!A2:E200) wrapped in FILTER to drop blanks.",
    },
    quiz: [
      {
        q: "Ranges stacked with VSTACK should share:",
        options: ["The same rows", "The same number of columns", "The same sheet", "Nothing"],
        answer: 1,
        explain: "Mismatched widths are padded with #N/A.",
      },
    ],
  },
  {
    slug: "hstack",
    title: "HSTACK - place ranges side by side",
    track: "arrays",
    level: "Intermediate",
    summary:
      "HSTACK joins ranges left to right, which is how you reorder columns or build a custom report layout with a formula.",
    objectives: ["Combine columns", "Reorder a table", "Assemble a report block"],
    syntax: "=HSTACK(array1; [array2]; …)",
    steps: [
      { title: "Join two columns", detail: "", formula: "=HSTACK(A2:A100;D2:D100)" },
      { title: "Reorder a table", detail: "Show column D before column B.", formula: "=HSTACK(A2:A100;D2:D100;B2:B100)" },
      { title: "Add a calculated column", detail: "", formula: "=HSTACK(A2:C100;C2:C100*1.2)" },
      {
        title: "Filter and reorder in one go",
        detail: "CHOOSECOLS is usually cleaner when you are picking from a single table.",
        formula: "=CHOOSECOLS(FILTER(A2:H100;E2:E100=\"Open\");1;5;3)",
      },
    ],
    practice: {
      task: "Build a two-column report of Customer (col A) and Balance (col G).",
      answer: "=HSTACK(A2:A500;G2:G500)  - or =CHOOSECOLS(A2:G500;1;7)",
    },
    quiz: [
      {
        q: "HSTACK combines ranges:",
        options: ["Vertically", "Horizontally", "Diagonally", "Randomly"],
        answer: 1,
        explain: "H = horizontal, side by side.",
      },
    ],
  },
  {
    slug: "chooserows",
    title: "CHOOSEROWS - pick specific rows from an array",
    track: "arrays",
    level: "Advanced",
    summary:
      "CHOOSEROWS extracts the rows you name, in the order you name them - top-N reports, every-other-row samples, reversals.",
    objectives: ["Select rows by number", "Use negative numbers to count from the end", "Build a top-N block"],
    syntax: "=CHOOSEROWS(array; row_num1; [row_num2]; …)",
    steps: [
      { title: "Pick rows 1, 3 and 5", detail: "", formula: "=CHOOSEROWS(A2:D100;1;3;5)" },
      { title: "The last three rows", detail: "Negative numbers count from the bottom.", formula: "=CHOOSEROWS(A2:D100;-3;-2;-1)" },
      { title: "Top 10 after sorting", detail: "", formula: "=CHOOSEROWS(SORT(A2:D100;4;-1);SEQUENCE(10))" },
      { title: "Reverse the order", detail: "", formula: "=CHOOSEROWS(A2:D100;SEQUENCE(ROWS(A2:D100);1;ROWS(A2:D100);-1))" },
      { title: "Simpler siblings", detail: "TAKE and DROP handle the common 'first/last n rows' cases with less typing.", formula: "=TAKE(SORT(A2:D100;4;-1);10)" },
    ],
    practice: {
      task: "Show only the last 5 transactions of a live list.",
      answer: "=TAKE(A2:D500;-5)  - or =CHOOSEROWS(A2:D500;-5;-4;-3;-2;-1)",
    },
    quiz: [
      {
        q: "What does row_num -1 mean?",
        options: ["An error", "The last row", "Reverse the array", "Skip a row"],
        answer: 1,
        explain: "Negative indexes count from the end.",
      },
    ],
  },
  {
    slug: "choosecols",
    title: "CHOOSECOLS - pick specific columns from an array",
    track: "arrays",
    level: "Advanced",
    summary:
      "The column twin of CHOOSEROWS: build a slim report from a 30-column export without deleting anything.",
    objectives: ["Select and reorder columns", "Combine with FILTER"],
    syntax: "=CHOOSECOLS(array; col_num1; [col_num2]; …)",
    steps: [
      { title: "Pick three columns", detail: "", formula: "=CHOOSECOLS(A2:Z500;1;4;9)" },
      { title: "Reorder them", detail: "The output follows the order you list.", formula: "=CHOOSECOLS(A2:Z500;9;1;4)" },
      { title: "Last column", detail: "", formula: "=CHOOSECOLS(A2:Z500;-1)" },
      {
        title: "Slim, filtered report",
        detail: "Filter the rows and then keep only the columns you care about.",
        formula: "=CHOOSECOLS(FILTER(A2:Z500;E2:E500=\"Open\");1;4;9)",
      },
    ],
    practice: {
      task: "From a 26-column export, produce a report with columns C, A and Z in that order.",
      answer: "=CHOOSECOLS(A2:Z500;3;1;26)",
    },
    quiz: [
      {
        q: "CHOOSECOLS(A:Z;-1) returns:",
        options: ["Column A", "Column Z (the last)", "An error", "All columns"],
        answer: 1,
        explain: "-1 is the last column of the array.",
      },
    ],
  },
  {
    slug: "torow",
    title: "TOROW - flatten an array into a single row",
    track: "arrays",
    level: "Advanced",
    summary:
      "TOROW takes a rectangular block and lays every value out along one row, optionally skipping blanks.",
    objectives: ["Flatten a range", "Ignore blanks and errors", "Choose scan direction"],
    syntax: "=TOROW(array; [ignore]; [scan_by_column])",
    steps: [
      { title: "Flatten a block", detail: "", formula: "=TOROW(A2:D10)" },
      { title: "Skip blanks", detail: "ignore = 1 removes blanks, 2 removes errors, 3 removes both.", formula: "=TOROW(A2:D10;1)" },
      { title: "Scan down the columns instead", detail: "", formula: "=TOROW(A2:D10;1;TRUE)" },
      { title: "Join the flattened list", detail: "", formula: "=TEXTJOIN(\", \";TRUE;TOROW(A2:D10;1))" },
    ],
    practice: {
      task: "Turn a 5×4 block of names into a single comma-separated line with no gaps.",
      answer: "=TEXTJOIN(\", \";TRUE;TOROW(A2:D6;1))",
    },
    quiz: [
      {
        q: "TOROW's 'ignore' argument value 3 means:",
        options: ["Nothing", "Blanks only", "Errors only", "Blanks and errors"],
        answer: 3,
        explain: "0 none, 1 blanks, 2 errors, 3 both.",
      },
    ],
  },
  {
    slug: "tocol",
    title: "TOCOL - flatten an array into a single column",
    track: "arrays",
    level: "Advanced",
    summary:
      "TOCOL is the workhorse for turning a wide grid into a tidy single-column list you can filter, count or de-duplicate.",
    objectives: ["Flatten to one column", "Combine with UNIQUE and SORT", "Unpivot a small grid with a formula"],
    syntax: "=TOCOL(array; [ignore]; [scan_by_column])",
    steps: [
      { title: "Flatten", detail: "", formula: "=TOCOL(A2:D10)" },
      { title: "Drop the blanks", detail: "", formula: "=TOCOL(A2:D10;1)" },
      { title: "Unique sorted list from a grid", detail: "Turns a messy grid of tags into a clean list.", formula: "=SORT(UNIQUE(TOCOL(A2:D10;1)))" },
      { title: "Count entries in a grid", detail: "", formula: "=ROWS(TOCOL(A2:D10;1))" },
    ],
    practice: {
      task: "A 4-column grid of skills per employee needs a de-duplicated skill list.",
      answer: "=SORT(UNIQUE(TOCOL(B2:E200;1)))",
    },
    quiz: [
      {
        q: "TOCOL is most useful for:",
        options: [
          "Formatting",
          "Turning a grid into a tidy single-column list",
          "Sorting columns",
          "Charting",
        ],
        answer: 1,
        explain: "It is the formula-based mini-unpivot.",
      },
    ],
  },
  {
    slug: "lambda",
    title: "LAMBDA - build your own Excel function",
    track: "arrays",
    level: "Advanced",
    summary:
      "LAMBDA turns a formula into a reusable named function, so business logic lives in one place instead of in 400 copied cells.",
    objectives: [
      "Write and test a LAMBDA",
      "Save it in the Name Manager",
      "Use the helper functions BYROW, MAP and REDUCE",
    ],
    syntax: "=LAMBDA(parameter1; [parameter2]; …; calculation)",
    steps: [
      {
        title: "Test it inline first",
        detail:
          "Add the arguments in a second set of brackets so you can see the result before naming it.",
        formula: "=LAMBDA(rev;cost;(rev-cost)/rev)(B2;C2)",
      },
      {
        title: "Name it",
        detail:
          "Formulas → Name Manager → New. Name: MARGIN. Refers to: =LAMBDA(rev;cost;IF(rev=0;\"-\";(rev-cost)/rev))",
      },
      { title: "Use it like a built-in", detail: "Colleagues never need to see the logic.", formula: "=MARGIN(B2;C2)" },
      {
        title: "BYROW - apply a calculation to each row",
        detail: "Returns one result per row of the array.",
        formula: "=BYROW(B2:D100;LAMBDA(r;SUM(r)))",
      },
      { title: "MAP - transform every value", detail: "", formula: "=MAP(A2:A100;LAMBDA(v;UPPER(TRIM(v))))" },
      {
        title: "REDUCE - accumulate to one value",
        detail: "Running through the array carrying a total.",
        formula: "=REDUCE(0;A2:A100;LAMBDA(acc;v;acc+v))",
      },
      {
        title: "Document it",
        detail:
          "Use the Comment box in the Name Manager to describe the arguments - it appears in the tooltip.",
      },
    ],
    tips: [
      "LAMBDA requires Microsoft 365.",
      "Keep names in a dedicated 'Functions' workbook or add-in so they can be reused across files.",
    ],
    practice: {
      task: "Create a reusable VAT function that adds 20%.",
      answer: "Name: ADDVAT, Refers to: =LAMBDA(amount;amount*1.2). Use: =ADDVAT(B2)",
    },
    quiz: [
      {
        q: "Where do you save a LAMBDA so it can be reused?",
        options: ["In a cell", "Name Manager", "A macro module", "Power Query"],
        answer: 1,
        explain: "Define it as a name; then it behaves like a built-in function.",
      },
    ],
  },
  {
    slug: "formulatext",
    title: "FORMULATEXT - show the formula, not the result",
    track: "arrays",
    level: "Intermediate",
    summary:
      "FORMULATEXT returns another cell's formula as text - indispensable for documentation, teaching and auditing.",
    objectives: ["Display a formula as text", "Build a self-documenting model", "Audit a workbook"],
    syntax: "=FORMULATEXT(reference)",
    steps: [
      { title: "Show the formula next door", detail: "", formula: "=FORMULATEXT(B2)" },
      { title: "Handle plain values", detail: "A cell with no formula returns #N/A.", formula: "=IFNA(FORMULATEXT(B2);\"(constant)\")" },
      {
        title: "Document a model",
        detail:
          "Put =FORMULATEXT() next to every key calculation on a 'Documentation' sheet so reviewers can read the logic without clicking.",
      },
      {
        title: "Related audit tools",
        detail:
          "Ctrl + ` shows all formulas on the sheet; Formulas → Trace Precedents/Dependents draws the arrows; Evaluate Formula steps through the calculation.",
      },
    ],
    shortcuts: [
      ["Ctrl + `", "Show/hide all formulas"],
      ["Ctrl + [", "Jump to the precedent cells"],
    ],
    practice: {
      task: "Create an audit column showing the formula used in each row of column F.",
      answer: "=IFNA(FORMULATEXT(F2);\"hard-coded\")",
    },
    quiz: [
      {
        q: "FORMULATEXT on a cell containing 100 returns:",
        options: ["100", "\"100\"", "#N/A", "Empty"],
        answer: 2,
        explain: "It only works on cells that contain a formula.",
      },
    ],
  },
];
