import type { Lesson } from "./types";

export const foundations: Lesson[] = [
  {
    slug: "excel-basics",
    title: "Excel Basics - the interface, cells and your first formula",
    track: "foundations",
    level: "Beginner",
    summary:
      "Learn what a workbook, worksheet, cell, range and reference actually are, then write your very first formula.",
    objectives: [
      "Name every part of the Excel window",
      "Write a reference such as A1, A1:A10 and Sheet2!B3",
      "Type a formula that starts with = and updates automatically",
      "Understand the difference between a value, text and a formula",
    ],
    steps: [
      {
        title: "Open a workbook and meet the grid",
        detail:
          "A file is a workbook. Each tab at the bottom is a worksheet. A worksheet is a grid of columns (letters) and rows (numbers). The box where a column and a row cross is a cell, and the cell address is column + row - for example C5.",
      },
      {
        title: "Type data into a cell",
        detail:
          "Click a cell, type, then press Enter (moves down) or Tab (moves right). Numbers align right, text aligns left. That alignment is your first sanity check: if a number is stuck on the left it is really text.",
      },
      {
        title: "Start every formula with =",
        detail:
          "In cell C2 type =A2+B2 and press Enter. Excel shows the result but stores the formula. Change A2 and C2 recalculates instantly.",
        formula: "=A2+B2",
      },
      {
        title: "Use a range instead of many cells",
        detail:
          "A colon means 'through'. A2:A10 is nine cells. Functions take ranges, so =SUM(A2:A10) is far better than =A2+A3+A4…",
        formula: "=SUM(A2:A10)",
      },
      {
        title: "Copy the formula down",
        detail:
          "Select C2 and double-click the small square at the bottom-right (the fill handle), or press Ctrl + D after selecting C2:C20. References shift by row automatically - that is a relative reference.",
      },
      {
        title: "Lock a reference with $",
        detail:
          "Press F4 while editing a reference to cycle A2 → $A$2 → A$2 → $A2. Use $ when a formula must always point at the same cell, such as a VAT rate in $B$1.",
        formula: "=A2*$B$1",
      },
    ],
    example: {
      caption: "Type this into a blank sheet and follow along",
      headers: ["", "A (Item)", "B (Qty)", "C (Price)", "D (Total)"],
      rows: [
        ["2", "Keyboard", "3", "45", "=B2*C2 → 135"],
        ["3", "Mouse", "5", "18", "=B3*C3 → 90"],
        ["4", "Monitor", "2", "150", "=B4*C4 → 300"],
        ["5", "Total", "", "", "=SUM(D2:D4) → 525"],
      ],
    },
    tips: [
      "Press Ctrl + ` (backtick) to show every formula on the sheet at once - the fastest way to audit someone else's work.",
      "Ctrl + Z undoes, F2 edits in place, Esc cancels an edit before it is committed.",
      "Turn any block of data into a Table with Ctrl + T so formulas and charts grow with the data.",
    ],
    shortcuts: [
      ["Ctrl + Arrow", "Jump to the edge of the data"],
      ["Ctrl + Shift + Arrow", "Select to the edge of the data"],
      ["F2", "Edit the active cell"],
      ["F4", "Cycle absolute/relative references"],
      ["Ctrl + Enter", "Fill every selected cell at once"],
      ["Alt + =", "AutoSum"],
    ],
    mistakes: [
      "Forgetting the = sign - Excel then treats the whole thing as text.",
      "Typing 12 kg into a number cell; use a custom number format instead.",
      "Merged cells: they break sorting, filtering and pivots. Use 'Center Across Selection'.",
    ],
    practice: {
      task: "Put 10 in A1, 4 in A2. In A3 return the total, in A4 the difference, in A5 the product and in A6 A1 divided by A2.",
      answer: "=A1+A2  |  =A1-A2  |  =A1*A2  |  =A1/A2  → 14, 6, 40, 2.5",
    },
    quiz: [
      {
        q: "Which character must every Excel formula start with?",
        options: ["+", "=", "#", ":"],
        answer: 1,
        explain:
          "= tells Excel to calculate. (+ and - also work as legacy Lotus entries but Excel converts them to =.)",
      },
      {
        q: "What does $B$1 mean when you copy a formula?",
        options: [
          "The cell contains currency",
          "The reference never changes",
          "Only the column changes",
          "It refers to another workbook",
        ],
        answer: 1,
        explain:
          "Both the column and the row are locked, so the reference stays on B1 wherever you copy it.",
      },
    ],
  },
  {
    slug: "sum",
    title: "SUM - add up a range in one keystroke",
    track: "foundations",
    level: "Beginner",
    summary:
      "SUM adds numbers, ignores text and blanks, and has a keyboard shortcut that can total a whole table at once.",
    objectives: [
      "Write =SUM() over a range and over separate ranges",
      "Total an entire table with Alt + =",
      "Know when to use SUBTOTAL instead of SUM",
    ],
    syntax: "=SUM(number1; [number2]; …)",
    steps: [
      {
        title: "Select the cell under your numbers",
        detail:
          "Click the first empty cell below the column you want to total.",
      },
      {
        title: "Press Alt + =",
        detail:
          "Excel guesses the range above and writes the SUM for you. Press Enter to accept.",
        formula: "=SUM(B2:B13)",
      },
      {
        title: "Total the whole table at once",
        detail:
          "Select the data plus the empty row below AND the empty column to the right, then press Alt + =. Excel fills every total and the grand total corner in one go.",
      },
      {
        title: "Add non-adjacent ranges",
        detail:
          "Separate the arguments with a semicolon (comma on US/UK versions). Up to 255 arguments are allowed.",
        formula: "=SUM(B2:B13;D2:D13;F5)",
      },
      {
        title: "Sum only the visible rows",
        detail:
          "If your list is filtered, SUM still adds hidden rows. Use SUBTOTAL with function number 109, or press Alt + = inside a filtered Table and Excel writes it for you.",
        formula: "=SUBTOTAL(109;B2:B13)",
      },
    ],
    example: {
      caption: "Monthly sales",
      headers: ["A (Month)", "B (Sales)"],
      rows: [
        ["Jan", "12 400"],
        ["Feb", "9 850"],
        ["Mar", "15 300"],
        ["Total", "=SUM(B2:B4) → 37 550"],
      ],
    },
    formulas: [
      { label: "Simple total", code: "=SUM(B2:B13)" },
      { label: "Whole column (safe in a Table)", code: "=SUM(Sales[Amount])" },
      { label: "Running total", code: "=SUM($B$2:B2)" },
      { label: "Visible rows only", code: "=SUBTOTAL(109;B2:B13)" },
    ],
    tips: [
      "SUM ignores text and empty cells, so it never errors just because a cell says 'n/a'.",
      "A running total is just a SUM with the first reference locked: =SUM($B$2:B2) copied down.",
      "If SUM returns 0, your numbers are stored as text - check for left alignment and use VALUE or Text to Columns.",
    ],
    shortcuts: [["Alt + =", "Insert AutoSum"]],
    practice: {
      task: "You have costs in C2:C40 and a filter applied. Write the formula that totals only what you can see.",
      answer: "=SUBTOTAL(109;C2:C40)",
    },
    quiz: [
      {
        q: "Which formula totals only the rows left visible by a filter?",
        options: [
          "=SUM(C2:C40)",
          "=SUBTOTAL(109;C2:C40)",
          "=SUMIF(C2:C40;\">0\")",
          "=COUNT(C2:C40)",
        ],
        answer: 1,
        explain: "109 = SUM, ignoring hidden and filtered rows.",
      },
    ],
  },
  {
    slug: "max",
    title: "MAX - find the largest number",
    track: "foundations",
    level: "Beginner",
    summary:
      "MAX returns the biggest value in a range. Combine it with a lookup to also show which record it belongs to.",
    objectives: [
      "Return the highest value in a range",
      "Combine MAX with INDEX/MATCH to get the matching label",
      "Use MAX to build a 'never below zero' formula",
    ],
    syntax: "=MAX(number1; [number2]; …)",
    steps: [
      {
        title: "Write the formula",
        detail: "Point MAX at your numeric range.",
        formula: "=MAX(B2:B100)",
      },
      {
        title: "Show which item scored the maximum",
        detail:
          "Use MATCH to find the position of the maximum and INDEX to bring back the name.",
        formula: "=INDEX(A2:A100;MATCH(MAX(B2:B100);B2:B100;0))",
      },
      {
        title: "Use MAX as a floor",
        detail:
          "=MAX(0;B2-C2) never returns a negative number - perfect for stock levels or bonuses.",
        formula: "=MAX(0;B2-C2)",
      },
      {
        title: "Conditional maximum",
        detail:
          "For 'the largest sale in the North region' use MAXIFS (365/2019) - covered in its own lesson.",
        formula: "=MAXIFS(B2:B100;C2:C100;\"North\")",
      },
    ],
    example: {
      caption: "Best performing rep",
      headers: ["A (Rep)", "B (Sales)"],
      rows: [
        ["Ana", "9 000"],
        ["Ben", "14 500"],
        ["Cara", "11 250"],
        ["Result", "=MAX(B2:B4) → 14 500, =INDEX(A2:A4;MATCH(MAX(B2:B4);B2:B4;0)) → Ben"],
      ],
    },
    tips: [
      "MAX ignores text and blank cells; MAXA counts TRUE as 1 and text as 0.",
      "For the second largest value use =LARGE(B2:B100;2).",
    ],
    practice: {
      task: "Return the latest date in D2:D200.",
      answer: "=MAX(D2:D200) - dates are numbers, so MAX works and you just format the result as a date.",
    },
    quiz: [
      {
        q: "How do you return the 3rd highest value?",
        options: ["=MAX(range;3)", "=LARGE(range;3)", "=BIG(range;3)", "=MAX3(range)"],
        answer: 1,
        explain: "LARGE(range;k) returns the k-th largest; SMALL does the opposite.",
      },
    ],
  },
  {
    slug: "min",
    title: "MIN - find the smallest number",
    track: "foundations",
    level: "Beginner",
    summary:
      "MIN returns the lowest value and is the easiest way to build a cap in pricing and payroll formulas.",
    objectives: [
      "Return the lowest value in a range",
      "Cap a value with MIN",
      "Ignore zeros when finding a minimum",
    ],
    syntax: "=MIN(number1; [number2]; …)",
    steps: [
      { title: "Basic minimum", detail: "Point MIN at the range.", formula: "=MIN(B2:B100)" },
      {
        title: "Cap a number",
        detail:
          "=MIN(B2;1000) gives you B2 but never more than 1000 - an instant discount or commission cap.",
        formula: "=MIN(B2;1000)",
      },
      {
        title: "Minimum ignoring zeros",
        detail:
          "Zeros often mean 'no data'. Filter them out first with the modern MINIFS.",
        formula: "=MINIFS(B2:B100;B2:B100;\">0\")",
      },
      {
        title: "Clamp between two limits",
        detail: "Combine MIN and MAX to keep a value inside a band.",
        formula: "=MIN(MAX(B2;0);100)",
      },
    ],
    tips: [
      "MIN on an empty range returns 0, not an error - always check your range is not empty.",
      "=SMALL(range;2) gives the second smallest value.",
    ],
    practice: {
      task: "Commission is 5% of sales but never more than 2 000. Write it.",
      answer: "=MIN(B2*5%;2000)",
    },
    quiz: [
      {
        q: "What does =MIN(MAX(A1;0);100) do?",
        options: [
          "Returns the average",
          "Keeps A1 between 0 and 100",
          "Errors when A1 is negative",
          "Rounds A1",
        ],
        answer: 1,
        explain: "MAX pushes it up to at least 0, MIN pulls it down to at most 100.",
      },
    ],
  },
  {
    slug: "counta",
    title: "COUNTA - count everything that is not empty",
    track: "foundations",
    level: "Beginner",
    summary:
      "COUNTA counts filled cells of any type: text, numbers, dates, even errors. It is how you count records.",
    objectives: [
      "Count non-empty cells",
      "Know the difference between COUNT, COUNTA and COUNTBLANK",
      "Spot the invisible-empty-string trap",
    ],
    syntax: "=COUNTA(value1; [value2]; …)",
    steps: [
      {
        title: "Count the records in a list",
        detail:
          "Point COUNTA at the name column, not the number column, so rows without a value still count.",
        formula: "=COUNTA(A2:A1000)",
      },
      {
        title: "Compare with COUNT",
        detail:
          "COUNT only counts numbers. If A2:A1000 holds names, COUNT returns 0 and COUNTA returns the true number of records.",
      },
      {
        title: "Count the gaps",
        detail: "COUNTBLANK does the opposite and is a great data-quality check.",
        formula: "=COUNTBLANK(A2:A1000)",
      },
      {
        title: "Beware \"\" from formulas",
        detail:
          "A cell containing =IF(A1>0;A1;\"\") looks empty but COUNTA counts it. Use =COUNTIF(range;\"?*\")+COUNT(range) or SUMPRODUCT to exclude it.",
        formula: "=SUMPRODUCT(--(A2:A1000<>\"\"))",
      },
    ],
    example: {
      headers: ["Range content", "COUNT", "COUNTA", "COUNTBLANK"],
      rows: [
        ["10, \"apple\", blank, 5", "2", "3", "1"],
        ["\"\", \"x\", 3", "1", "3", "0"],
      ],
    },
    practice: {
      task: "How many invoices are in the list if invoice numbers are text codes in A2:A500?",
      answer: "=COUNTA(A2:A500)",
    },
    quiz: [
      {
        q: "Which function counts cells that contain text?",
        options: ["COUNT", "COUNTA", "COUNTBLANK", "SUM"],
        answer: 1,
        explain: "COUNT is numbers only; COUNTA counts anything non-empty.",
      },
    ],
  },
  {
    slug: "count",
    title: "COUNT - count numeric values",
    track: "foundations",
    level: "Beginner",
    summary:
      "COUNT tallies only cells containing numbers (and dates), which makes it a quick validity check on imported data.",
    objectives: [
      "Count numeric entries",
      "Use COUNT to detect numbers stored as text",
      "Count rows with COUNTA vs values with COUNT",
    ],
    syntax: "=COUNT(value1; [value2]; …)",
    steps: [
      { title: "Count numbers", detail: "Point COUNT at the value column.", formula: "=COUNT(B2:B500)" },
      {
        title: "Detect text-numbers",
        detail:
          "If =COUNTA(B2:B500) is 499 but =COUNT(B2:B500) is 412, then 87 of your 'numbers' are text. Fix them with Text to Columns > Finish, or =VALUE().",
        formula: "=COUNTA(B2:B500)-COUNT(B2:B500)",
      },
      {
        title: "Count dates",
        detail: "Dates are numbers, so COUNT includes them.",
      },
    ],
    tips: [
      "The status bar at the bottom right already shows Count, Numerical Count, Sum and Average for the selection - right-click it to choose which.",
    ],
    practice: {
      task: "Write a data-quality formula that returns how many cells in B2:B500 are non-numeric but filled.",
      answer: "=COUNTA(B2:B500)-COUNT(B2:B500)",
    },
    quiz: [
      {
        q: "=COUNT(A1:A5) where A1:A5 = 1, \"2 apples\", 3, blank, 01/01/2024 returns:",
        options: ["2", "3", "4", "5"],
        answer: 1,
        explain: "1, 3 and the date are numbers. \"2 apples\" is text and the blank counts for nothing.",
      },
    ],
  },
  {
    slug: "average",
    title: "AVERAGE - the arithmetic mean, done safely",
    track: "foundations",
    level: "Beginner",
    summary:
      "AVERAGE adds and divides for you, but how it treats blanks and zeros decides whether your KPI is right or wrong.",
    objectives: [
      "Calculate a mean",
      "Understand blanks vs zeros",
      "Average only rows that meet a condition",
    ],
    syntax: "=AVERAGE(number1; [number2]; …)",
    steps: [
      { title: "Basic average", detail: "Point at the range.", formula: "=AVERAGE(B2:B100)" },
      {
        title: "Blanks are skipped, zeros are not",
        detail:
          "A blank cell is excluded from the divisor; a 0 is included and drags the mean down. Decide which behaviour you want before you publish the number.",
      },
      {
        title: "Ignore zeros deliberately",
        detail: "Use AVERAGEIF with a criterion.",
        formula: "=AVERAGEIF(B2:B100;\"<>0\")",
      },
      {
        title: "Guard against an empty range",
        detail: "AVERAGE of nothing is #DIV/0!. Wrap it.",
        formula: "=IFERROR(AVERAGE(B2:B100);0)",
      },
      {
        title: "Weighted average",
        detail: "When each row has a weight, AVERAGE is wrong - use SUMPRODUCT.",
        formula: "=SUMPRODUCT(B2:B100;C2:C100)/SUM(C2:C100)",
      },
    ],
    practice: {
      task: "Average the scores in B2:B50 but exclude zeros.",
      answer: "=AVERAGEIF(B2:B50;\"<>0\")",
    },
    quiz: [
      {
        q: "Values are 10, 0, blank, 20. What does AVERAGE return?",
        options: ["10", "7.5", "15", "30"],
        answer: 0,
        explain: "It averages 10, 0 and 20 = 30/3 = 10. The blank is ignored, the zero is not.",
      },
    ],
  },
  {
    slug: "addition",
    title: "Addition in Excel - tutorial and practice",
    track: "foundations",
    level: "Beginner",
    summary:
      "Four ways to add: the + operator, SUM, AutoSum and Paste Special > Add for bulk changes.",
    objectives: [
      "Add two cells and a whole column",
      "Add a constant to hundreds of cells without a helper column",
    ],
    steps: [
      { title: "The + operator", detail: "For two or three cells.", formula: "=A2+B2" },
      { title: "SUM for ranges", detail: "For four or more.", formula: "=SUM(A2:A20)" },
      { title: "AutoSum", detail: "Select the cell below and press Alt + =." },
      {
        title: "Add a constant in place",
        detail:
          "Type the number in a spare cell, copy it, select your data, then Ctrl + Alt + V and choose Values + Add. Every selected number increases - no formulas left behind.",
      },
      { title: "Add across sheets", detail: "3-D reference adds the same cell on every sheet from Jan to Dec.", formula: "=SUM(Jan:Dec!B5)" },
    ],
    practice: {
      task: "Increase every price in C2:C300 by 15 without adding a column.",
      answer: "Type 15 anywhere, Ctrl + C, select C2:C300, Ctrl + Alt + V → Values + Add → OK.",
    },
    quiz: [
      {
        q: "Which Paste Special option adds a copied number to the selection?",
        options: ["Values", "Add", "Multiply", "Transpose"],
        answer: 1,
        explain: "Paste Special > Operation > Add.",
      },
    ],
  },
  {
    slug: "subtraction",
    title: "Subtraction (minus) in Excel - tutorial and practice",
    track: "foundations",
    level: "Beginner",
    summary:
      "There is no SUBTRACT function. You use the minus operator, or SUM with a negative range.",
    objectives: [
      "Subtract cells, columns and totals",
      "Subtract dates to get days between",
    ],
    steps: [
      { title: "Basic minus", detail: "Income minus cost.", formula: "=B2-C2" },
      {
        title: "Subtract many cells from one",
        detail: "Wrap the deductions in SUM.",
        formula: "=B2-SUM(C2:F2)",
      },
      {
        title: "Days between two dates",
        detail: "Dates are serial numbers, so a plain minus works.",
        formula: "=B2-A2",
      },
      {
        title: "Show negatives in red brackets",
        detail: "Ctrl + 1 → Custom → #,##0;[Red](#,##0)",
      },
    ],
    practice: {
      task: "Budget is in B2, actual in C2. Show the variance and make overspend red.",
      answer: "=B2-C2 then Ctrl + 1 → Custom → #,##0;[Red](#,##0)",
    },
    quiz: [
      {
        q: "Which is the correct way to subtract C2:F2 from B2?",
        options: ["=B2-C2:F2", "=B2-SUM(C2:F2)", "=SUBTRACT(B2;C2:F2)", "=MINUS(B2;C2:F2)"],
        answer: 1,
        explain: "Excel has no SUBTRACT function; SUM the deductions and subtract the result.",
      },
    ],
  },
  {
    slug: "multiply",
    title: "Multiply numbers in Excel - tutorial and practice",
    track: "foundations",
    level: "Beginner",
    summary:
      "Use * for two cells, PRODUCT for many and SUMPRODUCT to multiply and total two columns at once.",
    objectives: [
      "Multiply cells and columns",
      "Apply a fixed rate with an absolute reference",
      "Total qty × price in a single cell",
    ],
    steps: [
      { title: "Multiply two cells", detail: "Quantity times price.", formula: "=B2*C2" },
      { title: "PRODUCT for a range", detail: "Multiplies every number in the range.", formula: "=PRODUCT(B2:D2)" },
      {
        title: "Multiply by a fixed rate",
        detail: "Lock the rate cell with F4 so it does not drift when you copy down.",
        formula: "=B2*$E$1",
      },
      {
        title: "Multiply and total in one shot",
        detail: "SUMPRODUCT multiplies row by row and adds the results - no helper column needed.",
        formula: "=SUMPRODUCT(B2:B50;C2:C50)",
      },
    ],
    practice: {
      task: "Qty in B, unit price in C, rows 2 to 50. Get the order value without a helper column.",
      answer: "=SUMPRODUCT(B2:B50;C2:C50)",
    },
    quiz: [
      {
        q: "Why use $E$1 instead of E1 for a VAT rate?",
        options: [
          "It is faster",
          "So the reference does not shift when copied down",
          "It formats as currency",
          "It rounds the result",
        ],
        answer: 1,
        explain: "The $ signs make the reference absolute.",
      },
    ],
  },
  {
    slug: "division",
    title: "Division in Excel - tutorial and practice",
    track: "foundations",
    level: "Beginner",
    summary:
      "Divide with /, handle #DIV/0! gracefully, and use QUOTIENT and MOD when you need whole parts and remainders.",
    objectives: [
      "Divide cells and calculate percentages",
      "Prevent #DIV/0!",
      "Split a number into whole units and remainder",
    ],
    steps: [
      { title: "Basic division", detail: "Total divided by count.", formula: "=B2/C2" },
      {
        title: "Percentage of total",
        detail: "Lock the total so it stays put as you copy down, then format as %.",
        formula: "=B2/$B$20",
      },
      {
        title: "Stop the #DIV/0! error",
        detail: "Wrap the division so an empty denominator shows a dash instead.",
        formula: "=IFERROR(B2/C2;\"-\")",
      },
      {
        title: "Whole boxes and leftovers",
        detail: "QUOTIENT gives the whole number of boxes, MOD gives the remainder.",
        formula: "=QUOTIENT(B2;12) & \" boxes + \" & MOD(B2;12) & \" loose\"",
      },
    ],
    practice: {
      task: "Show each row's share of the grand total in B20, with no error on blank rows.",
      answer: "=IFERROR(B2/$B$20;\"\") formatted as Percentage",
    },
    quiz: [
      {
        q: "What causes #DIV/0!?",
        options: [
          "Dividing by a blank or zero cell",
          "Dividing by text",
          "Using / instead of ÷",
          "A missing $ sign",
        ],
        answer: 0,
        explain: "A blank cell counts as zero in a division.",
      },
    ],
  },
];
