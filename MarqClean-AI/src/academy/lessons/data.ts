import type { Lesson } from "./types";

export const dataTools: Lesson[] = [
  {
    slug: "filter-tool",
    title: "The Filter tool - see only the rows you need",
    track: "data",
    level: "Beginner",
    summary:
      "AutoFilter hides rows that do not match your criteria without deleting anything, and it is one keyboard shortcut away.",
    objectives: [
      "Turn filters on and off",
      "Filter by value, text, number, date and colour",
      "Total only what is visible",
    ],
    steps: [
      {
        title: "Turn it on",
        detail:
          "Click any cell in the data and press Ctrl + Shift + L (or Data → Filter). Drop-down arrows appear in the header row.",
      },
      {
        title: "Filter by value",
        detail:
          "Click an arrow, untick 'Select All', tick what you want. The search box lets you type instead of scrolling.",
      },
      {
        title: "Text / Number / Date filters",
        detail:
          "The submenu changes with the data type: 'contains', 'begins with', 'greater than', 'between', 'this month', 'top 10'.",
      },
      {
        title: "Filter by colour",
        detail: "If you use fills or conditional formatting, you can filter or sort by that colour directly.",
      },
      {
        title: "Total only visible rows",
        detail:
          "SUM still adds hidden rows. SUBTOTAL respects the filter. In a Table, the Total Row does this automatically.",
        formula: "=SUBTOTAL(109;B2:B500)",
      },
      {
        title: "Clear and reapply",
        detail: "Alt + D, F, S clears all filters. Ctrl + Alt + L reapplies after changing the data.",
      },
    ],
    tips: [
      "Filters need a single header row and no fully blank rows inside the data.",
      "Copy visible rows only: select, then Alt + ; (Go To Special → Visible cells only) before Ctrl + C.",
    ],
    shortcuts: [
      ["Ctrl + Shift + L", "Toggle filters"],
      ["Alt + ↓", "Open the filter drop-down of the active column"],
      ["Alt + ;", "Select visible cells only"],
    ],
    practice: {
      task: "Show only orders from the North over 5 000, then total them.",
      answer:
        "Ctrl + Shift + L → Region = North → Amount → Number Filters → Greater Than → 5000 → then =SUBTOTAL(109;Amount range).",
    },
    quiz: [
      {
        q: "Which function totals only filtered-in rows?",
        options: ["SUM", "SUBTOTAL(109;…)", "SUMIF", "COUNTA"],
        answer: 1,
        explain: "109 = SUM ignoring hidden rows.",
      },
    ],
  },
  {
    slug: "sort-data",
    title: "Sort Data - single, multi-level and custom sorting",
    track: "data",
    level: "Beginner",
    summary:
      "Sorting looks trivial until you sort one column on its own and scramble your table. Here is how to do it safely.",
    objectives: [
      "Sort on one and several columns",
      "Create a custom order (Mon-Sun, Small-Large)",
      "Sort left to right",
    ],
    steps: [
      {
        title: "Select one cell, not one column",
        detail:
          "Click a single cell inside the data and Excel selects the whole table for you. Selecting just a column and sorting is what destroys data.",
      },
      { title: "Quick sort", detail: "Data → A→Z or Z→A, or Alt + A, S, A." },
      {
        title: "Multi-level sort",
        detail:
          "Data → Sort → Add Level. Sort by Region A→Z, then by Amount Largest to Smallest. Levels are applied top-down.",
      },
      {
        title: "Custom list order",
        detail:
          "In the Sort dialog choose Order → Custom List to sort by Mon, Tue, Wed… or by your own priority list (High, Medium, Low).",
      },
      {
        title: "Sort left to right",
        detail: "Sort → Options → Sort left to right, when your months run across the columns.",
      },
      {
        title: "Formula alternative",
        detail: "SORT() returns a sorted copy that updates live and never touches the source.",
        formula: "=SORT(A2:D500;3;-1)",
      },
    ],
    mistakes: [
      "Sorting with merged cells - Excel refuses or mangles the data.",
      "Forgetting 'My data has headers' and sorting your header row into the middle.",
    ],
    practice: {
      task: "Sort by Department A→Z and then by Salary high→low.",
      answer: "Data → Sort → Level 1 Department A→Z → Add Level → Salary Largest to Smallest.",
    },
    quiz: [
      {
        q: "How do you sort by a custom order like High/Medium/Low?",
        options: [
          "Alphabetically",
          "Sort dialog → Order → Custom List",
          "You cannot",
          "Rename the values",
        ],
        answer: 1,
        explain: "Custom Lists let you define any sequence.",
      },
    ],
  },
  {
    slug: "flash-fill",
    title: "Flash Fill - Excel learns the pattern you type",
    track: "data",
    level: "Beginner",
    summary:
      "Type the result you want once, press Ctrl + E, and Excel fills the rest by example. No formula required.",
    objectives: ["Split and join text by example", "Reformat codes and phone numbers", "Know the limits"],
    steps: [
      {
        title: "Put your source data next to an empty column",
        detail: "Flash Fill only works when the example column is adjacent to the data.",
      },
      {
        title: "Type the first result manually",
        detail: "For 'Audun Danielsen' type 'A. Danielsen'.",
      },
      {
        title: "Press Ctrl + E",
        detail: "Excel proposes the whole column. Press Enter to accept, or keep typing to refine it.",
      },
      {
        title: "Give a second example when it guesses wrong",
        detail: "Type row two as well, then Ctrl + E again - accuracy jumps dramatically.",
      },
      {
        title: "Great use cases",
        detail:
          "Splitting names, extracting the year from ORD-2024-0091, formatting phone numbers, joining initials, changing capitalisation.",
      },
      {
        title: "The catch",
        detail:
          "Flash Fill produces static text. If the source changes, the result does not. Use TEXTSPLIT/TEXTBEFORE for a live version.",
      },
    ],
    shortcuts: [["Ctrl + E", "Flash Fill"]],
    practice: {
      task: "Column A has '+47 900 12 345'. Get '90012345'.",
      answer: "Type 90012345 in B2, press Ctrl + E.",
    },
    quiz: [
      {
        q: "Flash Fill results are:",
        options: ["Live formulas", "Static text", "Pivot fields", "Named ranges"],
        answer: 1,
        explain: "They do not update when the source changes.",
      },
    ],
  },
  {
    slug: "remove-duplicates",
    title: "Remove Duplicates - safely de-duplicate a list",
    track: "data",
    level: "Beginner",
    summary:
      "The Remove Duplicates button deletes rows permanently, so learn the safe workflow - and the formula alternatives.",
    objectives: [
      "Remove duplicates on one or several columns",
      "Find duplicates before deleting them",
      "Use UNIQUE for a non-destructive list",
    ],
    steps: [
      {
        title: "Copy the sheet first",
        detail: "Remove Duplicates is destructive. Right-click the tab → Move or Copy → Create a copy.",
      },
      {
        title: "Find them first",
        detail:
          "Home → Conditional Formatting → Highlight Cells Rules → Duplicate Values, or =COUNTIF($A$2:$A$500;A2)>1.",
      },
      {
        title: "Run the tool",
        detail:
          "Select a cell in the data → Data → Remove Duplicates → tick the columns that define a duplicate. Ticking every column means only fully identical rows are removed.",
      },
      {
        title: "Read the report",
        detail: "Excel tells you how many were removed and how many unique values remain. Note it down.",
      },
      {
        title: "Non-destructive alternative",
        detail: "UNIQUE spills a live list next to the data and leaves the original untouched.",
        formula: "=UNIQUE(A2:A500)",
      },
      {
        title: "Watch out",
        detail:
          "Trailing spaces and different capitalisation are treated as duplicates by the tool (case-insensitive) but not by exact-match formulas. TRIM first.",
      },
    ],
    practice: {
      task: "Get a live list of unique customer names from A2:A500 without deleting anything.",
      answer: "=SORT(UNIQUE(FILTER(A2:A500;A2:A500<>\"\")))",
    },
    quiz: [
      {
        q: "Remove Duplicates is:",
        options: ["Non-destructive", "Destructive - it deletes rows", "A formula", "A pivot feature"],
        answer: 1,
        explain: "Always work on a copy.",
      },
    ],
  },
  {
    slug: "conditional-formatting",
    title: "Conditional Formatting - make data visually scannable",
    track: "data",
    level: "Intermediate",
    summary:
      "Rules that colour cells automatically: value rules, data bars, and formula rules that highlight the whole row.",
    objectives: [
      "Apply built-in rules",
      "Write a formula rule with the right $ locking",
      "Manage and clean up rules",
    ],
    steps: [
      {
        title: "Built-in rules",
        detail:
          "Select the range → Home → Conditional Formatting → Highlight Cells Rules (greater than, between, text contains, duplicate values) or Top/Bottom Rules (top 10, above average).",
      },
      {
        title: "Data bars, colour scales, icon sets",
        detail:
          "Great for dashboards. In More Rules you can tick 'Show Bar Only' to hide the numbers and keep only the bar.",
      },
      {
        title: "Formula rules - highlight an entire row",
        detail:
          "Select A2:H500 (start at the first data row!) → New Rule → Use a formula → =$D2=\"Overdue\". The $ locks the column, the free row number lets the rule travel down.",
        formula: "=$D2=\"Overdue\"",
      },
      { title: "Highlight duplicates", detail: "", formula: "=COUNTIF($A$2:$A$500;$A2)>1" },
      { title: "Highlight weekends", detail: "", formula: "=WEEKDAY($A2;2)>5" },
      { title: "Highlight blanks", detail: "", formula: "=ISBLANK($B2)" },
      { title: "Compare to a threshold cell", detail: "", formula: "=$C2>$G$1" },
      {
        title: "Manage the rules",
        detail:
          "Home → Conditional Formatting → Manage Rules → 'This Worksheet' shows every rule, its range and its priority. Use 'Stop If True' to prevent rules from stacking.",
      },
    ],
    tips: [
      "Two colours plus grey is enough. If everything is highlighted, nothing is.",
      "Copying and pasting rows fragments rules into dozens of ranges - periodically delete the rules and reapply once.",
    ],
    practice: {
      task: "Shade the whole row light red when the due date in column E is before today and column F is not 'Paid'.",
      answer: "Select A2:H500 → New Rule → formula: =AND($E2<TODAY();$F2<>\"Paid\")",
    },
    quiz: [
      {
        q: "For a whole-row highlight, the reference in the rule should be:",
        options: ["$D$2", "$D2", "D$2", "D2"],
        answer: 1,
        explain: "Lock the column only, so the rule checks column D on each row.",
      },
    ],
  },
  {
    slug: "pivot-table",
    title: "Pivot Table - summarise thousands of rows in seconds",
    track: "data",
    level: "Intermediate",
    summary:
      "A PivotTable groups, totals and cross-tabs your data by dragging fields into four zones. No formulas involved.",
    objectives: [
      "Prepare data correctly",
      "Use the Rows / Columns / Values / Filters zones",
      "Change the calculation and show percentages",
      "Group dates and add slicers",
    ],
    steps: [
      {
        title: "Step 0 - make the data boring",
        detail:
          "One header row, no merged cells, no blank rows, no subtotals inside the data. Press Ctrl + T so the pivot source grows automatically.",
      },
      { title: "Insert", detail: "Click inside the table → Alt + N + V → Enter. A blank pivot and the Field List appear." },
      {
        title: "The four zones",
        detail:
          "Rows = what you list down the page (Customer). Columns = what you compare across (Year). Values = what you count or sum (Amount). Filters = what you slice off the top (Region).",
      },
      {
        title: "Change the calculation",
        detail:
          "Right-click a value → Summarize Values By → Sum / Count / Average / Max. Excel defaults to Count when the column contains any text - a classic 'why is my total wrong' moment.",
      },
      {
        title: "Show percentages instead of numbers",
        detail:
          "Right-click a value → Show Values As → % of Grand Total / % of Column Total / Running Total In. The most underused pivot feature there is.",
      },
      {
        title: "Group dates",
        detail:
          "Right-click any date in the pivot → Group → tick Months and Years. Instant monthly report.",
      },
      {
        title: "Make it readable",
        detail:
          "Design → Report Layout → Show in Tabular Form, then Repeat All Item Labels. Right-click a value → Number Format → #,##0.",
      },
      {
        title: "Slicers and refresh",
        detail:
          "PivotTable Analyze → Insert Slicer for clickable filters. Data changed? Press Alt + F5 to refresh (Ctrl + Alt + F5 refreshes all).",
      },
      {
        title: "Drill down",
        detail:
          "Double-click any number in the pivot and Excel creates a new sheet listing the exact rows behind it. Perfect for audits.",
      },
    ],
    shortcuts: [
      ["Alt + N + V", "Insert PivotTable"],
      ["Alt + F5", "Refresh the pivot"],
      ["Ctrl + T", "Make the source a Table first"],
    ],
    practice: {
      task: "From a sales table, build monthly revenue per region, showing each region's share of the total.",
      answer:
        "Rows = Region, Columns = Date grouped by Month, Values = Amount (Sum), then Show Values As → % of Column Total.",
    },
    quiz: [
      {
        q: "Your pivot shows Count instead of Sum. Why?",
        options: [
          "The pivot is broken",
          "The value column contains text or blanks",
          "You need a slicer",
          "The data is not sorted",
        ],
        answer: 1,
        explain: "Excel defaults to Count when it detects non-numeric entries in the field.",
      },
      {
        q: "What does double-clicking a pivot value do?",
        options: [
          "Nothing",
          "Creates a sheet with the underlying rows",
          "Deletes the value",
          "Refreshes the pivot",
        ],
        answer: 1,
        explain: "That is the drill-down feature.",
      },
    ],
  },
  {
    slug: "text-to-columns",
    title: "Text to Columns - split a column in three clicks",
    track: "data",
    level: "Beginner",
    summary:
      "The classic wizard that splits delimited or fixed-width data - and secretly the fastest fix for text-numbers and foreign date formats.",
    objectives: [
      "Split by delimiter and by width",
      "Skip unwanted columns",
      "Fix DMY dates and text-numbers",
    ],
    steps: [
      {
        title: "Make room",
        detail:
          "Text to Columns overwrites the columns to the right. Insert enough blank columns first.",
      },
      { title: "Start the wizard", detail: "Select the column → Data → Text to Columns (Alt + A, E)." },
      {
        title: "Delimited or Fixed width",
        detail:
          "Delimited for comma/semicolon/tab/space data. Fixed width for old mainframe reports - you click to place the break lines.",
      },
      {
        title: "Choose the delimiter",
        detail:
          "Tick as many as you need, use 'Other' for | or /, and tick 'Treat consecutive delimiters as one' for space-separated data.",
      },
      {
        title: "Step 3 is the powerful one",
        detail:
          "Click a column in the preview and set it to Text (to keep leading zeros), to Date: DMY (to fix European dates importing as text), or to 'Do not import' to discard it entirely.",
      },
      {
        title: "The hidden trick",
        detail:
          "Select a column of text-numbers, run Text to Columns and just press Finish. Every value is re-parsed as a real number.",
      },
    ],
    shortcuts: [["Alt + A, E", "Text to Columns"]],
    practice: {
      task: "Split 'Smith, John' in column A into surname and first name.",
      answer: "Insert a blank column B → Data → Text to Columns → Delimited → Comma → Finish → TRIM the results.",
    },
    quiz: [
      {
        q: "How do you keep leading zeros while splitting?",
        options: [
          "Impossible",
          "Set that column's format to Text in step 3",
          "Use General",
          "Format afterwards",
        ],
        answer: 1,
        explain: "General strips leading zeros; Text preserves them.",
      },
    ],
  },
  {
    slug: "excel-wildcards",
    title: "Excel Wildcards - * ? and ~",
    track: "data",
    level: "Intermediate",
    summary:
      "Three characters that turn exact matching into pattern matching in COUNTIF, SUMIF, SEARCH, Find & Replace and filters.",
    objectives: ["Use * and ?", "Escape a literal * or ?", "Know where wildcards do NOT work"],
    steps: [
      { title: "* means any number of characters", detail: "Counts every cell that contains 'ltd'.", formula: "=COUNTIF(A2:A500;\"*ltd*\")" },
      { title: "? means exactly one character", detail: "Matches AB-1, AB-7 but not AB-12.", formula: "=COUNTIF(A2:A500;\"AB-?\")" },
      { title: "Begins with / ends with", detail: "", formula: "=SUMIF(A2:A500;\"INV*\";B2:B500)" },
      {
        title: "~ escapes a real asterisk",
        detail: "To find a literal question mark, search for ~?",
        formula: "=COUNTIF(A2:A500;\"~*\")",
      },
      {
        title: "Where they work",
        detail:
          "COUNTIF(S), SUMIF(S), AVERAGEIF(S), MATCH (with match_type 0), SEARCH, VLOOKUP/HLOOKUP, XLOOKUP with match_mode 2, Find & Replace, AutoFilter.",
      },
      {
        title: "Where they do NOT work",
        detail:
          "Plain comparisons (=A2=\"*ltd*\" is FALSE), FIND, SUBSTITUTE and most maths functions. Use SEARCH there instead.",
      },
    ],
    practice: {
      task: "Sum every invoice whose reference starts with 'UK-' and ends with '-24'.",
      answer: "=SUMIF(A2:A500;\"UK-*-24\";B2:B500)",
    },
    quiz: [
      {
        q: "Which wildcard matches exactly one character?",
        options: ["*", "?", "~", "#"],
        answer: 1,
        explain: "? = one character, * = any number of characters.",
      },
    ],
  },
  {
    slug: "advanced-filter",
    title: "Advanced Filter - criteria ranges and extracting to a new place",
    track: "data",
    level: "Advanced",
    summary:
      "Advanced Filter handles OR logic, formula criteria and can copy the matching rows somewhere else - things AutoFilter cannot do.",
    objectives: [
      "Build a criteria range",
      "Combine AND (same row) and OR (different rows)",
      "Extract unique records to another location",
    ],
    steps: [
      {
        title: "Build the criteria range",
        detail:
          "Above your data, copy the exact header names and put the conditions underneath. Same row = AND. Different rows = OR.",
      },
      {
        title: "Example criteria",
        detail:
          "Header Region / Amount. Row 1: North | >5000. Row 2: South | (blank). Meaning: North AND over 5000, OR anything from the South.",
      },
      {
        title: "Run it",
        detail:
          "Data → Advanced → choose 'Filter the list, in-place' or 'Copy to another location' → set List range, Criteria range, and Copy to.",
      },
      {
        title: "Extract unique records",
        detail: "Tick 'Unique records only' to get a de-duplicated extract without touching the source.",
      },
      {
        title: "Formula criteria",
        detail:
          "Leave the criteria header BLANK (or use a name that is not a real header) and write a formula referring to the first data row, e.g. =D2>AVERAGE($D$2:$D$500). That gives you above-average filtering.",
      },
      {
        title: "Modern alternative",
        detail: "In Microsoft 365 the FILTER function does most of this live, with no dialog.",
        formula: "=FILTER(A2:D500;((C2:C500=\"North\")*(B2:B500>5000))+(C2:C500=\"South\"))",
      },
    ],
    practice: {
      task: "Copy every order over 10 000 from the North or South into a new sheet, unique only.",
      answer:
        "Criteria range with Region/Amount → North|>10000 on one row, South|>10000 on the next → Data → Advanced → Copy to another location → Unique records only.",
    },
    quiz: [
      {
        q: "In a criteria range, two conditions on different rows mean:",
        options: ["AND", "OR", "NOT", "Invalid"],
        answer: 1,
        explain: "Same row = AND, different rows = OR.",
      },
    ],
  },
  {
    slug: "power-query-combine-workbooks",
    title: "Power Query - combine data from multiple workbooks",
    track: "data",
    level: "Advanced",
    summary:
      "Point Power Query at a folder and it appends every file inside into one table that refreshes with a single click each month.",
    objectives: [
      "Import from a folder",
      "Use the sample-file transform",
      "Refresh next month in one click",
    ],
    steps: [
      {
        title: "Put the files in one folder",
        detail:
          "Every workbook must have the same column layout and the same sheet or table name. Nothing else matters.",
      },
      { title: "Data → Get Data → From File → From Folder", detail: "Browse to the folder and press Open, then Transform Data." },
      {
        title: "Filter the file list",
        detail:
          "You now see one row per file. Filter Extension to .xlsx and remove any ~$ temporary files. Keep Name and Date modified if you want them in the output.",
      },
      {
        title: "Click Combine Files",
        detail:
          "The double-arrow icon on the Content column. Choose the sheet or table to take from each workbook. Power Query builds a sample query and a helper function automatically.",
      },
      {
        title: "Clean once, applies to all",
        detail:
          "Promote headers, set data types, remove junk columns, rename. Every step is recorded and will run on every file forever.",
      },
      {
        title: "Add the source file name",
        detail:
          "The Source.Name column tells you which workbook each row came from - keep it for auditing.",
      },
      { title: "Close & Load To…", detail: "Load as a Table on a new sheet, or straight into the Data Model for a pivot." },
      {
        title: "Next month",
        detail:
          "Drop the new file into the folder and press Data → Refresh All (Ctrl + Alt + F5). That is the whole monthly process.",
      },
    ],
    tips: [
      "If a file has an extra column, Power Query adds it with nulls instead of failing - check for unexpected nulls.",
      "Use a folder path stored in a cell (a parameter) so the query works on a colleague's machine too.",
    ],
    practice: {
      task: "12 monthly sales files in one folder need to become one table.",
      answer:
        "Get Data → From Folder → Transform → filter extensions → Combine Files → pick the sheet → set types → Close & Load. Refresh monthly.",
    },
    quiz: [
      {
        q: "What must be true of the files you combine?",
        options: [
          "They must be CSV",
          "They need the same structure/sheet name",
          "They must be under 1 MB",
          "They must be open",
        ],
        answer: 1,
        explain: "Consistent structure is what lets one transform apply to all of them.",
      },
    ],
  },
  {
    slug: "power-query-column-from-examples",
    title: "Power Query - Column From Examples",
    track: "data",
    level: "Intermediate",
    summary:
      "Flash Fill inside Power Query: type the output you want and it writes the M code for you, permanently and refreshably.",
    objectives: ["Create derived columns by example", "Read and tweak the generated M code"],
    steps: [
      { title: "Open the query", detail: "Data → Get Data → From Table/Range, or double-click an existing query." },
      {
        title: "Add Column → Column From Examples → From Selection",
        detail: "Select the source column(s) first so the suggestions are focused.",
      },
      {
        title: "Type one or two examples",
        detail:
          "Enter the result for row 1. Power Query fills the rest and shows the transformation it inferred at the top.",
      },
      {
        title: "Refine",
        detail:
          "Correct a wrong row and it re-learns instantly. Press OK when the preview is right.",
      },
      {
        title: "Look at the M code",
        detail:
          "The generated step is real M, e.g. Text.BetweenDelimiters([Code];\"-\";\"-\"). Edit it in the formula bar to make it more robust.",
      },
      {
        title: "Why it beats Flash Fill",
        detail:
          "It is part of the query, so it re-runs automatically on next month's data instead of being static text.",
      },
    ],
    practice: {
      task: "From 'ORD-2024-0091' create a Year column.",
      answer:
        "Add Column → Column From Examples → type 2024 → OK. Generated step: Text.BetweenDelimiters([Ref];\"-\";\"-\").",
    },
    quiz: [
      {
        q: "The main advantage over Flash Fill is:",
        options: ["It is faster", "It refreshes with the data", "It uses less memory", "It supports colour"],
        answer: 1,
        explain: "Power Query steps re-run every refresh.",
      },
    ],
  },
  {
    slug: "unpivot-columns-power-query",
    title: "Unpivot columns using Power Query",
    track: "data",
    level: "Intermediate",
    summary:
      "Turn a wide, human-friendly report (months across the top) into the tall, tidy format that pivots and charts actually need.",
    objectives: [
      "Recognise wide vs tidy data",
      "Unpivot other columns",
      "Keep the query robust when new months arrive",
    ],
    steps: [
      {
        title: "Understand the goal",
        detail:
          "Wide: Region | Jan | Feb | Mar. Tidy: Region | Month | Amount. Every pivot, chart and SUMIFS wants the tidy shape.",
      },
      { title: "Load the table", detail: "Select the data → Data → From Table/Range." },
      {
        title: "Select the columns to KEEP",
        detail:
          "Click the Region (and any other identifier) column, right-click → Unpivot Other Columns. Choosing 'other columns' means new months added later are handled automatically.",
      },
      {
        title: "Rename the results",
        detail: "'Attribute' → Month, 'Value' → Amount. Set Amount to a decimal number type.",
      },
      {
        title: "Clean up",
        detail:
          "Remove Total rows and columns before unpivoting, otherwise your totals become data rows and everything double-counts.",
      },
      { title: "Close & Load", detail: "Load to a table or the Data Model and build your pivot on the tidy output." },
      {
        title: "Reverse it",
        detail: "Need wide again for a printed report? Select Month → Transform → Pivot Column → Values = Amount.",
      },
    ],
    practice: {
      task: "A budget sheet has Department in A and 12 month columns. Make it pivot-ready.",
      answer:
        "From Table/Range → select Department → Unpivot Other Columns → rename Attribute/Value → Close & Load.",
    },
    quiz: [
      {
        q: "Why choose 'Unpivot Other Columns' rather than 'Unpivot Columns'?",
        options: [
          "It is quicker to click",
          "New columns added later are included automatically",
          "It keeps formatting",
          "No difference",
        ],
        answer: 1,
        explain: "The step is defined by the columns you keep, so extra months need no maintenance.",
      },
    ],
  },
];
