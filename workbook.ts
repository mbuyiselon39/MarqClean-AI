import type { Lesson } from "./types";

export const workbook: Lesson[] = [
  {
    slug: "excel-interface-tour",
    title: "The Excel window: ribbon, tabs, name box and status bar",
    track: "workbook",
    level: "Beginner",
    summary:
      "Before any formula, learn to name every part of the screen so instructions like 'Home > Alignment > Wrap Text' make instant sense.",
    objectives: [
      "Name the ribbon, tabs, groups, dialog launchers and the Quick Access Toolbar",
      "Use the Name Box and the Formula Bar",
      "Read the status bar totals and change the view and zoom",
    ],
    steps: [
      {
        title: "The ribbon and its tabs",
        detail:
          "Along the top: File (backstage), Home, Insert, Page Layout, Formulas, Data, Review, View, and Developer if you switch it on. Each tab holds groups, and a small arrow at the bottom-right of a group (the dialog launcher) opens the full options window.",
      },
      {
        title: "Collapse it when you need space",
        detail: "Ctrl + F1 hides and shows the ribbon. Double-click any tab does the same.",
      },
      {
        title: "The Quick Access Toolbar",
        detail:
          "The tiny icons above or below the ribbon. Right-click any ribbon button and choose 'Add to Quick Access Toolbar'. Then Alt + 1, Alt + 2 and so on run them from the keyboard.",
      },
      {
        title: "Name Box and Formula Bar",
        detail:
          "Top-left shows the active cell address; type a reference such as C500 there and press Enter to jump straight to it. The Formula Bar to the right shows what the cell really contains. Ctrl + Shift + U expands it for long formulas.",
      },
      {
        title: "The sheet tabs",
        detail:
          "Right-click a tab to insert, rename, colour, move, copy, hide or protect a sheet. Ctrl + Page Down and Ctrl + Page Up move between sheets.",
      },
      {
        title: "The status bar is a free calculator",
        detail:
          "Select any numbers and the bottom-right shows Average, Count and Sum. Right-click it to add Min, Max and Numerical Count.",
      },
      {
        title: "Views and zoom",
        detail:
          "Bottom-right buttons switch between Normal, Page Layout and Page Break Preview. Ctrl + mouse wheel zooms. View > Freeze Panes keeps headings on screen.",
      },
      {
        title: "Backstage",
        detail:
          "File opens the backstage view: New, Open, Info (protect workbook, inspect document), Save As, Print with live preview, Share, Export to PDF and Options.",
      },
    ],
    shortcuts: [
      ["Ctrl + F1", "Collapse or show the ribbon"],
      ["Alt", "Show the KeyTips for every ribbon command"],
      ["Ctrl + G / F5", "Go To a cell or range"],
      ["Ctrl + Page Up / Down", "Previous or next sheet"],
    ],
    practice: {
      task: "Jump to cell AA500, name the range A1:D10 as SalesData, then return to A1 in one keystroke.",
      answer:
        "Type AA500 in the Name Box and press Enter. Select A1:D10, type SalesData in the Name Box, press Enter. Press Ctrl + Home to return to A1.",
    },
    quiz: [
      {
        q: "What does the Name Box do?",
        options: [
          "Renames the workbook",
          "Shows the active cell and lets you jump to or name a range",
          "Renames a sheet",
          "Shows the file path",
        ],
        answer: 1,
        explain: "It is both an address indicator and a navigation and naming tool.",
      },
    ],
  },
  {
    slug: "entering-editing-data",
    title: "Entering and editing data accurately",
    track: "workbook",
    level: "Beginner",
    summary:
      "How Excel decides whether your entry is a number, a date or text, and the fill techniques that save hours of typing.",
    objectives: [
      "Enter numbers, text, dates and times so Excel recognises them",
      "Use AutoFill and custom lists",
      "Edit, undo, and fill a whole selection in one keystroke",
    ],
    steps: [
      {
        title: "Watch the alignment",
        detail:
          "Numbers and dates align right, text aligns left. If a number sits on the left it is text and will break your sums.",
      },
      {
        title: "Enter special values correctly",
        detail:
          "Leading zeros need a Text format or a leading apostrophe. Fractions need a zero first (0 1/2). Percentages can be typed with the % sign. Ctrl + ; enters today's date as a fixed value.",
      },
      {
        title: "AutoFill",
        detail:
          "Drag the small square at the bottom-right of the selection. One number copies; two selected numbers continue the pattern. Dates, months and weekdays extend automatically.",
      },
      {
        title: "AutoFill options",
        detail:
          "After dragging, the small icon that appears lets you choose Copy Cells, Fill Series, Fill Formatting Only or Fill Without Formatting.",
      },
      {
        title: "Flash Fill",
        detail: "Type one example of the result you want and press Ctrl + E.",
      },
      {
        title: "Fill a whole selection at once",
        detail:
          "Select the range, type the value or formula, then press Ctrl + Enter instead of Enter.",
      },
      {
        title: "Custom lists",
        detail:
          "File > Options > Advanced > Edit Custom Lists lets you teach Excel your own sequence (departments, regions, sizes). It then works with AutoFill and with custom sorting.",
      },
      {
        title: "Editing safely",
        detail:
          "F2 edits in the cell, Esc abandons the edit, Ctrl + Z undoes, Ctrl + Y redoes. Delete clears contents only; Home > Clear > Clear All also removes formatting and comments.",
      },
    ],
    shortcuts: [
      ["Ctrl + Enter", "Fill every selected cell"],
      ["Ctrl + D / Ctrl + R", "Fill down / fill right"],
      ["Ctrl + ;", "Insert today's date"],
      ["Alt + Enter", "New line inside a cell"],
      ["Ctrl + E", "Flash Fill"],
    ],
    practice: {
      task: "Fill A1:A12 with Jan to Dec, then put the value 100 into B1:B12 in a single action.",
      answer:
        "Type Jan in A1 and drag the fill handle to A12. Select B1:B12, type 100, press Ctrl + Enter.",
    },
    quiz: [
      {
        q: "Which keystroke fills every selected cell with what you just typed?",
        options: ["Enter", "Ctrl + Enter", "Alt + Enter", "Shift + Enter"],
        answer: 1,
        explain: "Ctrl + Enter commits the entry to the whole selection.",
      },
    ],
  },
  {
    slug: "formatting-cells",
    title: "Formatting cells and building custom number formats",
    track: "workbook",
    level: "Beginner",
    summary:
      "Number formats change how a value looks without changing the value, which is how you show units, hide zeros and colour negatives.",
    objectives: [
      "Apply the built-in number, alignment, font and border options",
      "Write custom format codes",
      "Copy formatting quickly with the Format Painter and cell styles",
    ],
    steps: [
      {
        title: "Open Format Cells with Ctrl + 1",
        detail:
          "Six tabs: Number, Alignment, Font, Border, Fill and Protection. Everything in this lesson lives here.",
      },
      {
        title: "Pick the right number category",
        detail:
          "General, Number (decimals and thousands separator), Currency, Accounting (aligned symbols), Date, Time, Percentage, Fraction, Scientific, Text and Custom.",
      },
      {
        title: "Custom codes: the four sections",
        detail:
          "positive;negative;zero;text. For example #,##0;[Red](#,##0);\"-\";@ shows thousands, red brackets for negatives, a dash for zero.",
        formula: '#,##0;[Red](#,##0);"-";@',
      },
      {
        title: "Show a unit without breaking the maths",
        detail:
          'Type the unit in quotes. The cell still contains a real number, so SUM keeps working.',
        formula: '0" kg"    0.0" lbs"    #,##0" pcs"',
      },
      {
        title: "Hide values or shorten big numbers",
        detail:
          "Three semicolons ;;; hides a cell's contents on screen. A trailing comma divides the display by a thousand.",
        formula: '#,##0,"k"      0.0,,"m"',
      },
      {
        title: "Alignment that avoids merged cells",
        detail:
          "Wrap Text, Shrink to Fit, Indent and vertical alignment are all here. Instead of Merge & Center, use Horizontal > Center Across Selection, which looks identical but does not break sorting or filtering.",
      },
      {
        title: "Copy formatting",
        detail:
          "Format Painter (Home > paintbrush) copies formatting once; double-click it to keep painting until you press Esc.",
      },
      {
        title: "Cell styles and themes",
        detail:
          "Home > Cell Styles defines reusable looks (Input, Calculation, Total). Page Layout > Themes changes the whole colour and font scheme in one click.",
      },
    ],
    example: {
      caption: "Custom format codes and what they show",
      headers: ["Code", "Value", "Displays"],
      rows: [
        ['0" kg"', "12", "12 kg"],
        ['#,##0;[Red]-#,##0', "-2500", "-2,500 in red"],
        ['0.00%', "0.4567", "45.67%"],
        ['dd mmm yyyy', "45355", "04 Mar 2024"],
        [';;;', "anything", "nothing visible"],
        ['[>=1000]0,"k";0', "2500", "2k"],
      ],
    },
    shortcuts: [
      ["Ctrl + 1", "Format Cells"],
      ["Ctrl + Shift + 1", "Number, 2 decimals"],
      ["Ctrl + Shift + 5", "Percentage"],
      ["Ctrl + Shift + 3", "Date format"],
      ["Alt + H, O, I", "Autofit column width"],
    ],
    practice: {
      task: "Show weights as '12.5 kg', negatives in red brackets and zeros as a dash, while keeping SUM working.",
      answer: '0.0" kg";[Red](0.0" kg");"-"',
    },
    quiz: [
      {
        q: "Why is a custom number format better than typing '12 kg'?",
        options: [
          "It is faster to type",
          "The cell keeps a real number so calculations still work",
          "It prints better",
          "It uses less memory",
        ],
        answer: 1,
        explain: "Typing the unit converts the entry to text and breaks arithmetic.",
      },
    ],
  },
  {
    slug: "rows-columns-worksheets",
    title: "Managing rows, columns and worksheets",
    track: "workbook",
    level: "Beginner",
    summary:
      "Insert, delete, resize, hide and group the structure of your workbook without breaking the formulas that point at it.",
    objectives: [
      "Insert, delete, hide and unhide rows, columns and sheets",
      "Resize precisely and autofit",
      "Group sheets to edit several at once",
    ],
    steps: [
      {
        title: "Select first",
        detail:
          "Shift + Space selects the row, Ctrl + Space selects the column. Then Ctrl + Shift + + inserts and Ctrl + - deletes.",
      },
      {
        title: "Resize",
        detail:
          "Drag a border, or double-click it to autofit. Right-click > Row Height / Column Width sets an exact value for several selected rows at once.",
      },
      {
        title: "Hide and unhide",
        detail:
          "Ctrl + 9 hides rows, Ctrl + 0 hides columns. To unhide column A, use the Name Box: type A1, press Enter, then Home > Format > Hide & Unhide > Unhide Columns.",
      },
      {
        title: "Group and outline",
        detail:
          "Data > Group creates collapsible sections with the 1 / 2 / 3 buttons on the left. Data > Subtotal builds the groups and the SUBTOTAL formulas for you.",
      },
      {
        title: "Move rows without cut and paste",
        detail:
          "Select the row, hold Shift and drag its border. The other rows shuffle out of the way instead of being overwritten.",
      },
      {
        title: "Worksheets",
        detail:
          "Shift + F11 inserts a sheet. Right-click a tab for Rename, Tab Color, Move or Copy (tick 'Create a copy'), Hide and Protect Sheet.",
      },
      {
        title: "Group sheets to edit them together",
        detail:
          "Ctrl-click several tabs; anything you type or format applies to all of them. Click a single tab afterwards to ungroup - forgetting this is a classic way to wreck twelve sheets at once.",
      },
      {
        title: "3-D formulas across sheets",
        detail: "A single formula can total the same cell on every sheet.",
        formula: "=SUM(Jan:Dec!B5)",
      },
    ],
    shortcuts: [
      ["Shift + Space / Ctrl + Space", "Select row / column"],
      ["Ctrl + Shift + +", "Insert"],
      ["Ctrl + -", "Delete"],
      ["Shift + F11", "New worksheet"],
      ["Ctrl + 9 / Ctrl + 0", "Hide row / column"],
    ],
    practice: {
      task: "Insert three rows above row 5 and autofit every column in one action.",
      answer:
        "Select rows 5:7, press Ctrl + Shift + +. Then Ctrl + A to select all and press Alt + H, O, I.",
    },
    quiz: [
      {
        q: "How do you unhide column A?",
        options: [
          "Right-click column B",
          "Type A1 in the Name Box, then Home > Format > Unhide Columns",
          "Delete column B",
          "It cannot be unhidden",
        ],
        answer: 1,
        explain: "You must first select the hidden column, and the Name Box is the way in.",
      },
    ],
  },
  {
    slug: "named-ranges",
    title: "Named ranges and structured Table references",
    track: "workbook",
    level: "Intermediate",
    summary:
      "Names turn =SUMIFS($B$2:$B$5000;$C$2:$C$5000;E2) into =SUMIFS(Amount;Region;E2). Readable formulas are correct formulas.",
    objectives: [
      "Create, edit and delete names",
      "Use Tables so ranges grow automatically",
      "Build a dynamic dropdown from a name",
    ],
    steps: [
      {
        title: "The quickest way to name a range",
        detail: "Select the cells, click the Name Box, type the name and press Enter.",
      },
      {
        title: "Create from selection",
        detail:
          "Select the data including its headers, then Formulas > Create from Selection > Top row. Every column gets a name from its header in one step.",
      },
      {
        title: "The Name Manager",
        detail:
          "Ctrl + F3 lists every name, its scope (workbook or one sheet), what it refers to and a comment field. Delete unused and #REF! names here.",
      },
      {
        title: "Naming rules",
        detail:
          "No spaces, cannot look like a cell reference (Q1 is illegal, Qtr1 is fine), case-insensitive, up to 255 characters. Use underscores or CamelCase.",
      },
      {
        title: "Tables are self-naming ranges",
        detail:
          "Ctrl + T turns a range into a Table. Formulas then read =SUM(Sales[Amount]) and the range grows automatically when you add rows.",
        formula: "=SUMIFS(Sales[Amount];Sales[Region];$E2)",
      },
      {
        title: "Table shorthand",
        detail:
          "Sales[#Headers], Sales[#Totals], Sales[@Amount] for this row, and Sales[[Name]:[Price]] for a block of columns.",
      },
      {
        title: "A dropdown that maintains itself",
        detail:
          "Name a spilled UNIQUE result, then Data > Data Validation > List > Source: =Departments (or =$H$2# to point at the spill range).",
      },
      {
        title: "Names can hold constants and formulas",
        detail:
          "A name does not have to be a range. VATRate can simply refer to =0.2, and you can then write =B2*VATRate everywhere.",
      },
    ],
    shortcuts: [
      ["Ctrl + F3", "Name Manager"],
      ["F3", "Paste a name into a formula"],
      ["Ctrl + T", "Create a Table"],
    ],
    practice: {
      task: "Name column B 'Amount' and column C 'Region', then total the North rows using names only.",
      answer: '=SUMIF(Region;"North";Amount)',
    },
    quiz: [
      {
        q: "Which name is illegal in Excel?",
        options: ["Qtr1", "Tax_Rate", "Q1", "SalesData"],
        answer: 2,
        explain: "Q1 looks like a cell reference, so Excel rejects it.",
      },
    ],
  },
  {
    slug: "data-validation",
    title: "Data Validation: dropdown lists and input rules",
    track: "workbook",
    level: "Intermediate",
    summary:
      "Stop bad data at the door. Validation restricts what may be typed, shows a helpful prompt and can even cascade one list from another.",
    objectives: [
      "Create list, number, date and text-length rules",
      "Add input messages and error alerts",
      "Build a dependent (cascading) dropdown and find invalid data",
    ],
    steps: [
      {
        title: "Open the dialog",
        detail: "Select the cells, then Data > Data Validation (Alt + A, V, V).",
      },
      {
        title: "Allow: List",
        detail:
          "Type the options separated by commas, or point at a range or a name. Tick 'In-cell dropdown'.",
        formula: "Source: =Departments",
      },
      {
        title: "Allow: Whole number, Decimal, Date, Time, Text length",
        detail:
          "Each gives between / not between / greater than options. Example: a date between today and one year ahead.",
        formula: "Start date: =TODAY()   End date: =TODAY()+365",
      },
      {
        title: "Allow: Custom formula",
        detail:
          "Anything that returns TRUE is accepted. Force unique entries or a required prefix.",
        formula: '=COUNTIF($A:$A;A2)=1        =LEFT(A2;3)="INV"',
      },
      {
        title: "Input Message",
        detail:
          "The second tab shows a small yellow tooltip when the cell is selected. Use it to explain the expected format.",
      },
      {
        title: "Error Alert",
        detail:
          "Stop blocks the entry, Warning allows it after confirmation, Information just notifies. Write a message that says what IS allowed.",
      },
      {
        title: "Dependent dropdowns",
        detail:
          "Name each sub-list exactly as its parent value, then set the second dropdown's source to INDIRECT of the first cell.",
        formula: "Source: =INDIRECT($A2)",
      },
      {
        title: "Audit existing data",
        detail:
          "Data > Data Validation > Circle Invalid Data rings every value that breaks the rule, even if it was pasted in. 'Clear Validation Circles' removes the rings.",
      },
      {
        title: "The paste loophole",
        detail:
          "Pasting over a validated cell can replace the rule. Protect the sheet as well if the workbook goes to other people.",
      },
    ],
    practice: {
      task: "Column C must only accept a date in the current year, with a helpful message.",
      answer:
        'Data Validation > Allow: Date > between > Start =DATE(YEAR(TODAY());1;1) End =DATE(YEAR(TODAY());12;31), then write the Input Message and a Stop error alert.',
    },
    quiz: [
      {
        q: "Which feature builds a cascading dropdown?",
        options: ["VLOOKUP", "INDIRECT with named ranges", "Conditional Formatting", "Flash Fill"],
        answer: 1,
        explain: "INDIRECT converts the chosen text into the name of the matching sub-list.",
      },
    ],
  },
  {
    slug: "find-replace-goto",
    title: "Find, Replace and Go To Special",
    track: "workbook",
    level: "Beginner",
    summary:
      "Three dialogs that fix in seconds what people fix manually for hours: search with wildcards, replace across a workbook, and select exactly the cells you need.",
    objectives: [
      "Search with options, wildcards and formats",
      "Replace across the whole workbook",
      "Select blanks, formulas, constants and visible cells only",
    ],
    steps: [
      {
        title: "Find (Ctrl + F)",
        detail:
          "Click Options to search the whole workbook, match case, match entire cell contents, or look in Values instead of Formulas.",
      },
      {
        title: "Replace (Ctrl + H)",
        detail:
          "Replace All reports how many changes it made. Leaving 'Replace with' empty deletes the text found.",
      },
      {
        title: "Wildcards",
        detail:
          "* is any run of characters, ? is one character, ~ escapes a literal * or ?. Replace 'INV-*' with nothing to strip prefixes.",
      },
      {
        title: "Replace formats",
        detail:
          "The Format buttons let you find every cell with a red fill and replace the formatting rather than the value.",
      },
      {
        title: "Go To Special (F5 then Alt + S)",
        detail:
          "Choose Blanks, Constants, Formulas, Errors, Comments, Visible cells only, Row differences or Conditional formats.",
      },
      {
        title: "The classic blanks trick",
        detail:
          "Select the column, Go To Special > Blanks, type = then press the Up arrow, then Ctrl + Enter. Every gap fills with the value above.",
      },
      {
        title: "Copy visible cells only",
        detail:
          "After filtering, press Alt + ; before Ctrl + C so hidden rows are not copied.",
      },
      {
        title: "Find every formula",
        detail:
          "Go To Special > Formulas highlights every calculated cell - the fastest audit of an inherited workbook.",
      },
    ],
    shortcuts: [
      ["Ctrl + F / Ctrl + H", "Find / Replace"],
      ["F5 then Alt + S", "Go To Special"],
      ["Alt + ;", "Visible cells only"],
      ["Ctrl + `", "Show all formulas"],
    ],
    practice: {
      task: "An import left every second cell in column B empty. Fill each gap with the value above it.",
      answer:
        "Select column B, F5 > Special > Blanks > OK, type = then press Up arrow, then Ctrl + Enter. Finally copy and Paste Special > Values.",
    },
    quiz: [
      {
        q: "Which shortcut selects only the visible cells after filtering?",
        options: ["Ctrl + A", "Alt + ;", "Ctrl + Shift + V", "F4"],
        answer: 1,
        explain: "Alt + ; is Go To Special > Visible cells only.",
      },
    ],
  },
  {
    slug: "printing-page-layout",
    title: "Printing and page layout that actually fits",
    track: "workbook",
    level: "Beginner",
    summary:
      "Repeat headings on every page, force sensible page breaks and fit a wide report to one page wide without shrinking it to nothing.",
    objectives: [
      "Set a print area, titles, margins and scaling",
      "Control page breaks",
      "Add headers, footers and page numbers",
    ],
    steps: [
      {
        title: "Set the print area",
        detail:
          "Select the range, then Page Layout > Print Area > Set Print Area. Clear Print Area removes it.",
      },
      {
        title: "Repeat the header row on every page",
        detail:
          "Page Layout > Print Titles > 'Rows to repeat at top' and click row 1. Use 'Columns to repeat at left' for wide reports.",
        formula: "Rows to repeat at top: $1:$1",
      },
      {
        title: "Fit to one page wide",
        detail:
          "Page Layout > Scale to Fit: set Width to 1 page and leave Height blank (Automatic). Setting both to 1 is what makes the text unreadable.",
      },
      {
        title: "Orientation, size and margins",
        detail:
          "Landscape for wide tables. Custom Margins has a 'Center on page: Horizontally' tick box that instantly improves any printout.",
      },
      {
        title: "Page breaks",
        detail:
          "View > Page Break Preview shows blue lines you can drag. Page Layout > Breaks > Insert Page Break forces a break above and left of the selected cell.",
      },
      {
        title: "Headers and footers",
        detail:
          "Insert > Header & Footer, or Page Setup > Header/Footer > Custom. Use the buttons for page number, total pages, date, file path and sheet name.",
      },
      {
        title: "Sheet options worth knowing",
        detail:
          "Page Setup > Sheet lets you print Gridlines, Row and column headings, Comments, and choose 'Over, then down' page order for wide reports.",
      },
      {
        title: "Preview and export",
        detail:
          "Ctrl + P shows a live preview. File > Export > Create PDF/XPS produces a PDF that keeps your print settings.",
      },
    ],
    shortcuts: [
      ["Ctrl + P", "Print preview"],
      ["Alt + P, R, S", "Set print area"],
      ["Alt + P, I", "Print titles"],
    ],
    practice: {
      task: "A 40-page list must repeat its header row and fit one page wide, with 'Page 1 of 40' in the footer.",
      answer:
        "Print Titles > Rows to repeat at top = $1:$1. Scale to Fit > Width = 1 page, Height = Automatic. Custom Footer > insert Page Number, type ' of ', insert Number of Pages.",
    },
    quiz: [
      {
        q: "How do you repeat the header row on every printed page?",
        options: [
          "Freeze Panes",
          "Page Layout > Print Titles > Rows to repeat at top",
          "Copy it manually",
          "Insert a header",
        ],
        answer: 1,
        explain: "Freeze Panes only affects the screen; Print Titles affects paper.",
      },
    ],
  },
  {
    slug: "protect-share-workbook",
    title: "Protecting, sharing and finalising a workbook",
    track: "workbook",
    level: "Intermediate",
    summary:
      "Lock the formulas, leave the inputs open, and hand a workbook to colleagues without it coming back broken.",
    objectives: [
      "Lock and unlock cells and protect a sheet or workbook",
      "Hide formulas and structure",
      "Use co-authoring, track changes and document inspection",
    ],
    steps: [
      {
        title: "Understand the two-stage rule",
        detail:
          "Every cell is Locked by default, but locking only takes effect once the sheet is protected. So you unlock the input cells first, then protect.",
      },
      {
        title: "Unlock the inputs",
        detail:
          "Select the cells people should fill in, press Ctrl + 1 > Protection tab, untick Locked.",
      },
      {
        title: "Protect the sheet",
        detail:
          "Review > Protect Sheet. Set a password if you need one, and tick what users may still do - usually 'Select unlocked cells', 'Sort' and 'Use AutoFilter'.",
      },
      {
        title: "Hide the formulas",
        detail:
          "In the same Protection tab tick Hidden. Once protected, the Formula Bar shows nothing for those cells.",
      },
      {
        title: "Protect the structure",
        detail:
          "Review > Protect Workbook stops sheets being added, deleted, renamed, moved or unhidden.",
      },
      {
        title: "Encrypt the file",
        detail:
          "File > Info > Protect Workbook > Encrypt with Password. This one cannot be recovered if forgotten, unlike sheet protection.",
      },
      {
        title: "Mark as Final and Inspect",
        detail:
          "File > Info > Check for Issues > Inspect Document removes hidden rows, personal data and comments before you send the file out. 'Mark as Final' makes it read-only as a courtesy, not as security.",
      },
      {
        title: "Share and co-author",
        detail:
          "Save to OneDrive or SharePoint, then Share. Several people can edit at once and you can see their cursors. Review > Comments (@mention someone) replaces the old Track Changes for most teams.",
      },
      {
        title: "Version history",
        detail:
          "File > Info > Version History restores an earlier copy of a cloud-saved workbook - the real undo for shared files.",
      },
    ],
    practice: {
      task: "Let users type only into B2:B10, hide your formulas and stop anyone renaming the sheets.",
      answer:
        "Select B2:B10, Ctrl + 1 > Protection > untick Locked. Select the formula cells, tick Hidden. Review > Protect Sheet. Then Review > Protect Workbook > Structure.",
    },
    quiz: [
      {
        q: "Why does unticking 'Locked' do nothing on its own?",
        options: [
          "It is a bug",
          "Locking only applies once the sheet is protected",
          "You need a password first",
          "It only works in Tables",
        ],
        answer: 1,
        explain: "Protection is the switch that activates the Locked attribute.",
      },
    ],
  },
  {
    slug: "what-if-analysis",
    title: "What-If Analysis: Goal Seek, Data Tables and Scenarios",
    track: "workbook",
    level: "Advanced",
    summary:
      "Run your model backwards, test hundreds of combinations at once, and store alternative sets of assumptions.",
    objectives: [
      "Use Goal Seek to find the input that gives a target result",
      "Build one and two variable Data Tables",
      "Save and compare Scenarios, and use Solver for constrained problems",
    ],
    steps: [
      {
        title: "Goal Seek: the reverse calculation",
        detail:
          "Data > What-If Analysis > Goal Seek. 'Set cell' must contain a formula, 'To value' is your target, 'By changing cell' must be a plain input number.",
        formula: "Set B8 to 1200 by changing B1 (loan amount)",
      },
      {
        title: "One-variable Data Table",
        detail:
          "List the input values down a column. In the cell one row up and one column right, put =your result cell. Select the whole block, then Data Table with the Column input cell set to your input.",
      },
      {
        title: "Two-variable Data Table",
        detail:
          "Rates down the left, terms across the top, the formula reference in the top-left corner cell. Fill in both Row input cell and Column input cell.",
      },
      {
        title: "Read the result",
        detail:
          "Excel fills the grid with a hidden TABLE() array formula. You cannot edit individual cells - delete the whole block to remove it.",
      },
      {
        title: "Scenario Manager",
        detail:
          "Data > What-If Analysis > Scenario Manager > Add. Name it (Best case), pick the changing cells, enter the values. Repeat for Worst and Base, then click Summary for a side-by-side report.",
      },
      {
        title: "Solver for constrained optimisation",
        detail:
          "File > Options > Add-ins > Excel Add-ins > Solver. Set an objective (maximise profit), the variable cells, and constraints such as hours <= 40 and quantities as integers.",
      },
      {
        title: "Forecast Sheet",
        detail:
          "Data > Forecast Sheet takes a date column and a value column and produces a projection with confidence bounds in one click.",
      },
      {
        title: "Keep the model auditable",
        detail:
          "Inputs in one clearly shaded block, calculations separate, no hard-coded numbers inside formulas. Every what-if tool depends on that discipline.",
      },
    ],
    practice: {
      task: "Your monthly budget must be 1 200. Find the maximum loan you can borrow at 5.5% over 25 years.",
      answer:
        "Build =-PMT(B2/12;B3*12;B1) in B8, then Goal Seek: Set cell B8, To value 1200, By changing cell B1.",
    },
    quiz: [
      {
        q: "Goal Seek can change:",
        options: [
          "Any cell",
          "Only a cell containing a plain number, not a formula",
          "Only formulas",
          "Only named ranges",
        ],
        answer: 1,
        explain: "The changing cell must be a hard input value.",
      },
    ],
  },
  {
    slug: "formula-auditing",
    title: "Formula auditing and error checking",
    track: "workbook",
    level: "Intermediate",
    summary:
      "Trace where numbers come from, step through a formula one calculation at a time, and find the circular reference that is blocking the workbook.",
    objectives: [
      "Trace precedents and dependents",
      "Evaluate a formula step by step",
      "Diagnose every Excel error and fix circular references",
    ],
    steps: [
      {
        title: "Show all formulas",
        detail: "Ctrl + ` toggles the whole sheet between results and formulas.",
      },
      {
        title: "Trace Precedents and Dependents",
        detail:
          "Formulas > Trace Precedents draws arrows to the cells feeding this one; Trace Dependents shows what relies on it. Remove Arrows clears them. Double-click a dashed arrow to jump to another sheet.",
      },
      {
        title: "Evaluate Formula",
        detail:
          "Formulas > Evaluate Formula steps through the calculation one piece at a time - the single best tool for a formula that returns the wrong number.",
      },
      {
        title: "The F9 trick",
        detail:
          "In the formula bar, select part of a formula and press F9 to see what that fragment evaluates to. Press Esc, never Enter, afterwards.",
      },
      {
        title: "Error Checking",
        detail:
          "Formulas > Error Checking walks through every flagged cell. The green triangle menu explains the issue and offers a fix.",
      },
      {
        title: "The Watch Window",
        detail:
          "Formulas > Watch Window pins key cells so you can see them update while you work on another sheet.",
      },
      {
        title: "Circular references",
        detail:
          "The status bar names the offending cell. Formulas > Error Checking > Circular References lists them. Fix the logic, or enable iterative calculation only if the model genuinely needs it.",
      },
      {
        title: "The error decoder",
        detail:
          "#NAME? misspelled function or missing quotes. #VALUE! text where a number is expected. #REF! deleted reference. #DIV/0! division by empty or zero. #N/A lookup found nothing. #SPILL! something blocks a dynamic array. #NUM! impossible maths. #CALC! empty array result.",
      },
    ],
    shortcuts: [
      ["Ctrl + `", "Show all formulas"],
      ["Ctrl + [", "Jump to precedents"],
      ["Ctrl + ]", "Jump to dependents"],
      ["F9 (in formula bar)", "Evaluate the selected fragment"],
    ],
    practice: {
      task: "A total is 10% too high and you cannot see why. Which tool do you reach for first?",
      answer:
        "Select the cell, Formulas > Evaluate Formula, and step through until a value looks wrong. Then Trace Precedents on that part.",
    },
    quiz: [
      {
        q: "Which tool steps through a formula one calculation at a time?",
        options: ["Trace Precedents", "Evaluate Formula", "Watch Window", "Error Checking"],
        answer: 1,
        explain: "Evaluate Formula shows each intermediate result.",
      },
    ],
  },
  {
    slug: "import-export-data",
    title: "Importing and exporting: CSV, text, web and PDF",
    track: "workbook",
    level: "Intermediate",
    summary:
      "Get data in without mangling dates and leading zeros, and get it out in the format the next system expects.",
    objectives: [
      "Import CSV and text files with the right data types",
      "Pull a table from a web page",
      "Export to CSV and PDF safely",
    ],
    steps: [
      {
        title: "Never just double-click a CSV",
        detail:
          "Opening a CSV directly lets Excel guess: leading zeros vanish and 03/04 becomes a date. Use Data > From Text/CSV instead.",
      },
      {
        title: "Import with Power Query",
        detail:
          "Data > From Text/CSV, check the delimiter and the File Origin (UTF-8 for accented characters), then click Transform Data to set each column type before loading.",
      },
      {
        title: "Force problem columns to Text",
        detail:
          "In the query editor, click the column type icon and choose Text for account numbers, postcodes and product codes.",
      },
      {
        title: "Fix foreign dates",
        detail:
          "Use Change Type > Using Locale and pick English (United Kingdom) or whatever produced the file.",
      },
      {
        title: "From Web",
        detail:
          "Data > From Web, paste the URL, and pick the table from the navigator. Refresh pulls the latest version.",
      },
      {
        title: "Refresh and connection properties",
        detail:
          "Data > Queries & Connections > right-click > Properties to refresh on open or every n minutes.",
      },
      {
        title: "Export to CSV",
        detail:
          "File > Save As > CSV UTF-8. Only the active sheet is saved, formulas become values and formatting is lost - keep the .xlsx as your master.",
      },
      {
        title: "Export to PDF",
        detail:
          "File > Export > Create PDF/XPS. Options lets you choose the selection, the active sheet or the entire workbook, and it obeys your print area and titles.",
      },
    ],
    practice: {
      task: "A CSV of invoices keeps losing the leading zeros on invoice numbers.",
      answer:
        "Data > From Text/CSV > Transform Data > click the invoice column > Data Type > Text > Close & Load.",
    },
    quiz: [
      {
        q: "Best way to open a CSV without losing leading zeros:",
        options: [
          "Double-click it",
          "Data > From Text/CSV and set the column to Text",
          "Rename it to .xlsx",
          "Format the column afterwards",
        ],
        answer: 1,
        explain: "Once Excel has parsed the value the zeros are already gone.",
      },
    ],
  },
  {
    slug: "keyboard-shortcuts-training",
    title: "Keyboard shortcuts: a structured drill",
    track: "workbook",
    level: "Beginner",
    summary:
      "Not a list of 200 to memorise. Six small groups, learned one group per day, that remove most of the mouse from your working week.",
    objectives: [
      "Navigate and select without the mouse",
      "Enter, format and structure data from the keyboard",
      "Use the Alt KeyTips system to reach any ribbon command",
    ],
    steps: [
      {
        title: "Day 1 - move",
        detail:
          "Ctrl + Arrow jumps to the edge of the data. Ctrl + Home returns to A1. Ctrl + Page Up/Down changes sheet. F5 goes anywhere.",
      },
      {
        title: "Day 2 - select",
        detail:
          "Add Shift to any movement key to select as you go. Ctrl + Shift + Arrow grabs a whole column of data. Ctrl + A selects the current region. Shift + Space and Ctrl + Space take the row and column.",
      },
      {
        title: "Day 3 - enter and edit",
        detail:
          "F2 edits, Alt + Enter adds a line inside the cell, Ctrl + Enter fills the selection, Ctrl + D and Ctrl + R fill down and right, Ctrl + E is Flash Fill, F4 toggles the dollar signs.",
      },
      {
        title: "Day 4 - format",
        detail:
          "Ctrl + 1 is Format Cells and covers everything. Ctrl + B, I, U. Ctrl + Shift + 1 number, 5 percent, 3 date, 4 currency. Alt + H, O, I autofits.",
      },
      {
        title: "Day 5 - structure",
        detail:
          "Ctrl + T table, Ctrl + Shift + L filters, Ctrl + Shift + + insert, Ctrl + - delete, Alt + = AutoSum, Alt + N + V PivotTable.",
      },
      {
        title: "Day 6 - the master key",
        detail:
          "Press and release Alt. Letters appear over every tab and button. Alt then H then O then I means Home, Format, Autofit. You no longer need to memorise anything else.",
      },
      {
        title: "Build your own",
        detail:
          "Right-click any command > Add to Quick Access Toolbar, then use Alt + 1 to Alt + 9. Macros can also be given a Ctrl + Shift + letter shortcut.",
      },
    ],
    shortcuts: [
      ["Ctrl + Arrow", "Edge of the data"],
      ["Ctrl + Shift + Arrow", "Select to the edge"],
      ["Ctrl + Home", "Back to A1"],
      ["F2", "Edit the cell"],
      ["F4", "Toggle absolute references"],
      ["Ctrl + Enter", "Fill the selection"],
      ["Ctrl + 1", "Format Cells"],
      ["Ctrl + T", "Create a Table"],
      ["Ctrl + Shift + L", "Toggle filters"],
      ["Alt + =", "AutoSum"],
      ["Alt + ;", "Visible cells only"],
      ["Alt", "Show ribbon KeyTips"],
    ],
    practice: {
      task: "Without touching the mouse: select the whole data block, make it a Table, add filters and total the last column.",
      answer:
        "Ctrl + A, Ctrl + T, Enter, Ctrl + Shift + L, then Ctrl + Down + one row and Alt + =.",
    },
    quiz: [
      {
        q: "What does pressing and releasing Alt do?",
        options: [
          "Nothing",
          "Shows KeyTip letters for every ribbon command",
          "Opens the File menu",
          "Cancels the entry",
        ],
        answer: 1,
        explain: "KeyTips let you reach any command without memorising a shortcut.",
      },
    ],
  },
];
