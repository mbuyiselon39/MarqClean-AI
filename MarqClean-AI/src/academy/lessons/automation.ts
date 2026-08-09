import type { Lesson } from "./types";

export const automation: Lesson[] = [
  {
    slug: "macros",
    title: "Macros - record repetitive work once and replay it forever",
    track: "automation",
    level: "Intermediate",
    summary:
      "The Macro Recorder writes VBA for you while you work. Learn to record cleanly, then edit the result into something reusable.",
    objectives: [
      "Show the Developer tab and record a macro",
      "Understand relative vs absolute recording",
      "Store macros so they are available in every workbook",
    ],
    steps: [
      {
        title: "Turn on the Developer tab",
        detail: "File → Options → Customize Ribbon → tick Developer.",
      },
      {
        title: "Plan the steps first",
        detail:
          "The recorder captures everything, including your mistakes. Rehearse the sequence once before you press record.",
      },
      {
        title: "Record",
        detail:
          "Developer → Record Macro. Give it a name with no spaces (FormatReport), an optional Ctrl+Shift shortcut, and choose where to store it.",
      },
      {
        title: "Absolute vs relative",
        detail:
          "By default the recorder writes exact cell addresses. Click 'Use Relative References' first if the macro should work wherever the cursor happens to be.",
      },
      {
        title: "Stop and test",
        detail:
          "Developer → Stop Recording. Undo your changes and run the macro (Alt + F8) to check it reproduces them.",
      },
      {
        title: "Clean up the code",
        detail:
          "Alt + F11 → find your module. Delete the .Select / .Activate lines the recorder loves and work with objects directly.",
        formula:
          "' Recorded:\nRange(\"A1\").Select\nSelection.Font.Bold = True\n' Cleaned:\nRange(\"A1\").Font.Bold = True",
      },
      {
        title: "Store in PERSONAL.XLSB",
        detail:
          "Choosing 'Personal Macro Workbook' when recording creates a hidden workbook that opens with Excel, making the macro available everywhere.",
      },
      {
        title: "Save as .xlsm",
        detail: "A normal .xlsx cannot contain macros - Excel will silently strip them.",
      },
    ],
    formulas: [
      {
        label: "Unhide every sheet",
        code: "Sub UnhideAllSheets()\n    Dim ws As Worksheet\n    For Each ws In ActiveWorkbook.Worksheets\n        ws.Visible = xlSheetVisible\n    Next ws\nEnd Sub",
      },
      {
        label: "Remove every filter on the sheet",
        code: "Sub ClearFilters()\n    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData\nEnd Sub",
      },
      {
        label: "Save a copy with today's date",
        code: "Sub SaveDated()\n    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & \"\\\" & Format(Date, \"yyyy-mm-dd\") & \".xlsm\"\nEnd Sub",
      },
    ],
    shortcuts: [
      ["Alt + F8", "Macro list"],
      ["Alt + F11", "VBA editor"],
      ["F5", "Run the macro in the editor"],
      ["F8", "Step through line by line"],
    ],
    practice: {
      task: "Automate: freeze the header row, autofit columns and apply a filter.",
      answer:
        "Sub TidySheet()\n    Rows(2).Select: ActiveWindow.FreezePanes = True\n    Cells.EntireColumn.AutoFit\n    If Not ActiveSheet.AutoFilterMode Then Range(\"A1\").AutoFilter\nEnd Sub",
    },
    quiz: [
      {
        q: "Which file format keeps macros?",
        options: [".xlsx", ".xlsm", ".csv", ".xlst"],
        answer: 1,
        explain: "Macro-enabled workbooks use .xlsm (or .xlsb).",
      },
    ],
  },
  {
    slug: "run-vba-macro",
    title: "How to run a VBA macro code - step by step",
    track: "automation",
    level: "Beginner",
    summary:
      "Someone gave you a block of VBA. Here is exactly where to paste it, how to run it, and how to make it safe.",
    objectives: [
      "Paste code into the right place",
      "Run it four different ways",
      "Handle macro security warnings",
    ],
    steps: [
      { title: "1. Save your work first", detail: "Macros usually cannot be undone with Ctrl + Z. Save, or work on a copy." },
      { title: "2. Open the VBA editor", detail: "Press Alt + F11." },
      {
        title: "3. Insert a module",
        detail:
          "In the Project Explorer on the left, find your workbook (VBAProject (Book1)), then Insert → Module. A blank white code window appears.",
      },
      {
        title: "4. Paste the code",
        detail:
          "Click into the module and press Ctrl + V. The code must begin with Sub Something() and end with End Sub.",
      },
      {
        title: "5. Run it",
        detail:
          "Click anywhere inside the Sub and press F5. Or go back to Excel (Alt + F11) and press Alt + F8, select the macro, click Run.",
      },
      {
        title: "6. Other ways to trigger it",
        detail:
          "Assign it to a button (Developer → Insert → Button), to a shape (right-click → Assign Macro), or to a Quick Access Toolbar icon you can hit with Alt + 1.",
      },
      {
        title: "7. Save as .xlsm",
        detail: "File → Save As → Excel Macro-Enabled Workbook, otherwise your code disappears.",
      },
      {
        title: "8. If nothing happens",
        detail:
          "The yellow Security Warning bar → Enable Content. Or File → Options → Trust Center → Trust Center Settings → Macro Settings. Files downloaded from the internet also need Properties → Unblock.",
      },
      {
        title: "9. Debugging basics",
        detail:
          "F8 steps line by line, and hovering over a variable shows its value. A yellow highlighted line is where it stopped - read the message, it is usually a missing sheet name.",
      },
    ],
    formulas: [
      {
        label: "A safe first macro to test with",
        code: "Sub HelloExcel()\n    MsgBox \"The macro ran successfully on \" & ActiveSheet.Name\nEnd Sub",
      },
    ],
    practice: {
      task: "Run the HelloExcel macro above and then change the message to include the workbook name.",
      answer: "MsgBox \"Ran in \" & ThisWorkbook.Name",
    },
    quiz: [
      {
        q: "Which shortcut opens the macro list in Excel?",
        options: ["Alt + F11", "Alt + F8", "Ctrl + M", "F5"],
        answer: 1,
        explain: "Alt + F8 lists macros; Alt + F11 opens the editor.",
      },
    ],
  },
  {
    slug: "cheat-sheet",
    title: "Excel Online Cheat Sheet - every shortcut and formula pattern",
    track: "automation",
    level: "Beginner",
    summary:
      "One page to keep open beside your work: the shortcuts that matter, the formula patterns you reuse, and what each error message means.",
    objectives: ["Find the right shortcut fast", "Recognise every error value", "Copy a ready-made formula pattern"],
    steps: [
      {
        title: "How to use this page",
        detail:
          "Skim it once, then come back with Ctrl + F. The goal is recognition, not memorisation - you will absorb the ten you actually use.",
      },
      {
        title: "Error decoder",
        detail:
          "#NAME? - misspelled function or missing quotes. #VALUE! - text where a number is expected. #REF! - a referenced cell was deleted. #DIV/0! - dividing by empty or zero. #N/A - lookup found nothing. #SPILL! - something blocks a dynamic array. #NUM! - impossible maths. #CALC! - an empty array result.",
      },
      {
        title: "Formula patterns worth copying",
        detail: "The ten formulas below solve about 80% of everyday office spreadsheet tasks.",
      },
    ],
    formulas: [
      { label: "Safe lookup", code: "=IFNA(XLOOKUP(A2;Ids;Names);\"Not found\")" },
      { label: "Conditional total", code: "=SUMIFS(Amount;Region;$E2;Date;\">=\"&$F$1;Date;\"<=\"&$F$2)" },
      { label: "Live unique list", code: "=SORT(UNIQUE(FILTER(A2:A500;A2:A500<>\"\")))" },
      { label: "Running total", code: "=SUM($B$2:B2)" },
      { label: "Percent of total", code: "=IFERROR(B2/$B$20;\"\")" },
      { label: "Highlight the row rule", code: "=$D2=\"Overdue\"" },
      { label: "Split at a delimiter", code: "=TEXTBEFORE(A2;\"-\")   =TEXTAFTER(A2;\"-\";-1)" },
      { label: "Clean imported text", code: "=TRIM(CLEAN(SUBSTITUTE(A2;CHAR(160);\" \")))" },
      { label: "Month-end due date", code: "=EOMONTH(A2;1)" },
      { label: "Visible-rows total", code: "=SUBTOTAL(109;B2:B500)" },
    ],
    shortcuts: [
      ["Ctrl + Arrow", "Jump to the edge of the data"],
      ["Ctrl + Shift + Arrow", "Select to the edge"],
      ["Ctrl + T", "Create a Table"],
      ["Ctrl + Shift + L", "Toggle filters"],
      ["Alt + =", "AutoSum"],
      ["Ctrl + E", "Flash Fill"],
      ["Ctrl + D / Ctrl + R", "Fill down / right"],
      ["Ctrl + Enter", "Fill all selected cells"],
      ["Alt + Enter", "New line inside a cell"],
      ["Ctrl + 1", "Format Cells"],
      ["Ctrl + Shift + 5", "Percent format"],
      ["Ctrl + ;", "Insert today's date"],
      ["F2", "Edit cell"],
      ["F4", "Toggle $ references"],
      ["F5 → Alt + S", "Go To Special"],
      ["Alt + ;", "Select visible cells only"],
      ["Ctrl + `", "Show all formulas"],
      ["Alt + N + V", "Insert PivotTable"],
      ["Alt + F8 / Alt + F11", "Macros / VBA editor"],
      ["Ctrl + Page Up/Down", "Move between sheets"],
    ],
    practice: {
      task: "Without looking: which shortcut fills every selected cell with what you just typed?",
      answer: "Ctrl + Enter",
    },
    quiz: [
      {
        q: "#SPILL! means:",
        options: [
          "A circular reference",
          "Something is blocking a dynamic array result",
          "A missing function",
          "Too many decimals",
        ],
        answer: 1,
        explain: "Clear the cells in the way and the formula spills.",
      },
    ],
  },
  {
    slug: "blank-worksheet",
    title: "Blank Excel Online Worksheet - practise right here",
    track: "automation",
    level: "Beginner",
    summary:
      "A live grid built into this page. Type values, write formulas starting with =, and watch them calculate - no install needed.",
    objectives: [
      "Type values and formulas in a real grid",
      "Practise SUM, AVERAGE, IF and cell references",
    ],
    steps: [
      {
        title: "Open the practice grid below the lesson",
        detail:
          "Every lesson page has a live worksheet at the bottom. Click a cell, type, press Enter.",
      },
      {
        title: "Try a reference",
        detail: "Put 10 in A1, 5 in A2, then type =A1+A2 in A3.",
        formula: "=A1+A2",
      },
      {
        title: "Try a function",
        detail: "The grid understands SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, ROUND, ABS and CONCAT.",
        formula: "=SUM(A1:A10)",
      },
      {
        title: "Try a decision",
        detail: "",
        formula: "=IF(A1>A2;\"bigger\";\"smaller\")",
      },
    ],
    practice: {
      task: "In the grid below, build a mini invoice: quantities in B, prices in C, line totals in D and a grand total.",
      answer: "D1: =B1*C1 copied down, then =SUM(D1:D5) for the total.",
    },
    quiz: [
      {
        q: "What must a formula begin with?",
        options: ["=", "+", "@", "#"],
        answer: 0,
        explain: "The equals sign tells Excel to calculate.",
      },
    ],
  },
  {
    slug: "practice-data-worksheet",
    title: "Practice Data Worksheet - a dataset to train on",
    track: "automation",
    level: "Beginner",
    summary:
      "A realistic 20-row sales dataset plus twelve graded exercises that use nearly every function on this site.",
    objectives: [
      "Work with a realistic dataset",
      "Apply lookups, conditional maths, text and date functions",
      "Check your answers against the model solutions",
    ],
    steps: [
      {
        title: "Copy the dataset into the grid",
        detail:
          "Use the practice grid at the bottom of the page, or copy the table below into your own Excel and press Ctrl + T.",
      },
      { title: "1. Total revenue", detail: "", formula: "=SUM(D2:D21)" },
      { title: "2. Average order value, excluding zeros", detail: "", formula: "=AVERAGEIF(D2:D21;\"<>0\")" },
      { title: "3. Number of orders from the North", detail: "", formula: "=COUNTIF(C2:C21;\"North\")" },
      { title: "4. Revenue from the North in Q1", detail: "", formula: "=SUMIFS(D2:D21;C2:C21;\"North\";A2:A21;\"<\"&DATE(2024;4;1))" },
      { title: "5. Biggest single order", detail: "", formula: "=MAX(D2:D21)" },
      { title: "6. Which customer placed it", detail: "", formula: "=INDEX(B2:B21;MATCH(MAX(D2:D21);D2:D21;0))" },
      { title: "7. A live list of unique customers", detail: "", formula: "=SORT(UNIQUE(B2:B21))" },
      { title: "8. Look up a customer's region", detail: "", formula: "=XLOOKUP(G2;B2:B21;C2:C21;\"Not found\")" },
      { title: "9. Order month name", detail: "", formula: "=TEXT(A2;\"mmmm\")" },
      { title: "10. Flag orders over 5 000", detail: "", formula: "=IF(D2>5000;\"Large\";\"Standard\")" },
      { title: "11. Every open order over 1 000", detail: "", formula: "=FILTER(A2:E21;(E2:E21=\"Open\")*(D2:D21>1000);\"None\")" },
      { title: "12. Customer initials", detail: "", formula: "=LEFT(B2;1)&LEFT(TEXTAFTER(B2;\" \");1)" },
    ],
    example: {
      caption: "Practice dataset - Date | Customer | Region | Amount | Status",
      headers: ["Date", "Customer", "Region", "Amount", "Status"],
      rows: [
        ["05/01/2024", "Acme Ltd", "North", "4 200", "Paid"],
        ["11/01/2024", "Borg AS", "South", "1 150", "Open"],
        ["19/01/2024", "Cirrus PLC", "North", "7 800", "Paid"],
        ["02/02/2024", "Delta GmbH", "East", "980", "Open"],
        ["14/02/2024", "Acme Ltd", "North", "3 300", "Paid"],
        ["27/02/2024", "Everest SA", "West", "6 450", "Overdue"],
        ["04/03/2024", "Borg AS", "South", "2 000", "Paid"],
        ["12/03/2024", "Fjord AS", "North", "12 500", "Open"],
        ["21/03/2024", "Cirrus PLC", "North", "540", "Paid"],
        ["03/04/2024", "Delta GmbH", "East", "8 900", "Overdue"],
      ],
    },
    practice: {
      task: "Build a one-cell summary: 'North: 5 orders, 28 340 total, largest 12 500'.",
      answer:
        "=\"North: \"&COUNTIF(C2:C21;\"North\")&\" orders, \"&TEXT(SUMIF(C2:C21;\"North\";D2:D21);\"#,##0\")&\" total, largest \"&TEXT(MAXIFS(D2:D21;C2:C21;\"North\");\"#,##0\")",
    },
    quiz: [
      {
        q: "Which formula counts only the North orders?",
        options: [
          "=COUNT(C2:C21)",
          "=COUNTIF(C2:C21;\"North\")",
          "=COUNTA(C2:C21)",
          "=SUM(C2:C21)",
        ],
        answer: 1,
        explain: "COUNTIF counts by criterion.",
      },
    ],
  },
];
