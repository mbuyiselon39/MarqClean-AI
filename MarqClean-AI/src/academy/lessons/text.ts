import type { Lesson } from "./types";

export const text: Lesson[] = [
  {
    slug: "left-mid-right",
    title: "LEFT, MID and RIGHT - cut text by position",
    track: "text",
    level: "Beginner",
    summary:
      "Extract a fixed number of characters from the start, middle or end of a string - the base of every code-splitting formula.",
    objectives: [
      "Extract fixed-length parts of a code",
      "Combine with FIND for variable-length text",
    ],
    syntax: "=LEFT(text; n)   =RIGHT(text; n)   =MID(text; start; n)",
    steps: [
      { title: "LEFT - from the start", detail: "First 3 characters of a product code.", formula: "=LEFT(A2;3)" },
      { title: "RIGHT - from the end", detail: "Last 4 digits of an account number.", formula: "=RIGHT(A2;4)" },
      { title: "MID - from a position", detail: "Start at character 5 and take 2.", formula: "=MID(A2;5;2)" },
      {
        title: "Variable length: cut at a separator",
        detail: "FIND tells you where the dash is; take everything before it.",
        formula: "=LEFT(A2;FIND(\"-\";A2)-1)",
      },
      {
        title: "Everything after the separator",
        detail: "LEN minus the position of the separator gives the remaining characters.",
        formula: "=RIGHT(A2;LEN(A2)-FIND(\"-\";A2))",
      },
      {
        title: "In Microsoft 365, stop counting",
        detail: "TEXTBEFORE and TEXTAFTER do the same thing with no arithmetic at all.",
        formula: "=TEXTBEFORE(A2;\"-\")",
      },
    ],
    example: {
      caption: "Code ORD-2024-0091",
      headers: ["Formula", "Result"],
      rows: [
        ["=LEFT(A2;3)", "ORD"],
        ["=MID(A2;5;4)", "2024"],
        ["=RIGHT(A2;4)", "0091"],
      ],
    },
    tips: [
      "LEFT/RIGHT/MID always return text. Wrap in VALUE if you need to calculate with the result.",
      "RIGHT with more characters than the string has simply returns the whole string - no error.",
    ],
    practice: {
      task: "Cell A2 holds 'GB29-NWBK-6016'. Return the middle block.",
      answer: "=MID(A2;6;4)  → NWBK",
    },
    quiz: [
      {
        q: "=LEFT(\"Excel\";10) returns:",
        options: ["#VALUE!", "Excel", "Exc", "Empty"],
        answer: 1,
        explain: "Asking for more characters than exist just returns everything.",
      },
    ],
  },
  {
    slug: "textbefore-textafter",
    title: "TEXTBEFORE & TEXTAFTER - split without counting",
    track: "text",
    level: "Intermediate",
    summary:
      "Two Microsoft 365 functions that grab everything before or after a delimiter, including the nth occurrence.",
    objectives: [
      "Replace LEFT/FIND and RIGHT/LEN/FIND",
      "Use the instance argument and search from the end",
    ],
    syntax:
      "=TEXTBEFORE(text; delimiter; [instance]; [match_mode]; [match_end]; [if_not_found])",
    steps: [
      { title: "Everything before the first space", detail: "First name from a full name.", formula: "=TEXTBEFORE(A2;\" \")" },
      { title: "Everything after the @", detail: "Email domain.", formula: "=TEXTAFTER(A2;\"@\")" },
      {
        title: "The nth occurrence",
        detail: "Third argument = which delimiter to use. 2 means the second one.",
        formula: "=TEXTBEFORE(A2;\"-\";2)",
      },
      {
        title: "Count from the end",
        detail: "A negative instance searches backwards - perfect for file extensions.",
        formula: "=TEXTAFTER(A2;\".\";-1)",
      },
      {
        title: "Safe fallback",
        detail: "If the delimiter is missing you normally get #N/A; supply if_not_found instead.",
        formula: "=TEXTBEFORE(A2;\"-\";1;0;0;A2)",
      },
    ],
    practice: {
      task: "A2 = 'north/uk/london'. Return 'uk'.",
      answer: "=TEXTBEFORE(TEXTAFTER(A2;\"/\");\"/\")  - or =TEXTAFTER(TEXTBEFORE(A2;\"/\";2);\"/\")",
    },
    quiz: [
      {
        q: "How do you get the text after the LAST dot?",
        options: [
          "=TEXTAFTER(A2;\".\";0)",
          "=TEXTAFTER(A2;\".\";-1)",
          "=TEXTAFTER(A2;\".\";99)",
          "=RIGHT(A2;3)",
        ],
        answer: 1,
        explain: "A negative instance number counts occurrences from the end.",
      },
    ],
  },
  {
    slug: "textsplit",
    title: "TEXTSPLIT - turn one cell into many",
    track: "text",
    level: "Intermediate",
    summary:
      "TEXTSPLIT is a formula version of Text to Columns: it spills a delimited string across columns and/or down rows.",
    objectives: ["Split across columns", "Split into rows", "Handle empty values and multiple delimiters"],
    syntax:
      "=TEXTSPLIT(text; col_delimiter; [row_delimiter]; [ignore_empty]; [match_mode]; [pad_with])",
    steps: [
      { title: "Split across columns", detail: "Comma-separated values spill to the right.", formula: "=TEXTSPLIT(A2;\",\")" },
      { title: "Split down rows", detail: "Leave the column delimiter empty and set the row one.", formula: "=TEXTSPLIT(A2;;\";\")" },
      { title: "Both at once", detail: "Rebuild a whole grid from one pasted blob.", formula: "=TEXTSPLIT(A2;\",\";CHAR(10))" },
      {
        title: "Several possible delimiters",
        detail: "Pass an array in curly braces.",
        formula: "=TEXTSPLIT(A2;{\",\";\";\";\" \"};;TRUE)",
      },
      { title: "Ignore blanks", detail: "The fourth argument removes empty results caused by double delimiters.", formula: "=TEXTSPLIT(A2;\",\";;TRUE)" },
    ],
    tips: [
      "TEXTSPLIT is live: change the source cell and the split updates. Text to Columns is one-off.",
      "Not available before Microsoft 365 - use Text to Columns or Power Query there.",
    ],
    practice: {
      task: "A2 = 'red, green, , blue'. Split into columns without the empty item.",
      answer: "=TEXTSPLIT(A2;\", \";;TRUE)",
    },
    quiz: [
      {
        q: "Which argument makes TEXTSPLIT split into rows?",
        options: ["1st", "2nd", "3rd", "4th"],
        answer: 2,
        explain: "row_delimiter is the third argument.",
      },
    ],
  },
  {
    slug: "concatenate",
    title: "Concatenate - join text with & and CONCATENATE",
    track: "text",
    level: "Beginner",
    summary:
      "Joining text is the easiest thing in Excel once you know that & glues values and \" \" adds the space.",
    objectives: ["Join cells with &", "Insert separators and line breaks", "Keep numbers formatted while joining"],
    syntax: "=A2 & \" \" & B2      =CONCATENATE(A2;\" \";B2)",
    steps: [
      { title: "The & operator", detail: "The fastest way - first name, space, last name.", formula: "=A2&\" \"&B2" },
      { title: "Add literal text", detail: "Anything in quotes is printed as-is.", formula: "=\"Dear \"&A2&\",\"" },
      {
        title: "Line break inside a cell",
        detail: "CHAR(10) is a newline. Turn on Wrap Text or you will not see it.",
        formula: "=A2&CHAR(10)&B2",
      },
      {
        title: "Keep the number format",
        detail: "Joining a date gives you the ugly serial number unless you wrap it in TEXT.",
        formula: "=\"Due \"&TEXT(B2;\"dd mmm yyyy\")",
      },
      {
        title: "CONCATENATE is retired",
        detail: "It still works but Microsoft replaced it with CONCAT and TEXTJOIN.",
      },
    ],
    practice: {
      task: "Build 'SMITH, John (Sales)' from surname A2, first name B2, department C2.",
      answer: "=UPPER(A2)&\", \"&B2&\" (\"&C2&\")\"",
    },
    quiz: [
      {
        q: "Why wrap a date in TEXT() when joining?",
        options: [
          "To avoid errors",
          "Otherwise the serial number appears",
          "To make it bold",
          "It is not needed",
        ],
        answer: 1,
        explain: "Concatenation strips formatting, leaving the underlying serial number such as 45678.",
      },
    ],
  },
  {
    slug: "textjoin",
    title: "TEXTJOIN - join a whole range with a delimiter",
    track: "text",
    level: "Intermediate",
    summary:
      "TEXTJOIN adds a separator between every item and can skip blanks - the difference between a clean list and 'a,,b,,c'.",
    objectives: ["Join a range", "Skip empty cells", "Join only the rows that match a condition"],
    syntax: "=TEXTJOIN(delimiter; ignore_empty; text1; [text2]; …)",
    steps: [
      { title: "Join a column", detail: "Comma and space between every value.", formula: "=TEXTJOIN(\", \";TRUE;A2:A50)" },
      { title: "Keep the blanks", detail: "Set the second argument to FALSE when position matters.", formula: "=TEXTJOIN(\"|\";FALSE;A2:A50)" },
      {
        title: "Conditional join",
        detail: "FILTER first, then join - a one-cell 'who is in this group' list.",
        formula: "=TEXTJOIN(\", \";TRUE;FILTER(A2:A50;B2:B50=\"Oslo\"))",
      },
      {
        title: "Older-Excel conditional join",
        detail: "IF returns an array of names or blanks; TEXTJOIN drops the blanks.",
        formula: "=TEXTJOIN(\", \";TRUE;IF(B2:B50=\"Oslo\";A2:A50;\"\"))",
      },
      {
        title: "Build an SQL IN clause",
        detail: "Four quotes produce one literal quote character.",
        formula: "=\"IN (\"&TEXTJOIN(\",\";TRUE;\"\"\"\"&A2:A50&\"\"\"\")&\")\"",
      },
    ],
    practice: {
      task: "Join all email addresses in C2:C200 with a semicolon, skipping empties.",
      answer: "=TEXTJOIN(\"; \";TRUE;C2:C200)",
    },
    quiz: [
      {
        q: "What does the second argument of TEXTJOIN control?",
        options: [
          "The delimiter",
          "Whether empty cells are skipped",
          "Case sensitivity",
          "The maximum length",
        ],
        answer: 1,
        explain: "TRUE = ignore empty cells.",
      },
    ],
  },
  {
    slug: "concat",
    title: "CONCAT - join ranges without a delimiter",
    track: "text",
    level: "Beginner",
    summary:
      "CONCAT replaced CONCATENATE and accepts whole ranges, but it puts nothing between the items.",
    objectives: ["Join a range with no separator", "Choose between CONCAT, TEXTJOIN and &"],
    syntax: "=CONCAT(text1; [text2]; …)",
    steps: [
      { title: "Join a range", detail: "Every value in the range end to end.", formula: "=CONCAT(A2:D2)" },
      {
        title: "Choosing the right one",
        detail:
          "& for two or three cells. CONCAT when you need no separator across a range. TEXTJOIN when you need a separator or want blanks skipped.",
      },
      {
        title: "Build a binary/flag string",
        detail: "CONCAT over a row of Y/N flags creates a compact signature you can compare or count.",
        formula: "=CONCAT(IF(B2:H2=\"Y\";\"1\";\"0\"))",
      },
    ],
    practice: {
      task: "Join the characters in A2:A6 into one word.",
      answer: "=CONCAT(A2:A6)",
    },
    quiz: [
      {
        q: "The difference between CONCAT and TEXTJOIN is:",
        options: [
          "CONCAT is older",
          "TEXTJOIN adds a delimiter and can skip blanks",
          "CONCAT only takes 2 arguments",
          "There is none",
        ],
        answer: 1,
        explain: "CONCAT simply glues everything together.",
      },
    ],
  },
  {
    slug: "len",
    title: "LEN - count characters",
    track: "text",
    level: "Beginner",
    summary:
      "LEN counts every character including spaces, which makes it the detective tool for invisible data problems.",
    objectives: ["Count characters", "Detect trailing spaces", "Validate code lengths"],
    syntax: "=LEN(text)",
    steps: [
      { title: "Count characters", detail: "Spaces count too.", formula: "=LEN(A2)" },
      {
        title: "Find hidden spaces",
        detail: "If LEN and LEN(TRIM()) differ, the cell has extra spaces that will break lookups.",
        formula: "=LEN(A2)-LEN(TRIM(A2))",
      },
      {
        title: "Validate a code length",
        detail: "Flag account numbers that are not exactly 10 characters.",
        formula: "=IF(LEN(A2)<>10;\"Check\";\"OK\")",
      },
      {
        title: "Count occurrences of a character",
        detail: "Length minus length-without-the-character.",
        formula: "=LEN(A2)-LEN(SUBSTITUTE(A2;\",\";\"\"))",
      },
    ],
    practice: {
      task: "How many words are in A2? (Words = spaces + 1.)",
      answer: "=LEN(TRIM(A2))-LEN(SUBSTITUTE(TRIM(A2);\" \";\"\"))+1",
    },
    quiz: [
      {
        q: "=LEN(\"Excel \") returns:",
        options: ["5", "6", "4", "#VALUE!"],
        answer: 1,
        explain: "The trailing space is a character.",
      },
    ],
  },
  {
    slug: "find",
    title: "FIND - locate text, case-sensitive",
    track: "text",
    level: "Intermediate",
    summary:
      "FIND returns the position of one string inside another. It is case-sensitive and errors when there is no match.",
    objectives: ["Locate a character", "Feed the position into LEFT/MID", "Find the nth occurrence"],
    syntax: "=FIND(find_text; within_text; [start_num])",
    steps: [
      { title: "Position of a character", detail: "Where is the dash?", formula: "=FIND(\"-\";A2)" },
      { title: "Start later in the string", detail: "The third argument skips ahead - this is how you find the 2nd dash.", formula: "=FIND(\"-\";A2;FIND(\"-\";A2)+1)" },
      { title: "Use it to cut", detail: "Everything before the dash.", formula: "=LEFT(A2;FIND(\"-\";A2)-1)" },
      {
        title: "Handle 'not found'",
        detail: "FIND returns #VALUE! when the text is missing. Wrap it.",
        formula: "=IFERROR(FIND(\"-\";A2);0)",
      },
      { title: "Case matters", detail: "FIND(\"a\";\"Apple\") is 0 characters found → error; SEARCH would find it." },
    ],
    practice: {
      task: "A2 = 'Smith, John'. Return the surname.",
      answer: "=LEFT(A2;FIND(\",\";A2)-1)",
    },
    quiz: [
      {
        q: "FIND differs from SEARCH because it:",
        options: [
          "Is case-sensitive and does not allow wildcards",
          "Is faster",
          "Returns text",
          "Works only on numbers",
        ],
        answer: 0,
        explain: "SEARCH is case-insensitive and accepts * and ?.",
      },
    ],
  },
  {
    slug: "search",
    title: "SEARCH - locate text, case-insensitive with wildcards",
    track: "text",
    level: "Intermediate",
    summary:
      "SEARCH is the forgiving twin of FIND: it ignores case and understands * and ?, which makes it perfect for 'does this contain…' tests.",
    objectives: ["Do a contains test", "Use wildcards", "Categorise rows by keyword"],
    syntax: "=SEARCH(find_text; within_text; [start_num])",
    steps: [
      { title: "Contains test", detail: "ISNUMBER turns the position into a clean TRUE/FALSE.", formula: "=ISNUMBER(SEARCH(\"invoice\";A2))" },
      { title: "Wildcards", detail: "? = one character, * = any run of characters.", formula: "=SEARCH(\"IN?-*\";A2)" },
      {
        title: "Categorise with nested IF",
        detail: "The order decides the priority when a row contains two keywords.",
        formula:
          "=IFS(ISNUMBER(SEARCH(\"refund\";A2));\"Refund\";ISNUMBER(SEARCH(\"invoice\";A2));\"Invoice\";TRUE;\"Other\")",
      },
      {
        title: "Highlight matching rows",
        detail:
          "Conditional Formatting → Use a formula → =ISNUMBER(SEARCH($F$1;$A2)) highlights every row containing whatever you type in F1.",
      },
    ],
    practice: {
      task: "Flag any description in A2 that mentions 'urgent', in any capitalisation.",
      answer: "=IF(ISNUMBER(SEARCH(\"urgent\";A2));\"Urgent\";\"\")",
    },
    quiz: [
      {
        q: "=SEARCH(\"a\";\"Banana\") returns:",
        options: ["0", "2", "4", "#VALUE!"],
        answer: 1,
        explain: "The first 'a' (ignoring case) is at position 2.",
      },
    ],
  },
  {
    slug: "substitute",
    title: "SUBSTITUTE - replace text inside a cell",
    track: "text",
    level: "Intermediate",
    summary:
      "SUBSTITUTE swaps one piece of text for another - the formula version of Find & Replace, and the trick behind counting characters.",
    objectives: [
      "Replace text and specific occurrences",
      "Strip unwanted characters",
      "Know when to use REPLACE instead",
    ],
    syntax: "=SUBSTITUTE(text; old_text; new_text; [instance_num])",
    steps: [
      { title: "Replace everything", detail: "Every dash becomes a slash.", formula: "=SUBSTITUTE(A2;\"-\";\"/\")" },
      { title: "Replace only the 2nd occurrence", detail: "The fourth argument picks the instance.", formula: "=SUBSTITUTE(A2;\"-\";\"/\";2)" },
      { title: "Delete characters", detail: "Replace with nothing to strip currency symbols or spaces.", formula: "=SUBSTITUTE(SUBSTITUTE(A2;\"£\";\"\");\",\";\"\")" },
      {
        title: "Remove non-breaking spaces from web data",
        detail: "CHAR(160) is the invisible culprit that TRIM cannot remove.",
        formula: "=TRIM(SUBSTITUTE(A2;CHAR(160);\" \"))",
      },
      {
        title: "SUBSTITUTE vs REPLACE",
        detail:
          "SUBSTITUTE matches by content; REPLACE(text;start;chars;new) works by position. Use REPLACE to mask an account number.",
        formula: "=REPLACE(A2;1;6;\"******\")",
      },
    ],
    practice: {
      task: "Convert '£1,250.00' text in A2 into a real number.",
      answer: "=VALUE(SUBSTITUTE(SUBSTITUTE(A2;\"£\";\"\");\",\";\"\"))",
    },
    quiz: [
      {
        q: "SUBSTITUTE is case-…",
        options: ["insensitive", "sensitive", "dependent on settings", "n/a"],
        answer: 1,
        explain: "SUBSTITUTE is case-sensitive; \"abc\" will not match \"ABC\".",
      },
    ],
  },
  {
    slug: "trim",
    title: "TRIM - remove the spaces that break your lookups",
    track: "text",
    level: "Beginner",
    summary:
      "TRIM deletes leading, trailing and duplicated inner spaces. It fixes more #N/A errors than any other function.",
    objectives: ["Clean imported text", "Combine TRIM with CLEAN and SUBSTITUTE"],
    syntax: "=TRIM(text)",
    steps: [
      { title: "Clean a value", detail: "'  John   Smith ' becomes 'John Smith'.", formula: "=TRIM(A2)" },
      { title: "Clean inside a lookup", detail: "Stops the invisible-space #N/A.", formula: "=XLOOKUP(TRIM(A2);Ids;Names;\"-\")" },
      { title: "Remove line breaks too", detail: "CLEAN removes non-printing characters.", formula: "=TRIM(CLEAN(A2))" },
      {
        title: "The full clean-up combo",
        detail: "Non-breaking spaces from HTML need SUBSTITUTE first.",
        formula: "=TRIM(CLEAN(SUBSTITUTE(A2;CHAR(160);\" \")))",
      },
      {
        title: "Make it permanent",
        detail: "Copy the cleaned column, then Paste Special > Values over the original.",
      },
    ],
    practice: {
      task: "Your VLOOKUP fails on a few rows only. What is the first thing you try?",
      answer: "Compare =LEN(A2) with =LEN(TRIM(A2)) and wrap the lookup value in TRIM.",
    },
    quiz: [
      {
        q: "TRIM removes:",
        options: [
          "All spaces",
          "Leading, trailing and repeated inner spaces",
          "Only leading spaces",
          "Line breaks",
        ],
        answer: 1,
        explain: "Single spaces between words are kept; CLEAN handles line breaks.",
      },
    ],
  },
  {
    slug: "lower-upper-proper",
    title: "LOWER, UPPER and PROPER - fix capitalisation",
    track: "text",
    level: "Beginner",
    summary:
      "Three one-argument functions that standardise names, emails and codes before you import or compare them.",
    objectives: ["Change case", "Know PROPER's limitations", "Apply the change permanently"],
    syntax: "=LOWER(text)   =UPPER(text)   =PROPER(text)",
    steps: [
      { title: "LOWER for emails", detail: "Systems often reject mixed-case keys.", formula: "=LOWER(A2)" },
      { title: "UPPER for codes", detail: "", formula: "=UPPER(A2)" },
      { title: "PROPER for names", detail: "Capitalises the first letter of every word.", formula: "=PROPER(A2)" },
      {
        title: "PROPER's traps",
        detail:
          "'MCDONALD' becomes 'Mcdonald' and 'O'BRIEN' becomes \"O'Brien\" - correct - but 'IBM' becomes 'Ibm'. Always eyeball the result on real data.",
      },
      {
        title: "Bake it in",
        detail: "Copy the helper column → Paste Special > Values → delete the original.",
      },
      { title: "No-formula option", detail: "Flash Fill (Ctrl + E) also learns capitalisation patterns." },
    ],
    practice: {
      task: "Turn 'jane MCBRIDE' in A2 into 'Jane McBride'.",
      answer:
        "=SUBSTITUTE(PROPER(A2);\"Mc\";\"Mc\") won't fix it alone - use PROPER then correct exceptions with SUBSTITUTE(PROPER(A2);\"Mcbride\";\"McBride\").",
    },
    quiz: [
      {
        q: "=PROPER(\"john o'neill\") returns:",
        options: ["JOHN O'NEILL", "John O'Neill", "John o'neill", "#VALUE!"],
        answer: 1,
        explain: "PROPER capitalises after every non-letter, including the apostrophe.",
      },
    ],
  },
  {
    slug: "value",
    title: "VALUE - turn text into a real number",
    track: "text",
    level: "Beginner",
    summary:
      "When numbers arrive as text your SUMs return zero. VALUE converts them back so the maths works again.",
    objectives: ["Convert text to number", "Recognise the symptoms", "Fix a whole column at once"],
    syntax: "=VALUE(text)",
    steps: [
      { title: "Spot the problem", detail: "Numbers stuck on the left of the cell, a green triangle, SUM returning 0." },
      { title: "Convert one cell", detail: "", formula: "=VALUE(A2)" },
      { title: "Strip the junk first", detail: "VALUE fails on symbols and thousands separators it does not recognise.", formula: "=VALUE(SUBSTITUTE(SUBSTITUTE(A2;\" \";\"\");\",\";\"\"))" },
      {
        title: "The lazy alternatives",
        detail: "=A2*1 or =A2+0 or =--A2 all coerce text to a number just as well.",
        formula: "=--A2",
      },
      {
        title: "Fix a whole column in 5 seconds",
        detail:
          "Select the column → Data → Text to Columns → Next → Next → Finish. Excel re-parses every cell as a number.",
      },
      { title: "Dates too", detail: "DATEVALUE converts a text date; TIMEVALUE converts a text time." },
    ],
    practice: {
      task: "A2 contains the text \"1 250,50\" (space thousands, comma decimal). Convert it.",
      answer: "=VALUE(SUBSTITUTE(SUBSTITUTE(A2;\" \";\"\");\",\";\".\"))",
    },
    quiz: [
      {
        q: "Fastest fix for a whole column of text-numbers:",
        options: [
          "Retype them",
          "Text to Columns → Finish",
          "Bold them",
          "Sort the column",
        ],
        answer: 1,
        explain: "Text to Columns re-parses every value with no helper column.",
      },
    ],
  },
  {
    slug: "text-function",
    title: "TEXT - format a number as text you control",
    track: "text",
    level: "Intermediate",
    summary:
      "TEXT applies a number-format code and returns text, which is how you build readable sentences and titles from data.",
    objectives: ["Use format codes", "Build dynamic report titles", "Keep leading zeros"],
    syntax: "=TEXT(value; format_text)",
    steps: [
      { title: "Format a date", detail: "", formula: "=TEXT(A2;\"dd mmm yyyy\")" },
      { title: "Format currency", detail: "", formula: "=TEXT(A2;\"£#,##0.00\")" },
      { title: "Keep leading zeros", detail: "Codes and postcodes stay 6 digits.", formula: "=TEXT(A2;\"000000\")" },
      { title: "Percentages", detail: "", formula: "=TEXT(A2;\"0.0%\")" },
      {
        title: "Dynamic report title",
        detail: "Join TEXT results into a sentence that updates itself.",
        formula: "=\"Sales report - \"&TEXT(TODAY();\"mmmm yyyy\")&\": \"&TEXT(B1;\"£#,##0\")",
      },
      {
        title: "Remember",
        detail:
          "The result is TEXT. Do not use TEXT on numbers you still need to calculate with - use a cell number format instead (Ctrl + 1).",
      },
    ],
    example: {
      headers: ["Code", "Value 0.4567", "Result"],
      rows: [
        ["0.0%", "", "45.7%"],
        ["#,##0", "1234567", "1,234,567"],
        ["ddd dd/mm", "date", "Mon 04/03"],
        ["0\" kg\"", "12", "12 kg"],
      ],
    },
    practice: {
      task: "Show 'Week 12 of 2024' from a date in A2.",
      answer: "=\"Week \"&TEXT(A2;\"ww\")&\" of \"&TEXT(A2;\"yyyy\")",
    },
    quiz: [
      {
        q: "What does TEXT return?",
        options: ["A number", "Text", "A date serial", "An error"],
        answer: 1,
        explain: "Always text - which is why it can no longer be summed.",
      },
    ],
  },
];
