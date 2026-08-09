import type { Lesson } from "./types";

export const dates: Lesson[] = [
  {
    slug: "day-month-year",
    title: "DAY, MONTH and YEAR - break a date apart",
    track: "dates",
    level: "Beginner",
    summary:
      "Every Excel date is just a number counting days from 1 January 1900. DAY, MONTH and YEAR pull out the pieces you need for grouping.",
    objectives: [
      "Extract parts of a date",
      "Group data by month and year",
      "Understand date serial numbers",
    ],
    syntax: "=DAY(date)   =MONTH(date)   =YEAR(date)",
    steps: [
      {
        title: "See the number behind the date",
        detail:
          "Type a date and format the cell as General: 04/03/2024 becomes 45355. That is why you can subtract dates.",
      },
      { title: "Pull out the parts", detail: "Three simple functions.", formula: "=YEAR(A2) & \"-\" & MONTH(A2)" },
      {
        title: "Month name instead of number",
        detail: "TEXT gives you the label for a pivot or chart axis.",
        formula: "=TEXT(A2;\"mmmm\")",
      },
      {
        title: "Group by month safely",
        detail:
          "Always include the year, otherwise January 2023 and January 2024 merge. A month-start key sorts correctly too.",
        formula: "=DATE(YEAR(A2);MONTH(A2);1)",
      },
      { title: "Quarter", detail: "", formula: "=\"Q\"&ROUNDUP(MONTH(A2)/3;0)" },
    ],
    practice: {
      task: "Create a sortable 'YYYY-MM' key from the date in A2.",
      answer: "=TEXT(A2;\"yyyy-mm\")",
    },
    quiz: [
      {
        q: "Why can you subtract two dates directly?",
        options: [
          "Excel has a DATEDIFF operator",
          "Dates are stored as serial numbers",
          "Only if they are in the same month",
          "You cannot",
        ],
        answer: 1,
        explain: "A date is a number of days, so B2-A2 gives days between.",
      },
    ],
  },
  {
    slug: "date-function",
    title: "DATE - build a valid date from parts",
    track: "dates",
    level: "Beginner",
    summary:
      "DATE(year;month;day) assembles a real date and, brilliantly, rolls over impossible values like month 13 or day 0.",
    objectives: ["Build dates from columns", "Add months safely", "Rescue text dates from an import"],
    syntax: "=DATE(year; month; day)",
    steps: [
      { title: "Assemble from three columns", detail: "Year in A, month in B, day in C.", formula: "=DATE(A2;B2;C2)" },
      {
        title: "Roll-over is a feature",
        detail: "DATE(2024;13;1) returns 1 Jan 2025 and DATE(2024;3;0) returns 29 Feb 2024 (last day of the previous month).",
      },
      { title: "Add months", detail: "Same day, six months later - with automatic year roll-over.", formula: "=DATE(YEAR(A2);MONTH(A2)+6;DAY(A2))" },
      { title: "First and last day of the month", detail: "", formula: "=DATE(YEAR(A2);MONTH(A2);1) and =DATE(YEAR(A2);MONTH(A2)+1;0)" },
      {
        title: "Fix an imported text date like 20240304",
        detail: "Slice the number with LEFT/MID/RIGHT and feed the pieces to DATE.",
        formula: "=DATE(LEFT(A2;4);MID(A2;5;2);RIGHT(A2;2))",
      },
    ],
    practice: {
      task: "A2 holds the text 31122024 (ddmmyyyy). Convert it into a real date.",
      answer: "=DATE(RIGHT(A2;4);MID(A2;3;2);LEFT(A2;2))",
    },
    quiz: [
      {
        q: "=DATE(2024;3;0) returns:",
        options: ["#VALUE!", "29 Feb 2024", "1 Mar 2024", "0"],
        answer: 1,
        explain: "Day 0 of March is the last day of February.",
      },
    ],
  },
  {
    slug: "weekday",
    title: "WEEKDAY - which day of the week is it?",
    track: "dates",
    level: "Beginner",
    summary:
      "WEEKDAY returns a number for the day of the week. Choosing the right return_type is what makes weekend logic work.",
    objectives: ["Return the day number", "Flag weekends", "Highlight weekends automatically"],
    syntax: "=WEEKDAY(date; [return_type])",
    steps: [
      {
        title: "Pick the numbering system",
        detail:
          "Type 1 (default): Sunday = 1. Type 2: Monday = 1 - this is the one you want in most of Europe. Type 3: Monday = 0.",
        formula: "=WEEKDAY(A2;2)",
      },
      { title: "Flag weekends", detail: "With type 2, Saturday = 6 and Sunday = 7.", formula: "=IF(WEEKDAY(A2;2)>5;\"Weekend\";\"Weekday\")" },
      { title: "Day name", detail: "Easier to read than a number.", formula: "=TEXT(A2;\"dddd\")" },
      {
        title: "Highlight weekend rows",
        detail:
          "Conditional Formatting → Use a formula → =WEEKDAY($A2;2)>5 → grey fill. Applies to the whole row.",
      },
      {
        title: "Working days between two dates",
        detail: "NETWORKDAYS counts weekdays and can exclude a holiday list.",
        formula: "=NETWORKDAYS(A2;B2;Holidays)",
      },
    ],
    practice: {
      task: "Return 'Mon'-'Sun' for the date in A2 and mark Saturdays and Sundays.",
      answer: "=TEXT(A2;\"ddd\")&IF(WEEKDAY(A2;2)>5;\" (weekend)\";\"\")",
    },
    quiz: [
      {
        q: "With return_type 2, what number is Sunday?",
        options: ["1", "6", "7", "0"],
        answer: 2,
        explain: "Type 2 starts the week on Monday = 1, so Sunday is 7.",
      },
    ],
  },
  {
    slug: "eomonth",
    title: "EOMONTH - month ends and clean date arithmetic",
    track: "dates",
    level: "Intermediate",
    summary:
      "EOMONTH returns the last day of a month a chosen number of months away - the backbone of payment terms and reporting periods.",
    objectives: ["Find month ends and starts", "Build payment due dates", "Generate a month-end series"],
    syntax: "=EOMONTH(start_date; months)",
    steps: [
      { title: "End of this month", detail: "0 months away.", formula: "=EOMONTH(A2;0)" },
      { title: "End of next month", detail: "", formula: "=EOMONTH(A2;1)" },
      { title: "Start of this month", detail: "Last day of the previous month, plus one.", formula: "=EOMONTH(A2;-1)+1" },
      {
        title: "Payment terms: end of month + 30 days",
        detail: "The classic invoice rule in one formula.",
        formula: "=EOMONTH(A2;0)+30",
      },
      {
        title: "Same day next month, safely",
        detail:
          "EDATE keeps the day number and never invents 31 February.",
        formula: "=EDATE(A2;1)",
      },
      {
        title: "A list of 12 month ends",
        detail: "In Microsoft 365, SEQUENCE spills a whole year at once.",
        formula: "=EOMONTH(TODAY();SEQUENCE(12;1;0))",
      },
    ],
    practice: {
      task: "An invoice dated A2 is due at the end of the following month. Give the due date.",
      answer: "=EOMONTH(A2;1)",
    },
    quiz: [
      {
        q: "=EOMONTH(\"15/01/2024\";-1)+1 returns:",
        options: ["01/01/2024", "31/12/2023", "15/12/2023", "31/01/2024"],
        answer: 0,
        explain: "End of December plus one day = 1 January - the start of the month.",
      },
    ],
  },
];
