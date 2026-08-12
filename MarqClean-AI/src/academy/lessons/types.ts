export type QuizQ = {
  q: string;
  options: string[];
  answer: number;
  explain: string;
};

export type Example = {
  caption?: string;
  headers: string[];
  rows: string[][];
};

export type Lesson = {
  slug: string;
  title: string;
  track: TrackId;
  level: "Beginner" | "Intermediate" | "Advanced";
  summary: string;
  objectives: string[];
  syntax?: string;
  steps: { title: string; detail: string; formula?: string }[];
  example?: Example;
  formulas?: { label: string; code: string }[];
  tips?: string[];
  shortcuts?: [string, string][];
  mistakes?: string[];
  practice?: { task: string; answer: string };
  quiz: QuizQ[];
};

export type TrackId =
  | "foundations"
  | "workbook"
  | "charts"
  | "logic"
  | "lookup"
  | "text"
  | "dates"
  | "data"
  | "arrays"
  | "finance"
  | "automation";

export const TRACKS: {
  id: TrackId;
  name: string;
  icon: string;
  blurb: string;
}[] = [
  {
    id: "foundations",
    name: "Excel Foundations",
    icon: "🧱",
    blurb:
      "The interface, arithmetic and the six aggregation functions every report is built on.",
  },
  {
    id: "workbook",
    name: "Workbook & Worksheet Skills",
    icon: "🗂️",
    blurb:
      "The interface, formatting, named ranges, validation, printing, protection, auditing and what-if analysis.",
  },
  {
    id: "charts",
    name: "Charts & Dashboards",
    icon: "📈",
    blurb:
      "Choosing the right chart, formatting it honestly, combo charts, sparklines and a full interactive dashboard build.",
  },
  {
    id: "logic",
    name: "Logic & Conditional Maths",
    icon: "🧠",
    blurb:
      "IF and the whole *IF / *IFS family - how Excel makes decisions and conditional totals.",
  },
  {
    id: "lookup",
    name: "Lookup & Reference",
    icon: "🔎",
    blurb:
      "Pull data from another table with VLOOKUP, XLOOKUP, HLOOKUP and INDEX/MATCH.",
  },
  {
    id: "text",
    name: "Text & Cleaning",
    icon: "✂️",
    blurb:
      "Split, join, trim, find and reformat messy text that arrives from other systems.",
  },
  {
    id: "dates",
    name: "Dates & Time",
    icon: "📅",
    blurb: "Serial numbers, date building blocks, weekdays and month ends.",
  },
  {
    id: "data",
    name: "Data Tools & Analysis",
    icon: "📊",
    blurb:
      "Sort, filter, remove duplicates, PivotTables, Conditional Formatting and Power Query.",
  },
  {
    id: "arrays",
    name: "Dynamic Arrays (365)",
    icon: "⚡",
    blurb:
      "The modern spill functions: UNIQUE, SORT, FILTER, LET, LAMBDA, stacking and shaping.",
  },
  {
    id: "finance",
    name: "Financial Functions",
    icon: "💷",
    blurb: "PMT, PV, FV, PPMT/IPMT and a full mortgage calculator build.",
  },
  {
    id: "automation",
    name: "Automation & Practice",
    icon: "🤖",
    blurb:
      "Macros, VBA, the cheat sheet and live practice worksheets you can type into.",
  },
];
