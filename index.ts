import { foundations } from "./foundations";
import { workbook } from "./workbook";
import { charts } from "./charts";
import { logic } from "./logic";
import { lookup } from "./lookup";
import { text } from "./text";
import { dates } from "./dates";
import { dataTools } from "./data";
import { arrays } from "./arrays";
import { finance } from "./finance";
import { automation } from "./automation";
import { TRACKS, type Lesson, type TrackId } from "./types";

export { TRACKS };
export type { Lesson, TrackId };

export const LESSONS: Lesson[] = [
  ...foundations,
  ...workbook,
  ...charts,
  ...logic,
  ...lookup,
  ...text,
  ...dates,
  ...dataTools,
  ...arrays,
  ...finance,
  ...automation,
];

export const bySlug = (slug: string) => LESSONS.find((l) => l.slug === slug);

export const byTrack = (track: TrackId) =>
  LESSONS.filter((l) => l.track === track);

export function neighbours(slug: string) {
  const i = LESSONS.findIndex((l) => l.slug === slug);
  return {
    prev: i > 0 ? LESSONS[i - 1] : null,
    next: i >= 0 && i < LESSONS.length - 1 ? LESSONS[i + 1] : null,
    index: i,
  };
}

/** Quick "shortcut to solution" index used by the search palette. */
export const SOLUTIONS: { q: string; slug: string; answer: string }[] = [
  { q: "Add up a column", slug: "sum", answer: "=SUM(B2:B100) or press Alt + =" },
  { q: "Total only visible / filtered rows", slug: "filter-tool", answer: "=SUBTOTAL(109;B2:B100)" },
  { q: "Remove duplicates from a list", slug: "unique", answer: "=SORT(UNIQUE(A2:A500)) - live and non-destructive" },
  { q: "Look up a value in another table", slug: "xlookup", answer: "=XLOOKUP(A2;Ids;Values;\"Not found\")" },
  { q: "Look up to the left", slug: "index-match", answer: "=INDEX(C:C;MATCH(A2;E:E;0))" },
  { q: "Fix #N/A in a lookup", slug: "iferror", answer: "=IFNA(XLOOKUP(...);\"Not found\") and TRIM the lookup value" },
  { q: "Stop #DIV/0!", slug: "division", answer: "=IFERROR(B2/C2;\"-\")" },
  { q: "Total by category", slug: "sumif", answer: "=SUMIF(C:C;\"North\";B:B)" },
  { q: "Total with two or more conditions", slug: "sumifs", answer: "=SUMIFS(B:B;C:C;\"North\";D:D;\"2024\")" },
  { q: "Count how many times a value appears", slug: "countif", answer: "=COUNTIF(A:A;A2)" },
  { q: "Split first and last name", slug: "textbefore-textafter", answer: "=TEXTBEFORE(A2;\" \") and =TEXTAFTER(A2;\" \")" },
  { q: "Split a column by comma", slug: "text-to-columns", answer: "Data → Text to Columns → Delimited, or =TEXTSPLIT(A2;\",\")" },
  { q: "Join text from many cells", slug: "textjoin", answer: "=TEXTJOIN(\", \";TRUE;A2:A50)" },
  { q: "Remove extra spaces breaking a lookup", slug: "trim", answer: "=TRIM(A2) - check with =LEN(A2)-LEN(TRIM(A2))" },
  { q: "Numbers stored as text / SUM returns 0", slug: "value", answer: "Select the column → Data → Text to Columns → Finish, or =VALUE(A2)" },
  { q: "Highlight an entire row by a condition", slug: "conditional-formatting", answer: "Conditional Formatting → formula → =$D2=\"Overdue\"" },
  { q: "Fill every empty cell in imported data", slug: "flash-fill", answer: "F5 → Special → Blanks → type = ↑ → Ctrl + Enter" },
  { q: "Summarise thousands of rows", slug: "pivot-table", answer: "Alt + N + V → drag fields into Rows / Values" },
  { q: "Combine 12 monthly files", slug: "power-query-combine-workbooks", answer: "Data → Get Data → From Folder → Combine Files" },
  { q: "Turn months-across into tidy data", slug: "unpivot-columns-power-query", answer: "Power Query → select keys → Unpivot Other Columns" },
  { q: "Filter with a formula", slug: "filter-function", answer: "=FILTER(A2:D500;C2:C500=\"North\";\"None\")" },
  { q: "Calculate a loan payment", slug: "pmt", answer: "=-PMT(rate/12;years*12;loan)" },
  { q: "Last day of the month", slug: "eomonth", answer: "=EOMONTH(A2;0)" },
  { q: "Working out days between dates", slug: "day-month-year", answer: "=B2-A2 (or NETWORKDAYS for weekdays only)" },
  { q: "Unhide all sheets at once", slug: "macros", answer: "Alt + F11 → paste the UnhideAllSheets macro → F5" },
  { q: "Run a VBA macro someone sent me", slug: "run-vba-macro", answer: "Alt + F11 → Insert Module → paste → F5 → save as .xlsm" },
  { q: "Make my own function", slug: "lambda", answer: "Name Manager → =LAMBDA(x;…)" },
  { q: "Show the formula instead of the result", slug: "formulatext", answer: "=FORMULATEXT(B2) or press Ctrl + `" },
];
