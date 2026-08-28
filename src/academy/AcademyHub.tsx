import React, { useState } from "react";
import {
  GraduationCap,
  CheckCircle,
  Copy,
  Check,
  Calculator,
  Search,
  Code,
  Lightbulb,
  ArrowLeft,
} from "lucide-react";
import { CopyDataButton } from "../components/CopyDataButton";

interface AcademyHubProps {
  onExit?: () => void;
}

interface CourseModule {
  id: string;
  category: string;
  title: string;
  level: "Beginner" | "Intermediate" | "Advanced" | "Mastery";
  description: string;
  syntax: string;
  example: string;
  outputPreview: string;
  tips: string[];
}

const ACADEMY_COURSES: CourseModule[] = [
  {
    id: "xlookup-mastery",
    category: "Lookup & Reference",
    title: "XLOOKUP Modern Exact & Approximate Matching",
    level: "Intermediate",
    description:
      "Supersedes VLOOKUP and HLOOKUP with bi-directional scanning, native exact matching by default, and robust fallback error values.",
    syntax: "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
    example: '=XLOOKUP(A2, Ledger!A:A, Ledger!E:E, "Not Found", 0)',
    outputPreview: "Returns matching cell from column E without counting column indexes or breaking on column insertions.",
    tips: [
      "No more column indexing (1, 2, 3...) — direct column reference.",
      "Built-in error handling via the 4th parameter `[if_not_found]`.",
      "Works from right-to-left just as easily as left-to-right.",
    ],
  },
  {
    id: "dynamic-filter",
    category: "Dynamic Array Formulas",
    title: "FILTER & UNIQUE Dynamic Array Spills",
    level: "Advanced",
    description:
      "Extract subsets of table records that meet multiple criteria and automatically spill clean lists without copying or manual drag-fills.",
    syntax: "=FILTER(array, include, [if_empty])",
    example: '=FILTER(A2:D500, (B2:B500="EMEA") * (D2:D500>10000), "No Records")',
    outputPreview: "Instantly spills 42 matching client rows across columns A through D with dynamic live updating.",
    tips: [
      "Use `*` for AND conditions, and `+` for OR conditions across columns.",
      "Combine with `=SORT(FILTER(...))` for auto-alphabetized results.",
      "Ensure spill cells below and to the right are empty to avoid #SPILL! errors.",
    ],
  },
  {
    id: "subtotal-grouping",
    category: "Financial Consolidation",
    title: "SUBTOTAL 109 Visible Sum for Filtered Client Funds",
    level: "Intermediate",
    description:
      "Calculate aggregated sum totals that dynamically ignore hidden or filtered rows, essential for client fund reconciliations.",
    syntax: "=SUBTOTAL(function_num, ref1, [ref2], ...)",
    example: "=SUBTOTAL(109, D2:D250)",
    outputPreview: "Evaluates only currently filtered rows, excluding manually hidden or filtered-out entries.",
    tips: [
      "Function 109 = SUM excluding hidden rows.",
      "Function 101 = AVERAGE excluding hidden rows.",
      "Function 103 = COUNTA excluding hidden rows.",
    ],
  },
  {
    id: "let-lambda-optimization",
    category: "Calculation Efficiency",
    title: "LET Function Variable Definitions & Performance",
    level: "Mastery",
    description:
      "Assign friendly names to intermediate calculation steps, drastically eliminating redundant sub-calculations in massive enterprise workbooks.",
    syntax: "=LET(name1, name_value1, [name2, name_value2], ..., calculation)",
    example: "=LET(gross, C2:C1000, tax, gross * 0.15, gross - tax)",
    outputPreview: "Calculates net balance 8x faster by storing the intermediate array evaluation once in memory.",
    tips: [
      "Improves readability of complex nested IF / SWITCH formulas.",
      "Drastically cuts workbook recalculation latency on 50,000+ rows.",
      "Define up to 126 variables in a single LET statement.",
    ],
  },
  {
    id: "textsplit-cleaner",
    category: "Data Sanitization",
    title: "TEXTSPLIT & TEXTJOIN Modern Text Normalization",
    level: "Intermediate",
    description:
      "Split delimited string exports by commas, tabs, or line breaks into dynamic columns or rows without legacy Text-to-Columns wizards.",
    syntax: "=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])",
    example: '=TEXTSPLIT(A2, ", ")',
    outputPreview: "Spills comma-separated tokens across adjacent columns instantly.",
    tips: [
      "Can split both by columns AND rows in a single formula pass.",
      "Combine with TRIM to eliminate unwanted leading or trailing whitespaces.",
    ],
  },
];

export const AcademyHub: React.FC<AcademyHubProps> = ({ onExit }) => {
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [activeCourseId, setActiveCourseId] = useState<string>("xlookup-mastery");
  const [copiedId, setCopiedId] = useState<string | null>(null);

  const categories = ["all", "Lookup & Reference", "Dynamic Array Formulas", "Financial Consolidation", "Calculation Efficiency", "Data Sanitization"];

  const filteredCourses = ACADEMY_COURSES.filter((course) => {
    const matchesCat = selectedCategory === "all" || course.category === selectedCategory;
    const matchesSearch =
      course.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      course.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
      course.syntax.toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCat && matchesSearch;
  });

  const activeCourse = ACADEMY_COURSES.find((c) => c.id === activeCourseId) || ACADEMY_COURSES[0];

  const handleCopyFormula = (formula: string, id: string) => {
    navigator.clipboard.writeText(formula);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  return (
    <div className="space-y-8">
      {/* Header Banner */}
      <div className="rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <GraduationCap className="h-3.5 w-3.5" />
              MarqClean Academy &bull; Enterprise Formula Hub
            </div>
            <h1 className="mt-2 text-2xl font-extrabold text-white sm:text-3xl">
              Modern Excel Automation &amp; Master Formula Engineering
            </h1>
            <p className="mt-1.5 text-xs sm:text-sm text-slate-400 max-w-2xl">
              Comprehensive interactive guides, syntaxes, best practices, and copy-ready formula patterns for high-performance financial data modeling.
            </p>
          </div>

          <div className="flex items-center gap-2">
            {onExit ? (
              <button
                onClick={onExit}
                className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-900 px-3.5 py-1.5 text-xs font-medium text-slate-300 transition hover:border-slate-500 hover:text-white"
              >
                <ArrowLeft className="h-3.5 w-3.5" />
                Back to Tools
              </button>
            ) : null}
            <CopyDataButton
              payload={{
                type: "records",
                data: ACADEMY_COURSES.map((c) => ({
                  title: c.title,
                  category: c.category,
                  level: c.level,
                  description: c.description,
                  syntax: c.syntax,
                  example: c.example,
                  outputPreview: c.outputPreview,
                })),
              }}
              showJson={true}
              showCsv={true}
              showTsv={true}
              size="sm"
              buttonTheme="cyan"
            />
          </div>
        </div>

        {/* Filter bar */}
        <div className="mt-6 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-2">
            {categories.map((cat) => (
              <button
                key={cat}
                onClick={() => setSelectedCategory(cat)}
                className={`rounded-xl px-3 py-1.5 text-xs font-semibold capitalize transition ${
                  selectedCategory === cat
                    ? "bg-cyan-500 text-slate-950 font-bold shadow-md shadow-cyan-500/20"
                    : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
                }`}
              >
                {cat === "all" ? "All Curriculum" : cat}
              </button>
            ))}
          </div>

          <div className="relative w-full sm:w-64">
            <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400" />
            <input
              type="text"
              placeholder="Search formulas &amp; functions..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full rounded-xl border border-slate-800 bg-slate-900/80 py-1.5 pl-9 pr-3 text-xs text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* Main 2-column curriculum view */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left column: course list */}
        <div className="lg:col-span-5 space-y-3">
          <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 px-1">
            Formulas &amp; Architectural Patterns ({filteredCourses.length})
          </h3>
          <div className="space-y-2.5">
            {filteredCourses.map((course) => {
              const isSelected = course.id === activeCourseId;
              return (
                <div
                  key={course.id}
                  onClick={() => setActiveCourseId(course.id)}
                  className={`p-4 rounded-2xl border transition cursor-pointer ${
                    isSelected
                      ? "border-cyan-400 bg-cyan-950/40 shadow-lg shadow-cyan-500/10"
                      : "border-slate-800 bg-slate-950/60 hover:border-slate-700 hover:bg-slate-900/60"
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[10px] font-mono font-bold uppercase px-2 py-0.5 rounded-full bg-slate-800 text-cyan-300">
                      {course.category}
                    </span>
                    <span
                      className={`text-[10px] font-semibold px-2 py-0.5 rounded-full ${
                        course.level === "Mastery"
                          ? "bg-purple-950 text-purple-300 border border-purple-800"
                          : course.level === "Advanced"
                          ? "bg-blue-950 text-blue-300 border border-blue-800"
                          : "bg-emerald-950 text-emerald-300 border border-emerald-800"
                      }`}
                    >
                      {course.level}
                    </span>
                  </div>

                  <h4 className="mt-2 text-sm font-bold text-white group-hover:text-cyan-300 transition">
                    {course.title}
                  </h4>
                  <p className="mt-1 text-xs text-slate-400 line-clamp-2">{course.description}</p>
                </div>
              );
            })}
          </div>
        </div>

        {/* Right column: deep-dive preview */}
        <div className="lg:col-span-7">
          {activeCourse ? (
            <div className="rounded-3xl border border-cyan-500/30 bg-slate-950/90 p-6 sm:p-8 backdrop-blur-xl shadow-2xl space-y-6">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-mono font-bold text-cyan-400 uppercase">
                    {activeCourse.category} &bull; {activeCourse.level}
                  </span>
                </div>
                <h2 className="mt-1.5 text-2xl font-extrabold text-white">{activeCourse.title}</h2>
                <p className="mt-2 text-sm text-slate-300">{activeCourse.description}</p>
              </div>

              {/* Syntax Box */}
              <div className="space-y-2">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
                  <Code className="h-3.5 w-3.5 text-cyan-400" />
                  Formula Syntax
                </span>
                <div className="relative rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-cyan-200 overflow-x-auto">
                  <code>{activeCourse.syntax}</code>
                </div>
              </div>

              {/* Live Working Example */}
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
                    <Calculator className="h-3.5 w-3.5 text-emerald-400" />
                    Production Ready Implementation
                  </span>
                  <button
                    onClick={() => handleCopyFormula(activeCourse.example, activeCourse.id)}
                    className="flex items-center gap-1.5 text-xs font-semibold text-cyan-300 hover:text-white transition"
                  >
                    {copiedId === activeCourse.id ? (
                      <>
                        <Check className="h-3.5 w-3.5 text-emerald-400" />
                        <span>Copied to Clipboard</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-3.5 w-3.5" />
                        <span>Copy Formula</span>
                      </>
                    )}
                  </button>
                </div>
                <div className="rounded-2xl border border-cyan-500/30 bg-slate-900/90 p-4 font-mono text-sm text-emerald-300 font-bold overflow-x-auto">
                  <code>{activeCourse.example}</code>
                </div>
              </div>

              {/* Expected Output */}
              <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-4">
                <span className="text-[10px] font-mono font-bold uppercase text-slate-400">
                  Execution Output / Impact
                </span>
                <p className="mt-1 text-xs text-slate-300 font-sans">{activeCourse.outputPreview}</p>
              </div>

              {/* Pro Tips */}
              <div className="space-y-2">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
                  <Lightbulb className="h-3.5 w-3.5 text-amber-400" />
                  Enterprise Engineering Tips
                </span>
                <div className="space-y-1.5">
                  {activeCourse.tips.map((tip, idx) => (
                    <div key={idx} className="flex items-start gap-2 text-xs text-slate-300">
                      <CheckCircle className="h-3.5 w-3.5 text-cyan-400 mt-0.5 flex-shrink-0" />
                      <span>{tip}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
};

export default AcademyHub;
