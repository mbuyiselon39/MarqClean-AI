import { useMemo, useRef, useState } from "react";
import { extractToTable, downloadBlob, type DataTable } from "./engine";
import { runQueryPipeline, buildTransformationSummary, type QueryStep, type QueryStepType, type DataTypeTarget } from "./query-engine";
import { validateRelationship, computeMeasure, type Relationship, type Cardinality, type Measure, type MeasureAggFn, type RelationshipConflict } from "./data-model";
import { DAX_PATTERNS, DAX_CATEGORIES, checkDaxSyntaxShape, type DaxCategory } from "./dax-patterns";
import { GlobalHeader, ToolLaunchpad } from "./WorkspaceShell";
import {
  EXCEL_FUNCTIONS,
  aggregateColumn,
  andOrColumn,
  averageifs,
  correl,
  stdevS,
  varS,
  forecastLinear,
  forecastETS,
  maxifs,
  minifs,
  caseColumn,
  concatColumns,
  countifs,
  datedif,
  datePart,
  dateValue,
  timeValue,
  eomonth,
  filterTableByCriterion,
  formatDateColumn,
  hlookup,
  ifColumn,
  iferrorColumn,
  ifsBanding,
  indexMatch,
  interpretInstruction,
  ipmt,
  lenColumn,
  letCalc,
  networkdays,
  offsetAggregate,
  percentileColumn,
  pmt,
  rankColumn,
  roundColumn,
  regexReplaceColumn,
  regexTestColumn,
  regexExtractColumn,
  formulaTextColumn,
  errorTypeColumn,
  sortTableByColumn,
  substituteColumn,
  substringColumn,
  sumifs,
  sumproduct,
  switchColumn,
  textBeforeAfter,
  textJoin,
  textSplit,
  trimCleanColumn,
  typeCheckColumn,
  indirectColumn,
  sequenceColumn,
  uniqueValues,
  chooseCols,
  vstackTables,
  tocol,
  vlookup,
  weekdayColumn,
  xirr,
  irr,
  xlookup,
  xnpv,
  npv,
  type AggFn,
  type Band,
  type CaseMode,
  type Criterion,
  type CriteriaOp,
  type FunctionKey,
  type FunctionResult,
  type RoundMode,
  type SubstringMode,
  type SwitchMap,
  type TypeCheckMode,
} from "./excel-functions";
import {
  autoFormatColumn,
  classifyPredict,
  columnDistribution,
  computeClarityScore,
  downloadMatrixExcel,
  downloadTableCsv,
  downloadTableExcel,
  extractHtmlTable,
  filterTable,
  findFuzzyDuplicates,
  fuzzySearch,
  linearPredict,
  mergeAppend,
  mergeJoin,
  removeDuplicatesByColumns,
  smartFill,
  timesheetSummary,
  timesheetToMatrix,
  type ClarityReport,
  type FilterOp,
  type FormatMode,
  type TimesheetEntry,
} from "./toolbox";

type ToolKey = "functions" | "power-query" | "data-model" | "dax-reference" | "merge" | "dedupe" | "fuzzy" | "clarity" | "explorer" | "predict" | "extract" | "timesheet";

const TOOLS: Array<{ key: ToolKey; label: string; blurb: string; icon: string }> = [
  { key: "functions", label: "Advanced Excel Functions", blurb: "Workbook-backed cleanup, analysis, reconciliation, statistics, and finance formulas.", icon: "📊" },
  { key: "power-query", label: "Power Query-Style Workflow", blurb: "Import, preview, transform, combine, and export with a step-by-step log.", icon: "🧩" },
  { key: "data-model", label: "Data Model & Relationships", blurb: "Link tables, validate keys, and build cross-table measures, Power Pivot-style.", icon: "🕸️" },
  { key: "dax-reference", label: "DAX Formula Reference", blurb: "Look up common DAX patterns, their equivalent here, and check formula syntax.", icon: "🧮" },
  { key: "merge", label: "Merge & Combine", blurb: "Merge and join multiple CSV or Excel files by column.", icon: "🔗" },
  { key: "dedupe", label: "Remove Duplicates", blurb: "Exact and fuzzy duplicate detection and removal.", icon: "🧹" },
  { key: "fuzzy", label: "Fuzzy Match Search", blurb: "Similarity searching for names and addresses.", icon: "🔍" },
  { key: "clarity", label: "Data Health & Clarity Score", blurb: "Deduplicate, gap-fill, autoformat, and score data quality.", icon: "✅" },
  { key: "explorer", label: "Data Explorer", blurb: "Filter, query, and view distributions visually.", icon: "📈" },
  { key: "predict", label: "Smart Predict", blurb: "One-click prediction and classification from your data.", icon: "🤖" },
  { key: "extract", label: "Web Table Extractor", blurb: "Turn pasted HTML tables or lists into structured data.", icon: "🌐" },
  { key: "timesheet", label: "Timesheet Builder", blurb: "Track hours and generate billable time reports.", icon: "⏱️" },
];

export default function DataToolbox({ onExit }: { onExit: () => void }) {
  const [tool, setTool] = useState<ToolKey>("functions");

  return (
    <div className="ws-root ws-scope relative min-h-screen">
      <GlobalHeader />
      <div className="ws-grid pointer-events-none absolute inset-x-0 top-0 h-[38rem]" aria-hidden="true" />

      <div className="relative mx-auto mb-2 flex max-w-[110rem] flex-wrap items-center justify-between gap-3 px-5 pt-24 lg:px-8">
        <div className="flex items-center gap-3">
          <span className="flex h-10 w-10 items-center justify-center rounded-xl bg-canvas from-[rgb(var(--accent))]/20 to-[rgb(var(--accent))]/20 text-xl ring-1 ring-inset ring-white/10">🧰</span>
          <div>
            <h2 className="mc-display text-xl font-medium text-ink">Data Toolbox</h2>
            <p className="text-xs text-[#b0bacb]">Advanced Excel functions, cleansing, matching and transformation. All tools free, browser-based.</p>
          </div>
        </div>
        <button className="ws-btn-secondary rounded-md px-4 py-2 text-sm font-medium" onClick={onExit}>Back to MarqClean AI</button>
      </div>

      <div className="relative mx-auto max-w-[110rem] px-5 pb-16 pt-4 lg:px-8">
        <div className="mb-2 flex items-center justify-between">
          <p className="mc-mono text-[11px] uppercase tracking-[0.2em] text-[#8a94a8]">Tools Available ({TOOLS.length})</p>
        </div>
        <ToolLaunchpad tools={TOOLS} active={tool} onSelect={setTool} />

        <main id="main-content" className="mt-6 min-w-0">
          {tool === "functions" ? <FunctionsTool /> : null}
          {tool === "power-query" ? <PowerQueryTool /> : null}
          {tool === "data-model" ? <DataModelTool /> : null}
          {tool === "dax-reference" ? <DaxReferenceTool /> : null}
          {tool === "merge" ? <MergeTool /> : null}
          {tool === "dedupe" ? <DedupeTool /> : null}
          {tool === "fuzzy" ? <FuzzyTool /> : null}
          {tool === "clarity" ? <ClarityTool /> : null}
          {tool === "explorer" ? <ExplorerTool /> : null}
          {tool === "predict" ? <PredictTool /> : null}
          {tool === "extract" ? <ExtractTool /> : null}
          {tool === "timesheet" ? <TimesheetTool /> : null}
        </main>
      </div>
    </div>
  );
}

function Panel({ title, description, children, centered }: { title: string; description: string; children: React.ReactNode; centered?: boolean }) {
  return (
    <section>
      <div className={centered ? "mx-auto max-w-3xl text-center" : undefined}>
        <h1 className="text-3xl font-medium tracking-[-0.03em]">{title}</h1>
        <p className={centered ? "mt-2 text-justify leading-7 text-ink-2 [hyphens:auto]" : "mt-2 max-w-3xl text-ink-2"}>{description}</p>
      </div>
      <div className="mt-6">{children}</div>
    </section>
  );
}

function useSingleUpload() {
  const [table, setTable] = useState<DataTable | null>(null);
  const [error, setError] = useState("");
  const ref = useRef<HTMLInputElement>(null);

  async function onFile(file: File) {
    setError("");
    try {
      const t = await extractToTable(file);
      if (!t.headers.length) throw new Error("No readable data found.");
      setTable(t);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "The file could not be read.");
    }
  }

  return { table, setTable, error, ref, onFile };
}

function UploadBox({ label, table, refEl, onFile }: { label: string; table: DataTable | null; refEl: React.RefObject<HTMLInputElement | null>; onFile: (f: File) => void }) {
  return (
    <div
      className="rounded-xl border border-dashed border-line bg-canvas p-5 text-center transition hover:border-line"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) onFile(f); }}
    >
      <input ref={refEl} className="sr-only" type="file" accept=".pdf,.xlsx,.xls,.csv,.txt,.docx,.doc" onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); if (refEl.current) refEl.current.value = ""; }} />
      <h3 className="text-sm font-medium text-ink">{label}</h3>
      <p className="mt-1 text-xs text-ink-3">{table ? `${table.sourceName} - ${table.rows.length} rows` : "Drop or choose a CSV, Excel, PDF, or Word file."}</p>
      <button className="mt-3 rounded-md bg-ink px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => refEl.current?.click()}>
        {table ? "Replace" : "Choose file"}
      </button>
    </div>
  );
}

function TablePreview({ table, max = 8 }: { table: DataTable; max?: number }) {
  return (
    <div className="mt-4 overflow-auto rounded-xl border border-line bg-canvas" style={{ maxHeight: "22rem" }}>
      <table className="min-w-full text-left text-sm">
        <thead className="bg-surface text-xs uppercase tracking-wide text-ink-3">
          <tr>{table.headers.map((h, i) => <th key={i} className="whitespace-nowrap px-3 py-2 font-medium">{h}</th>)}</tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {table.rows.slice(0, max).map((row, r) => (
            <tr key={r}>{table.headers.map((_, c) => <td key={c} className="max-w-48 truncate px-3 py-2 text-ink-2">{row[c] || "-"}</td>)}</tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function DownloadButtons({ table, name }: { table: DataTable; name: string }) {
  return (
    <div className="mt-4 flex flex-wrap gap-2">
      <button className="rounded-md bg-accent-tint px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => downloadTableExcel(table, `${name}.xlsx`, name)}>Download Excel</button>
      <button className="rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={() => downloadTableCsv(table, `${name}.csv`)}>Download CSV</button>
    </div>
  );
}

function PowerQueryTool() {
  const [sources, setSources] = useState<Array<{ name: string; table: DataTable }>>([]);
  const [primaryName, setPrimaryName] = useState<string>("");
  const [steps, setSteps] = useState<QueryStep[]>([]);
  const [error, setError] = useState("");
  const importRef = useRef<HTMLInputElement>(null);

  const [stepType, setStepType] = useState<QueryStepType>("choose-columns");
  const [stepCol, setStepCol] = useState(0);
  const [stepCol2, setStepCol2] = useState(1);
  const [stepText, setStepText] = useState(",");
  const [stepOp, setStepOp] = useState<"=" | "!=" | ">" | "<" | ">=" | "<=" | "contains">("=");
  const [stepDir, setStepDir] = useState<"asc" | "desc">("asc");
  const [stepType2, setStepType2] = useState<DataTypeTarget>("number");
  const [chosenCols, setChosenCols] = useState<Set<number>>(new Set());
  const [stepSource, setStepSource] = useState("");

  async function onImport(file: File) {
    setError("");
    try {
      const table = await extractToTable(file);
      if (!table.headers.length) throw new Error("No readable data found.");
      setSources((prev) => [...prev, { name: file.name, table }]);
      if (!primaryName) setPrimaryName(file.name);
    } catch (err) {
      setError(err instanceof Error ? err.message : "The file could not be read.");
    }
  }

  const primary = sources.find((s) => s.name === primaryName) ?? null;
  const lookupTable = (name: string) => sources.find((s) => s.name === name)?.table;

  const outcomes = useMemo(() => {
    if (!primary) return [];
    return runQueryPipeline(primary.table, steps, lookupTable);
  }, [primary, steps, sources]);

  const currentTable = outcomes.length > 0 ? outcomes[outcomes.length - 1].table : primary?.table ?? null;

  function addStep() {
    if (!currentTable) return;
    const id = `${Date.now()}-${steps.length}`;
    let label = "";
    let params: QueryStep["params"] = {};

    if (stepType === "choose-columns") {
      if (chosenCols.size === 0) { setError("Select at least one column to keep."); return; }
      params = { columns: Array.from(chosenCols).sort((a, b) => a - b).join(",") };
      label = `Choose columns: ${Array.from(chosenCols).sort((a, b) => a - b).map((i) => currentTable.headers[i]).join(", ")}`;
    } else if (stepType === "change-type") {
      params = { column: stepCol, target: stepType2 };
      label = `Change type of "${currentTable.headers[stepCol]}" to ${stepType2}`;
    } else if (stepType === "split-column") {
      params = { column: stepCol, delimiter: stepText };
      label = `Split "${currentTable.headers[stepCol]}" by "${stepText}"`;
    } else if (stepType === "merge-columns") {
      params = { columnA: stepCol, columnB: stepCol2, separator: stepText };
      label = `Merge "${currentTable.headers[stepCol]}" + "${currentTable.headers[stepCol2]}"`;
    } else if (stepType === "remove-duplicates") {
      params = { columns: Array.from(chosenCols).join(",") };
      label = chosenCols.size > 0 ? `Remove duplicates by ${Array.from(chosenCols).map((i) => currentTable.headers[i]).join(", ")}` : "Remove duplicate rows (whole row)";
    } else if (stepType === "remove-blank-rows") {
      params = { column: stepCol };
      label = `Remove blank/error rows in "${currentTable.headers[stepCol]}"`;
    } else if (stepType === "filter-rows") {
      params = { column: stepCol, op: stepOp, value: stepText };
      label = `Filter where "${currentTable.headers[stepCol]}" ${stepOp} "${stepText}"`;
    } else if (stepType === "sort") {
      params = { column: stepCol, direction: stepDir };
      label = `Sort by "${currentTable.headers[stepCol]}" (${stepDir})`;
    } else if (stepType === "append") {
      if (!stepSource) { setError("Choose a source table to append."); return; }
      params = { source: stepSource };
      label = `Append rows from "${stepSource}"`;
    } else if (stepType === "join") {
      if (!stepSource) { setError("Choose a source table to join."); return; }
      params = { source: stepSource, leftKey: stepCol, rightKey: stepCol2 };
      label = `Join with "${stepSource}"`;
    }

    setError("");
    setSteps((prev) => [...prev, { id, type: stepType, label, params }]);
    setChosenCols(new Set());
  }

  function removeStep(id: string) {
    setSteps((prev) => prev.filter((s) => s.id !== id));
  }
  function moveStep(id: string, direction: -1 | 1) {
    setSteps((prev) => {
      const idx = prev.findIndex((s) => s.id === id);
      const target = idx + direction;
      if (idx === -1 || target < 0 || target >= prev.length) return prev;
      const next = prev.slice();
      [next[idx], next[target]] = [next[target], next[idx]];
      return next;
    });
  }

  function downloadSummary() {
    if (!primary) return;
    const summary = buildTransformationSummary(primary.name, outcomes);
    downloadBlob(summary, "transformation-summary.txt", "text/plain");
  }

  const otherSources = sources.filter((s) => s.name !== primaryName);

  return (
    <Panel title="Power Query-Style Workflow" description="Import one or more files, then build a repeatable pipeline of steps — choose columns, change types, split or merge columns, remove duplicates or blanks, filter, sort, append, or join — with a full step-by-step log, the same way Power Query's Applied Steps pane works.">
      <div className="rounded-xl border border-line bg-accent-tint p-4 text-sm text-accent">
        <p className="font-medium">What this actually is</p>
        <p className="mt-1 text-accent">
          This reproduces the <em>result</em> of a Power Query transformation — cleaned, combined, reshaped data with a
          full step log you can export. It does not generate real Power Query "M" code, and a saved pipeline will not
          appear as a live, refreshable query inside Excel's own Data → Queries &amp; Connections pane.
        </p>
      </div>

      {error ? <p className="mt-3 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</p> : null}

      <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <p className="text-sm font-medium text-ink">Sources ({sources.length})</p>
          <div>
            <input ref={importRef} type="file" className="sr-only" accept=".csv,.xlsx,.xls,.pdf,.docx,.doc,.txt" onChange={(e) => { const f = e.target.files?.[0]; if (f) onImport(f); if (importRef.current) importRef.current.value = ""; }} />
            <button className="rounded-md bg-ink px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => importRef.current?.click()}>Import file</button>
          </div>
        </div>
        {sources.length > 0 ? (
          <div className="mt-3 flex flex-wrap gap-2 text-xs">
            {sources.map((s) => (
              <button
                key={s.name}
                onClick={() => setPrimaryName(s.name)}
                className={`rounded-md border px-3 py-1.5 font-medium ${s.name === primaryName ? "border-line bg-accent-tint text-accent" : "border-line text-ink-2 hover:border-line"}`}
              >
                {s.name === primaryName ? "★ " : ""}{s.name} ({s.table.rows.length} rows)
              </button>
            ))}
          </div>
        ) : (
          <p className="mt-3 text-sm text-ink-3">Import at least one file to start building a pipeline. Import a second file if you plan to append or join.</p>
        )}
      </div>

      {primary && currentTable ? (
        <>
          <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
            <p className="text-sm font-medium text-ink">Applied steps ({steps.length})</p>
            {steps.length === 0 ? (
              <p className="mt-2 text-sm text-ink-3">No steps yet — add one below.</p>
            ) : (
              <ol className="mt-3 space-y-1.5 text-sm">
                {steps.map((s, i) => (
                  <li key={s.id} className="flex items-center justify-between gap-3 rounded-lg bg-surface px-3 py-2">
                    <span className="text-ink-2"><span className="mr-2 font-mono text-xs text-ink-3">{i + 1}.</span>{s.label}</span>
                    <span className="flex shrink-0 items-center gap-1">
                      <button className="rounded border border-line px-1.5 text-xs text-ink-3 hover:border-line" onClick={() => moveStep(s.id, -1)} aria-label={`Move step ${i + 1} up`}>↑</button>
                      <button className="rounded border border-line px-1.5 text-xs text-ink-3 hover:border-line" onClick={() => moveStep(s.id, 1)} aria-label={`Move step ${i + 1} down`}>↓</button>
                      <button className="rounded border border-rose-300 px-1.5 text-xs text-rose-600 hover:border-rose-500" onClick={() => removeStep(s.id)} aria-label={`Remove step ${i + 1}`}>✕</button>
                    </span>
                  </li>
                ))}
              </ol>
            )}
          </div>

          <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
            <p className="text-sm font-medium text-ink">Add a step</p>
            <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
              <label className="font-medium uppercase tracking-wide text-ink-3">Step type
                <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepType} onChange={(e) => setStepType(e.target.value as QueryStepType)}>
                  <option value="choose-columns">Choose columns</option>
                  <option value="change-type">Change data type</option>
                  <option value="split-column">Split column</option>
                  <option value="merge-columns">Merge columns</option>
                  <option value="remove-duplicates">Remove duplicates</option>
                  <option value="remove-blank-rows">Remove blank/error rows</option>
                  <option value="filter-rows">Filter rows</option>
                  <option value="sort">Sort</option>
                  <option value="append">Append another source</option>
                  <option value="join">Join another source</option>
                </select>
              </label>

              {(stepType === "choose-columns" || stepType === "remove-duplicates") ? (
                <div className="flex flex-col">
                  <span className="font-medium uppercase tracking-wide text-ink-3">{stepType === "choose-columns" ? "Columns to keep" : "Match columns (blank = whole row)"}</span>
                  <div className="mt-1 flex max-w-md flex-wrap gap-2">
                    {currentTable.headers.map((h, i) => (
                      <label key={i} className="flex items-center gap-1 rounded border border-line px-2 py-1 text-sm font-normal">
                        <input type="checkbox" checked={chosenCols.has(i)} onChange={() => setChosenCols((prev) => { const n = new Set(prev); if (n.has(i)) n.delete(i); else n.add(i); return n; })} />{h}
                      </label>
                    ))}
                  </div>
                </div>
              ) : null}

              {stepType === "change-type" ? (
                <>
                  <ColSelect2 table={currentTable} label="Column" value={stepCol} onChange={setStepCol} />
                  <label className="font-medium uppercase tracking-wide text-ink-3">Convert to
                    <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepType2} onChange={(e) => setStepType2(e.target.value as DataTypeTarget)}>
                      <option value="number">Number</option>
                      <option value="text">Text</option>
                      <option value="date-iso">Date (ISO)</option>
                    </select>
                  </label>
                </>
              ) : null}

              {stepType === "split-column" ? (
                <>
                  <ColSelect2 table={currentTable} label="Column" value={stepCol} onChange={setStepCol} />
                  <label className="font-medium uppercase tracking-wide text-ink-3">Delimiter<input className="mt-1 block w-20 rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepText} onChange={(e) => setStepText(e.target.value)} /></label>
                </>
              ) : null}

              {stepType === "merge-columns" ? (
                <>
                  <ColSelect2 table={currentTable} label="First column" value={stepCol} onChange={setStepCol} />
                  <ColSelect2 table={currentTable} label="Second column" value={stepCol2} onChange={setStepCol2} />
                  <label className="font-medium uppercase tracking-wide text-ink-3">Separator<input className="mt-1 block w-20 rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepText} onChange={(e) => setStepText(e.target.value)} /></label>
                </>
              ) : null}

              {stepType === "remove-blank-rows" ? <ColSelect2 table={currentTable} label="Column" value={stepCol} onChange={setStepCol} /> : null}

              {stepType === "filter-rows" ? (
                <>
                  <ColSelect2 table={currentTable} label="Column" value={stepCol} onChange={setStepCol} />
                  <label className="font-medium uppercase tracking-wide text-ink-3">Operator
                    <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepOp} onChange={(e) => setStepOp(e.target.value as typeof stepOp)}>
                      <option value="=">=</option><option value="!=">≠</option><option value=">">&gt;</option><option value="<">&lt;</option><option value=">=">≥</option><option value="<=">≤</option><option value="contains">contains</option>
                    </select>
                  </label>
                  <label className="font-medium uppercase tracking-wide text-ink-3">Value<input className="mt-1 block w-28 rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepText} onChange={(e) => setStepText(e.target.value)} /></label>
                </>
              ) : null}

              {stepType === "sort" ? (
                <>
                  <ColSelect2 table={currentTable} label="Column" value={stepCol} onChange={setStepCol} />
                  <label className="font-medium uppercase tracking-wide text-ink-3">Direction
                    <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepDir} onChange={(e) => setStepDir(e.target.value as "asc" | "desc")}>
                      <option value="asc">Ascending</option><option value="desc">Descending</option>
                    </select>
                  </label>
                </>
              ) : null}

              {(stepType === "append" || stepType === "join") ? (
                otherSources.length === 0 ? (
                  <p className="text-ink-3">Import a second file to append or join.</p>
                ) : (
                  <>
                    <label className="font-medium uppercase tracking-wide text-ink-3">Other source
                      <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={stepSource} onChange={(e) => setStepSource(e.target.value)}>
                        <option value="">Choose...</option>
                        {otherSources.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
                      </select>
                    </label>
                    {stepType === "join" ? (
                      <>
                        <ColSelect2 table={currentTable} label="Key in current" value={stepCol} onChange={setStepCol} />
                        <ColSelect2 table={sources.find((s) => s.name === stepSource)?.table ?? currentTable} label="Key in other" value={stepCol2} onChange={setStepCol2} />
                      </>
                    ) : null}
                  </>
                )
              ) : null}

              <button className="rounded-md bg-accent-tint px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={addStep}>Add step</button>
            </div>
          </div>

          <div className="mt-4">
            <p className="text-sm text-ink-2">Preview after {steps.length} step{steps.length === 1 ? "" : "s"}: {currentTable.rows.length} rows, {currentTable.headers.length} columns.</p>
            <TablePreview table={currentTable} />
            <DownloadButtons table={currentTable} name="power-query-output" />
            <button className="mt-2 rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={downloadSummary}>Download transformation summary</button>
          </div>
        </>
      ) : null}
    </Panel>
  );
}

function DataModelTool() {
  const [sources, setSources] = useState<Array<{ name: string; table: DataTable }>>([]);
  const importRef = useRef<HTMLInputElement>(null);
  const [error, setError] = useState("");
  const [relationships, setRelationships] = useState<Relationship[]>([]);
  const [measures, setMeasures] = useState<Measure[]>([]);
  const [result, setResult] = useState<DataTable | null>(null);

  const [oneTable, setOneTable] = useState("");
  const [oneKey, setOneKey] = useState(0);
  const [manyTable, setManyTable] = useState("");
  const [manyKey, setManyKey] = useState(0);
  const [cardinality, setCardinality] = useState<Cardinality>("one-to-many");

  const [measureName, setMeasureName] = useState("");
  const [measureRel, setMeasureRel] = useState("");
  const [measureCol, setMeasureCol] = useState(0);
  const [measureFn, setMeasureFn] = useState<MeasureAggFn>("sum");
  const [baseTable, setBaseTable] = useState("");

  async function onImport(file: File) {
    setError("");
    try {
      const table = await extractToTable(file);
      if (!table.headers.length) throw new Error("No readable data found.");
      setSources((prev) => [...prev, { name: file.name, table }]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "The file could not be read.");
    }
  }

  const tableMap = useMemo(() => new Map(sources.map((s) => [s.name, s.table])), [sources]);

  const conflictsByRel = useMemo(() => {
    const map = new Map<string, RelationshipConflict[]>();
    relationships.forEach((rel) => map.set(rel.id, validateRelationship(tableMap, rel)));
    return map;
  }, [relationships, tableMap]);

  function addRelationship() {
    if (!oneTable || !manyTable) { setError("Choose both a 'one' side table and a 'many' side table."); return; }
    if (oneTable === manyTable) { setError("The two sides of a relationship must be different tables."); return; }
    setError("");
    setRelationships((prev) => [...prev, { id: `${Date.now()}`, oneTable, oneKey, manyTable, manyKey, cardinality }]);
  }

  function removeRelationship(id: string) {
    setRelationships((prev) => prev.filter((r) => r.id !== id));
    setMeasures((prev) => prev.filter((m) => m.relationshipId !== id));
  }

  function addMeasure() {
    if (!measureName.trim()) { setError("Give the measure a name."); return; }
    if (!measureRel) { setError("Choose which relationship this measure aggregates across."); return; }
    setError("");
    setMeasures((prev) => [...prev, { id: `${Date.now()}`, name: measureName.trim(), relationshipId: measureRel, aggColumn: measureCol, aggFn: measureFn }]);
    setMeasureName("");
  }

  function removeMeasure(id: string) {
    setMeasures((prev) => prev.filter((m) => m.id !== id));
  }

  function buildModel() {
    if (!baseTable) { setError("Choose a base table to build the model output from."); return; }
    const hasErrors = relationships.some((r) => (conflictsByRel.get(r.id) ?? []).some((c) => c.severity === "error"));
    if (hasErrors) { setError("Fix relationship errors (shown in red below) before building the model."); return; }
    setError("");
    const working = new Map(tableMap);
    let output = working.get(baseTable)!;
    measures
      .filter((m) => relationships.find((r) => r.id === m.relationshipId)?.oneTable === baseTable)
      .forEach((m) => {
        output = computeMeasure(working, relationships, m);
        working.set(baseTable, output);
      });
    setResult(output);
  }

  const manyTableForMeasure = relationships.find((r) => r.id === measureRel);
  const measureSourceTable = manyTableForMeasure ? tableMap.get(manyTableForMeasure.manyTable) : undefined;

  return (
    <Panel title="Data Model & Relationships" description="Link tables the way Power Pivot links them: declare a relationship, get real validation against your actual data, see the model as a diagram, and build cross-table measures.">
      <div className="rounded-xl border border-line bg-accent-tint p-4 text-sm text-accent">
        <p className="font-medium">What this actually is</p>
        <p className="mt-1 text-accent">
          This is not an in-memory columnar database and does not support dynamic, arbitrary filter-context
          re-aggregation the way real Power Pivot measures do inside Excel. Measures here are computed once when you
          build the model, grouped by the matching key — recalculating after new data arrives means rebuilding.
        </p>
      </div>

      {error ? <p className="mt-3 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{error}</p> : null}

      <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <p className="text-sm font-medium text-ink">Tables ({sources.length})</p>
          <div>
            <input ref={importRef} type="file" className="sr-only" accept=".csv,.xlsx,.xls" onChange={(e) => { const f = e.target.files?.[0]; if (f) onImport(f); if (importRef.current) importRef.current.value = ""; }} />
            <button className="rounded-md bg-ink px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => importRef.current?.click()}>Import table</button>
          </div>
        </div>
        {sources.length === 0 ? <p className="mt-3 text-sm text-ink-3">Import at least two tables to build relationships between them.</p> : (
          <div className="mt-3 flex flex-wrap gap-2 text-xs">
            {sources.map((s) => <span key={s.name} className="rounded-md border border-line px-3 py-1.5 text-ink-2">{s.name} ({s.table.rows.length} rows)</span>)}
          </div>
        )}
      </div>

      {sources.length >= 2 ? (
        <>
          <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
            <p className="text-sm font-medium text-ink">Relationships</p>
            <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
              <label className="font-medium uppercase tracking-wide text-ink-3">"One" side table
                <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={oneTable} onChange={(e) => { setOneTable(e.target.value); setOneKey(0); }}>
                  <option value="">Choose...</option>
                  {sources.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
                </select>
              </label>
              {oneTable ? <ColSelect2 table={tableMap.get(oneTable)!} label="Key column" value={oneKey} onChange={setOneKey} /> : null}
              <label className="font-medium uppercase tracking-wide text-ink-3">"Many" side table
                <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={manyTable} onChange={(e) => { setManyTable(e.target.value); setManyKey(0); }}>
                  <option value="">Choose...</option>
                  {sources.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
                </select>
              </label>
              {manyTable ? <ColSelect2 table={tableMap.get(manyTable)!} label="Key column" value={manyKey} onChange={setManyKey} /> : null}
              <label className="font-medium uppercase tracking-wide text-ink-3">Cardinality
                <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={cardinality} onChange={(e) => setCardinality(e.target.value as Cardinality)}>
                  <option value="one-to-many">One-to-many</option>
                  <option value="one-to-one">One-to-one</option>
                </select>
              </label>
              <button className="rounded-md bg-accent-tint px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={addRelationship}>Add relationship</button>
            </div>

            {relationships.length > 0 ? (
              <ul className="mt-4 space-y-2 text-sm">
                {relationships.map((rel) => {
                  const conflicts = conflictsByRel.get(rel.id) ?? [];
                  const hasError = conflicts.some((c) => c.severity === "error");
                  return (
                    <li key={rel.id} className={`rounded-lg border p-3 ${hasError ? "border-rose-300 bg-rose-50" : conflicts.length > 0 ? "border-line bg-surface" : "border-line bg-surface"}`}>
                      <div className="flex items-center justify-between gap-3">
                        <span className="font-medium text-ink">
                          {rel.oneTable}.{tableMap.get(rel.oneTable)?.headers[rel.oneKey]} → {rel.manyTable}.{tableMap.get(rel.manyTable)?.headers[rel.manyKey]} ({rel.cardinality})
                        </span>
                        <button className="rounded border border-line px-2 py-0.5 text-xs text-ink-3 hover:border-rose-500" onClick={() => removeRelationship(rel.id)}>Remove</button>
                      </div>
                      {conflicts.length > 0 ? (
                        <ul className="mt-2 space-y-1 text-xs">
                          {conflicts.map((c, i) => (
                            <li key={i} className={c.severity === "error" ? "text-rose-700" : "text-ink-2"}>
                              {c.severity === "error" ? "⛔" : "⚠️"} {c.message}
                            </li>
                          ))}
                        </ul>
                      ) : (
                        <p className="mt-1 text-xs text-ink-2">✓ No conflicts found — every key on the "one" side is unique, and every row matches.</p>
                      )}
                    </li>
                  );
                })}
              </ul>
            ) : null}
          </div>

          {relationships.length > 0 ? (
            <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
              <p className="text-sm font-medium text-ink">Relationship diagram</p>
              <RelationshipDiagram sources={sources} relationships={relationships} conflictsByRel={conflictsByRel} />
            </div>
          ) : null}

          {relationships.length > 0 ? (
            <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
              <p className="text-sm font-medium text-ink">Measures</p>
              <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
                <label className="font-medium uppercase tracking-wide text-ink-3">Name<input className="mt-1 block w-36 rounded border border-line px-2 py-1.5 text-sm font-normal" value={measureName} onChange={(e) => setMeasureName(e.target.value)} placeholder="Total Orders" /></label>
                <label className="font-medium uppercase tracking-wide text-ink-3">Relationship
                  <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={measureRel} onChange={(e) => { setMeasureRel(e.target.value); setMeasureCol(0); }}>
                    <option value="">Choose...</option>
                    {relationships.map((r) => <option key={r.id} value={r.id}>{r.oneTable} ← {r.manyTable}</option>)}
                  </select>
                </label>
                {measureSourceTable ? <ColSelect2 table={measureSourceTable} label="Column to aggregate" value={measureCol} onChange={setMeasureCol} /> : null}
                <label className="font-medium uppercase tracking-wide text-ink-3">Aggregation
                  <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={measureFn} onChange={(e) => setMeasureFn(e.target.value as MeasureAggFn)}>
                    <option value="sum">Sum</option><option value="average">Average</option><option value="count">Count</option><option value="min">Min</option><option value="max">Max</option>
                  </select>
                </label>
                <button className="rounded-md bg-accent-tint px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={addMeasure}>Add measure</button>
              </div>
              {measures.length > 0 ? (
                <ul className="mt-3 space-y-1.5 text-sm">
                  {measures.map((m) => {
                    const rel = relationships.find((r) => r.id === m.relationshipId);
                    return (
                      <li key={m.id} className="flex items-center justify-between rounded-lg bg-surface px-3 py-2">
                        <span className="text-ink-2">{m.name}: {m.aggFn}({rel?.manyTable}.{rel ? tableMap.get(rel.manyTable)?.headers[m.aggColumn] : ""}) grouped by {rel?.oneTable}</span>
                        <button className="rounded border border-rose-300 px-2 py-0.5 text-xs text-rose-600 hover:border-rose-500" onClick={() => removeMeasure(m.id)}>Remove</button>
                      </li>
                    );
                  })}
                </ul>
              ) : null}
            </div>
          ) : null}

          {measures.length > 0 ? (
            <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
              <p className="text-sm font-medium text-ink">Build model</p>
              <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
                <label className="font-medium uppercase tracking-wide text-ink-3">Base table (the "one" side to attach measures to)
                  <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={baseTable} onChange={(e) => setBaseTable(e.target.value)}>
                    <option value="">Choose...</option>
                    {Array.from(new Set(relationships.map((r) => r.oneTable))).map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </label>
                <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={buildModel}>Build model</button>
              </div>
              {result ? (
                <div className="mt-4">
                  <TablePreview table={result} />
                  <DownloadButtons table={result} name="data-model-output" />
                </div>
              ) : null}
            </div>
          ) : null}
        </>
      ) : null}
    </Panel>
  );
}

function RelationshipDiagram({ sources, relationships, conflictsByRel }: { sources: Array<{ name: string; table: DataTable }>; relationships: Relationship[]; conflictsByRel: Map<string, RelationshipConflict[]> }) {
  const tableNames = Array.from(new Set(relationships.flatMap((r) => [r.oneTable, r.manyTable])));
  const width = 720;
  const nodeWidth = 160;
  const gap = tableNames.length > 1 ? (width - nodeWidth) / (tableNames.length - 1) : 0;
  const positions = new Map(tableNames.map((name, i) => [name, { x: i * gap + nodeWidth / 2, y: 60 }]));

  return (
    <svg viewBox={`0 0 ${width} 160`} className="mt-3 w-full" role="img" aria-label="Relationship diagram between tables">
      {relationships.map((rel) => {
        const from = positions.get(rel.oneTable)!;
        const to = positions.get(rel.manyTable)!;
        const hasError = (conflictsByRel.get(rel.id) ?? []).some((c) => c.severity === "error");
        return (
          <g key={rel.id}>
            <line x1={from.x} y1={from.y} x2={to.x} y2={to.y} stroke={hasError ? "#e11d48" : "rgb(var(--accent))"} strokeWidth={2} />
            <text x={(from.x + to.x) / 2} y={from.y - 8} textAnchor="middle" fontSize="10" fill="#64748b">{rel.cardinality}</text>
          </g>
        );
      })}
      {tableNames.map((name) => {
        const pos = positions.get(name)!;
        const table = sources.find((s) => s.name === name)?.table;
        return (
          <g key={name}>
            <rect x={pos.x - nodeWidth / 2} y={pos.y - 24} width={nodeWidth} height={48} rx={10} fill="rgb(var(--ink))" />
            <text x={pos.x} y={pos.y - 4} textAnchor="middle" fontSize="12" fill="#ffffff" fontWeight="600">{name.length > 20 ? name.slice(0, 18) + "…" : name}</text>
            <text x={pos.x} y={pos.y + 14} textAnchor="middle" fontSize="10" fill="#94a3b8">{table?.rows.length ?? 0} rows</text>
          </g>
        );
      })}
    </svg>
  );
}

function DaxReferenceTool() {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState<DaxCategory | "All">("All");
  const [formula, setFormula] = useState("");

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return DAX_PATTERNS.filter((p) => {
      if (category !== "All" && p.category !== category) return false;
      if (!q) return true;
      return p.name.toLowerCase().includes(q) || p.equivalent.toLowerCase().includes(q);
    });
  }, [query, category]);

  const checkResult = useMemo(() => (formula.trim() ? checkDaxSyntaxShape(formula) : null), [formula]);

  return (
    <Panel title="DAX Formula Reference" description="Look up common DAX patterns across aggregation, filtering, time intelligence, relationships, and finance, see the equivalent operation this platform can actually run, and check a formula's syntax shape.">
      <div className="rounded-xl border border-line bg-accent-tint p-4 text-sm text-accent">
        <p className="font-medium">What this actually is</p>
        <p className="mt-1 text-accent">
          This is a DAX formula reference and pattern library, not a DAX engine. Formulas are explained and mapped to
          an equivalent calculation this platform can run — they are not parsed, executed, or validated as real DAX,
          and are not guaranteed to work if pasted into Power BI or Excel's Power Pivot without adjustment. The syntax
          checker below only verifies shape (balanced parentheses, a recognized function name, a plausible argument
          count) — it does not evaluate row or filter context.
        </p>
      </div>

      <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
        <p className="text-sm font-medium text-ink">Check a formula's syntax shape</p>
        <textarea
          className="mt-2 w-full rounded-lg border border-line px-3 py-2 font-mono text-sm"
          rows={2}
          placeholder='e.g. CALCULATE(SUM(Orders[Amount]), Orders[Region] = "East")'
          value={formula}
          onChange={(e) => setFormula(e.target.value)}
        />
        {checkResult ? (
          <div className={`mt-2 rounded-lg border p-3 text-sm ${checkResult.ok ? "border-line bg-surface text-ink-2" : "border-line bg-surface text-ink-2"}`}>
            {checkResult.ok ? "✓ Syntax shape looks valid." : (
              <ul className="space-y-1">{checkResult.issues.map((issue, i) => <li key={i}>⚠️ {issue}</li>)}</ul>
            )}
            {checkResult.matchedFunction ? (
              <p className="mt-2 border-t border-current/20 pt-2 text-xs">
                <strong>{checkResult.matchedFunction.name}</strong> equivalent here: {checkResult.matchedFunction.equivalent}
              </p>
            ) : null}
          </div>
        ) : null}
      </div>

      <div className="mt-4 flex flex-wrap items-center gap-3">
        <input
          type="search"
          placeholder="Search DAX functions..."
          className="w-full max-w-xs rounded-lg border border-line px-3 py-2 text-sm"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <div className="flex flex-wrap gap-2 text-xs">
          {(["All", ...DAX_CATEGORIES] as const).map((c) => (
            <button key={c} onClick={() => setCategory(c)} className={`rounded-md border px-3 py-1.5 font-medium ${category === c ? "border-line bg-accent-tint text-accent" : "border-line text-ink-2"}`}>{c}</button>
          ))}
        </div>
      </div>

      <div className="mt-4 grid gap-3 md:grid-cols-2">
        {filtered.map((p) => (
          <div key={p.name} className="rounded-xl border border-line bg-canvas p-4">
            <div className="flex items-center justify-between gap-2">
              <span className="font-mono text-sm font-medium text-ink">{p.name}</span>
              <span className="rounded-md border border-line px-2 py-0.5 text-[10px] text-ink-3">{p.category}</span>
            </div>
            <pre className="mt-2 overflow-x-auto rounded bg-surface px-2 py-1.5 font-mono text-xs text-ink-2">{p.daxSyntax}</pre>
            <p className="mt-1 font-mono text-xs text-ink-2">{p.daxExample}</p>
            <p className="mt-2 text-sm text-ink-2">{p.equivalent}</p>
          </div>
        ))}
      </div>
    </Panel>
  );
}

function ColSelect2({ table, label, value, onChange }: { table: DataTable; label: string; value: number; onChange: (n: number) => void }) {
  return (
    <label className="font-medium uppercase tracking-wide text-ink-3">{label}
      <select className="mt-1 block rounded border border-line px-2 py-1.5 text-sm font-normal" value={Math.min(value, table.headers.length - 1)} onChange={(e) => onChange(Number(e.target.value))}>
        {table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
      </select>
    </label>
  );
}

function MergeTool() {
  const a = useSingleUpload();
  const b = useSingleUpload();
  const [mode, setMode] = useState<"append" | "join">("append");
  const [leftKey, setLeftKey] = useState(0);
  const [rightKey, setRightKey] = useState(0);
  const [result, setResult] = useState<DataTable | null>(null);

  function run() {
    if (!a.table || !b.table) return;
    if (mode === "append") setResult(mergeAppend([a.table, b.table]));
    else setResult(mergeJoin(a.table, b.table, leftKey, rightKey));
  }

  return (
    <Panel title="Merge & Combine" description="Combine two spreadsheets. Append stacks rows and aligns matching columns automatically. Join matches rows from a second file using a key column.">
      <div className="grid gap-4 md:grid-cols-2">
        <UploadBox label="First file" table={a.table} refEl={a.ref} onFile={a.onFile} />
        <UploadBox label="Second file" table={b.table} refEl={b.ref} onFile={b.onFile} />
      </div>
      {a.error || b.error ? <p className="mt-3 text-sm text-error">{a.error || b.error}</p> : null}
      {a.table && b.table ? (
        <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <label className="font-medium">Mode</label>
            <select className="rounded border border-line px-3 py-1.5" value={mode} onChange={(e) => setMode(e.target.value as "append" | "join")}>
              <option value="append">Append (stack rows)</option>
              <option value="join">Join (match by key)</option>
            </select>
            {mode === "join" ? (
              <>
                <label>First key</label>
                <select className="rounded border border-line px-2 py-1" value={leftKey} onChange={(e) => setLeftKey(Number(e.target.value))}>
                  {a.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
                <label>Second key</label>
                <select className="rounded border border-line px-2 py-1" value={rightKey} onChange={(e) => setRightKey(Number(e.target.value))}>
                  {b.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
              </>
            ) : null}
            <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={run}>Merge</button>
          </div>
        </div>
      ) : null}
      {result ? (<><p className="mt-4 text-sm text-ink-2">{result.rows.length} rows, {result.headers.length} columns.</p><TablePreview table={result} /><DownloadButtons table={result} name="merged" /></>) : null}
    </Panel>
  );
}

function DedupeTool() {
  const up = useSingleUpload();
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [result, setResult] = useState<{ table: DataTable; removed: number } | null>(null);
  const [fuzzyCol, setFuzzyCol] = useState(0);
  const [threshold, setThreshold] = useState(0.85);
  const [fuzzyPairs, setFuzzyPairs] = useState<ReturnType<typeof findFuzzyDuplicates> | null>(null);

  function toggle(i: number) { setSelected((prev) => { const n = new Set(prev); if (n.has(i)) n.delete(i); else n.add(i); return n; }); }

  return (
    <Panel title="Remove Duplicates" description="Remove exact duplicate rows based on selected columns, or find near-duplicate (fuzzy) records for review.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
          <h3 className="text-sm font-medium">Exact duplicate columns (leave all unchecked for whole-row match)</h3>
          <div className="mt-2 flex flex-wrap gap-2 text-sm">
            {up.table.headers.map((h, i) => (
              <label key={i} className="flex items-center gap-1 rounded border border-line px-2 py-1"><input type="checkbox" checked={selected.has(i)} onChange={() => toggle(i)} />{h}</label>
            ))}
          </div>
          <button className="mt-3 rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => setResult(removeDuplicatesByColumns(up.table!, Array.from(selected)))}>Remove duplicates</button>

          <div className="mt-6 border-t border-line pt-4">
            <h3 className="text-sm font-medium">Fuzzy duplicate finder</h3>
            <div className="mt-2 flex flex-wrap items-center gap-3 text-sm">
              <select className="rounded border border-line px-2 py-1" value={fuzzyCol} onChange={(e) => setFuzzyCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <label>Similarity</label>
              <input type="range" min={0.6} max={0.98} step={0.01} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
              <span>{Math.round(threshold * 100)}%</span>
              <button className="rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={() => setFuzzyPairs(findFuzzyDuplicates(up.table!, fuzzyCol, threshold))}>Find near-duplicates</button>
            </div>
            {fuzzyPairs ? (
              <div className="mt-3 max-h-64 overflow-auto text-sm">
                <p className="text-ink-2">{fuzzyPairs.length} likely duplicate pairs</p>
                <table className="mt-2 min-w-full text-left"><tbody className="divide-y divide-slate-100">
                  {fuzzyPairs.map((p, i) => <tr key={i}><td className="py-1 pr-4">{p.valueA}</td><td className="py-1 pr-4">{p.valueB}</td><td className="py-1 text-accent">{Math.round(p.score * 100)}%</td></tr>)}
                </tbody></table>
              </div>
            ) : null}
          </div>
        </div>
      ) : null}
      {result ? (<><p className="mt-4 text-sm text-ink-2">Removed {result.removed} duplicates. {result.table.rows.length} rows remain.</p><TablePreview table={result.table} /><DownloadButtons table={result.table} name="deduplicated" /></>) : null}
    </Panel>
  );
}

function FuzzyTool() {
  const up = useSingleUpload();
  const [col, setCol] = useState(0);
  const [query, setQuery] = useState("");
  const [threshold, setThreshold] = useState(0.7);
  const [results, setResults] = useState<ReturnType<typeof fuzzySearch> | null>(null);

  return (
    <Panel title="Fuzzy Match Search" description="Similarity-based searching and matching for names and addresses, tuned for compliance-grade approximate matching.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <label>Search column</label>
            <select className="rounded border border-line px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <input className="min-w-48 flex-1 rounded border border-line px-3 py-1.5" placeholder="Search name or address" value={query} onChange={(e) => setQuery(e.target.value)} />
            <label>Min similarity</label>
            <input type="range" min={0.4} max={0.95} step={0.01} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
            <span>{Math.round(threshold * 100)}%</span>
            <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => setResults(fuzzySearch(up.table!, col, query, threshold))}>Search</button>
          </div>
          {results ? (
            <div className="mt-3 max-h-72 overflow-auto text-sm">
              <p className="text-ink-2">{results.length} matches</p>
              <table className="mt-2 min-w-full text-left"><tbody className="divide-y divide-slate-100">
                {results.map((r, i) => <tr key={i}><td className="py-1 pr-4">{r.value}</td><td className="py-1 text-accent">{Math.round(r.score * 100)}%</td></tr>)}
              </tbody></table>
            </div>
          ) : null}
        </div>
      ) : null}
    </Panel>
  );
}

function ClarityTool() {
  const up = useSingleUpload();
  const [report, setReport] = useState<ClarityReport | null>(null);
  const [col, setCol] = useState(0);
  const [fmt, setFmt] = useState<FormatMode>("title");

  function recompute(next: DataTable) { up.setTable(next); setReport(computeClarityScore(next)); }

  return (
    <Panel title="Data Health & Clarity Score" description="Deduplicate, fill gaps, standardize formatting, and get a live clarity score that shows overall data quality.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={(f) => { up.onFile(f).then(() => undefined); }} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 space-y-4">
          <div className="rounded-xl border border-line bg-canvas p-5">
            <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => setReport(computeClarityScore(up.table!))}>Compute clarity score</button>
            {report ? (
              <div className="mt-4">
                <div className="flex items-center gap-4">
                  <div className="flex h-20 w-20 items-center justify-center rounded-md bg-accent-tint text-2xl font-medium text-accent">{report.score}</div>
                  <div className="text-sm text-ink-2">
                    <p>{report.totalRows} rows, {report.filledCells}/{report.totalCells} cells filled</p>
                    <p>{report.duplicateRows} duplicate rows</p>
                  </div>
                </div>
                {report.issues.length ? (
                  <div className="mt-4 max-h-48 overflow-auto text-sm">
                    <table className="min-w-full text-left"><thead className="text-xs uppercase text-ink-3"><tr><th className="py-1 pr-4">Column</th><th className="py-1 pr-4">Issue</th><th className="py-1">Count</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">{report.issues.map((iss, i) => <tr key={i}><td className="py-1 pr-4">{iss.column}</td><td className="py-1 pr-4 text-ink-2">{iss.issue}</td><td className="py-1">{iss.count}</td></tr>)}</tbody></table>
                  </div>
                ) : null}
              </div>
            ) : null}
          </div>

          <div className="rounded-xl border border-line bg-canvas p-5">
            <h3 className="text-sm font-medium">Auto clean actions</h3>
            <div className="mt-2 flex flex-wrap items-center gap-3 text-sm">
              <select className="rounded border border-line px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <button className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={() => { const r = smartFill(up.table!, col); recompute(r.table); }}>Fill gaps (mode)</button>
              <select className="rounded border border-line px-2 py-1" value={fmt} onChange={(e) => setFmt(e.target.value as FormatMode)}>
                <option value="title">Title Case</option><option value="upper">UPPERCASE</option><option value="lower">lowercase</option><option value="trim">Trim</option>
              </select>
              <button className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={() => recompute(autoFormatColumn(up.table!, col, fmt))}>Auto format</button>
              <button className="rounded-md border border-line px-3 py-1.5 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={() => { const r = removeDuplicatesByColumns(up.table!, []); recompute(r.table); }}>Remove duplicate rows</button>
            </div>
          </div>

          <TablePreview table={up.table} />
          <DownloadButtons table={up.table} name="cleaned" />
        </div>
      ) : null}
    </Panel>
  );
}

function ExplorerTool() {
  const up = useSingleUpload();
  const [col, setCol] = useState(0);
  const [op, setOp] = useState<FilterOp>("contains");
  const [value, setValue] = useState("");
  const [view, setView] = useState<DataTable | null>(null);
  const dist = useMemo(() => (up.table ? columnDistribution(up.table, col) : []), [up.table, col]);
  const distMax = Math.max(1, ...dist.map((d) => d.count));

  return (
    <Panel title="Data Explorer" description="Visually explore your data. Filter and query rows, and view auto-updating distributions for any column.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 grid gap-4 lg:grid-cols-2">
          <div className="rounded-xl border border-line bg-canvas p-5">
            <h3 className="text-sm font-medium">Filter</h3>
            <div className="mt-2 flex flex-wrap items-center gap-2 text-sm">
              <select className="rounded border border-line px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <select className="rounded border border-line px-2 py-1" value={op} onChange={(e) => setOp(e.target.value as FilterOp)}>
                <option value="contains">contains</option><option value="equals">equals</option><option value="gt">greater than</option><option value="lt">less than</option><option value="nonempty">is not empty</option><option value="empty">is empty</option>
              </select>
              <input className="rounded border border-line px-2 py-1" value={value} onChange={(e) => setValue(e.target.value)} placeholder="value" />
              <button className="rounded-md bg-ink px-4 py-1.5 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => setView(filterTable(up.table!, col, op, value))}>Apply</button>
            </div>
            {view ? (<><p className="mt-3 text-sm text-ink-2">{view.rows.length} matching rows</p><DownloadButtons table={view} name="filtered" /></>) : null}
          </div>
          <div className="rounded-xl border border-line bg-canvas p-5">
            <h3 className="text-sm font-medium">Distribution of {up.table.headers[col]}</h3>
            <div className="mt-3 space-y-2">
              {dist.map((d) => (
                <div key={d.label}>
                  <div className="flex justify-between text-xs text-ink-2"><span className="truncate pr-2">{d.label}</span><span className="font-medium">{d.count}</span></div>
                  <div className="mt-1 h-2 rounded-md bg-surface"><div className="h-2 rounded-md bg-accent-tint" style={{ width: `${Math.max((d.count / distMax) * 100, 2)}%` }} /></div>
                </div>
              ))}
            </div>
          </div>
        </div>
      ) : null}
      {view ? <TablePreview table={view} /> : up.table ? <TablePreview table={up.table} /> : null}
    </Panel>
  );
}

function PredictTool() {
  const up = useSingleUpload();
  const [feature, setFeature] = useState(0);
  const [target, setTarget] = useState(1);
  const [input, setInput] = useState("");
  const [mode, setMode] = useState<"regression" | "classify">("regression");
  const [output, setOutput] = useState<string>("");

  function run() {
    if (!up.table) return;
    const x = parseFloat(input);
    if (mode === "regression") {
      const r = linearPredict(up.table, feature, target, x);
      setOutput(r ? `Predicted ${up.table.headers[target]} = ${r.prediction.toFixed(2)} (model fit R2 ${(r.r2 * 100).toFixed(0)}%, ${r.points} points)` : "Not enough numeric data for a regression.");
    } else {
      const r = classifyPredict(up.table, feature, target, x);
      setOutput(r ? `Predicted ${up.table.headers[target]} = ${r.label} (from ${r.neighbours} nearest records)` : "Not enough labelled data to classify.");
    }
  }

  return (
    <Panel title="Smart Predict" description="One-click machine learning. Upload a CSV, pick a feature and a target, and get an instant prediction using regression or nearest-neighbour classification.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-xl border border-line bg-canvas p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <select className="rounded border border-line px-2 py-1" value={mode} onChange={(e) => setMode(e.target.value as "regression" | "classify")}>
              <option value="regression">Predict number (regression)</option>
              <option value="classify">Predict category (classification)</option>
            </select>
            <label>Feature</label>
            <select className="rounded border border-line px-2 py-1" value={feature} onChange={(e) => setFeature(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <label>Target</label>
            <select className="rounded border border-line px-2 py-1" value={target} onChange={(e) => setTarget(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <input className="w-28 rounded border border-line px-2 py-1" placeholder="feature value" value={input} onChange={(e) => setInput(e.target.value)} />
            <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={run}>Predict</button>
          </div>
          {output ? <p className="mt-4 rounded-xl bg-accent-tint px-4 py-3 text-sm font-medium text-accent">{output}</p> : null}
        </div>
      ) : null}
    </Panel>
  );
}

function ExtractTool() {
  const [html, setHtml] = useState("");
  const [table, setTable] = useState<DataTable | null>(null);
  const [error, setError] = useState("");

  function run() {
    setError("");
    const t = extractHtmlTable(html);
    if (!t) { setError("No table or list found in the pasted HTML."); setTable(null); return; }
    setTable(t);
  }

  return (
    <Panel title="Web Table Extractor" description="Turn web data into structured spreadsheets. Paste the HTML of a web table or list and export it to Excel or CSV, no code required.">
      <textarea className="h-40 w-full rounded-xl border border-line bg-canvas px-4 py-3 text-sm" placeholder="Paste HTML containing a <table> or list here" value={html} onChange={(e) => setHtml(e.target.value)} />
      <div className="mt-3 flex gap-2">
        <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={run}>Extract table</button>
      </div>
      {error ? <p className="mt-3 text-sm text-error">{error}</p> : null}
      {table ? (<><p className="mt-4 text-sm text-ink-2">{table.rows.length} rows, {table.headers.length} columns.</p><TablePreview table={table} /><DownloadButtons table={table} name="extracted" /></>) : null}
    </Panel>
  );
}

function FunctionsTool() {
  const up = useSingleUpload();
  const [fnKey, setFnKey] = useState<FunctionKey>("xlookup");
  const [instruction, setInstruction] = useState("");
  const [note, setNote] = useState("");
  const [result, setResult] = useState<FunctionResult | null>(null);

  // shared params
  const [lookupCol, setLookupCol] = useState(0);
  const [returnCol, setReturnCol] = useState(1);
  const [sumCol, setSumCol] = useState(0);
  const [valueCol, setValueCol] = useState(0);
  const [dateCol, setDateCol] = useState(1);
  const [colB, setColB] = useState(1);
  const [lookupValue, setLookupValue] = useState("");
  const [ifNotFound, setIfNotFound] = useState("Not found");
  const [criteria, setCriteria] = useState<Criterion[]>([{ col: 0, op: "=", value: "" }]);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [offsetStart, setOffsetStart] = useState(0);
  const [offsetHeight, setOffsetHeight] = useState(10);
  const [offsetMode, setOffsetMode] = useState<"sum" | "average">("sum");
  const [letAgg, setLetAgg] = useState<"sum" | "average" | "count" | "max" | "min">("sum");
  const [letOp, setLetOp] = useState<CriteriaOp>(">");
  const [letValue, setLetValue] = useState("0");
  const [rate, setRate] = useState(10);
  const [periods, setPeriods] = useState(12);
  const [presentValue, setPresentValue] = useState(100000);
  const [period, setPeriod] = useState(1);
  const [decimals, setDecimals] = useState(2);
  const [roundMode, setRoundMode] = useState<RoundMode>("round");
  const [aggFn, setAggFn] = useState<AggFn>("sum");
  const [percentileK, setPercentileK] = useState(0.9);
  const [sep, setSep] = useState(" ");
  const [delimiter, setDelimiter] = useState(",");
  const [substrMode, setSubstrMode] = useState<SubstringMode>("left");
  const [charCount, setCharCount] = useState(3);
  const [midStart, setMidStart] = useState(1);
  const [caseMode, setCaseMode] = useState<CaseMode>("proper");
  const [oldText, setOldText] = useState("");
  const [newText, setNewText] = useState("");
  const [beforeAfter, setBeforeAfter] = useState<"before" | "after">("before");
  const [dateUnit, setDateUnit] = useState<"days" | "months" | "years">("days");
  const [endDateCol, setEndDateCol] = useState(1);
  const [datePartMode, setDatePartMode] = useState<"year" | "month" | "day">("year");
  const [monthsOffset, setMonthsOffset] = useState(0);
  const [dateFormat, setDateFormat] = useState<"iso" | "us" | "long">("iso");
  const [ifTrue, setIfTrue] = useState("Yes");
  const [ifFalse, setIfFalse] = useState("No");
  const [ifOp, setIfOp] = useState<CriteriaOp>(">");
  const [ifValue, setIfValue] = useState("0");
  const [bands, setBands] = useState<Band[]>([{ threshold: 80, label: "A" }, { threshold: 50, label: "B" }, { threshold: 0, label: "C" }]);
  const [switchMaps, setSwitchMaps] = useState<SwitchMap[]>([{ from: "", to: "" }]);
  const [andOrMode, setAndOrMode] = useState<"and" | "or">("and");
  const [fallback, setFallback] = useState("N/A");
  const [typeCheckMode, setTypeCheckMode] = useState<TypeCheckMode>("blank");
  const [sequenceStart, setSequenceStart] = useState(1);
  const [sequenceStep, setSequenceStep] = useState(1);
  const [secondTable, setSecondTable] = useState<DataTable | null>(null);
  const secondUploadRef = useRef<HTMLInputElement>(null);
  const [targetX, setTargetX] = useState(0);
  const [seasonality, setSeasonality] = useState(12);
  const [otherColumns, setOtherColumns] = useState("1,3,5");
  const [regexPattern, setRegexPattern] = useState("[^A-Za-z0-9 ]");
  const [regexReplacement, setRegexReplacement] = useState("");

  const table = up.table;
  const headers = table?.headers ?? [];
  const currentDef = EXCEL_FUNCTIONS.find((f) => f.key === fnKey)!;

  async function onSecondFile(file: File) {
    try { const t = await extractToTable(file); if (!t.headers.length) throw new Error("No readable data found."); setSecondTable(t); }
    catch { setSecondTable(null); }
  }

  function runInterpret() {
    if (!table) return;
    const interp = interpretInstruction(instruction, table);
    if (!interp) { setNote("Could not detect a function. Try naming one, e.g. XLOOKUP, SUMIFS, FILTER, unique, sort, PMT."); return; }
    setFnKey(interp.functionKey);
    setNote(interp.note);
    const g = interp.columnGuesses;
    if (g.sumCol !== undefined) { setSumCol(g.sumCol); setValueCol(g.sumCol); }
    if (g.dateCol !== undefined) setDateCol(g.dateCol);
    if (g.sortDir) setSortDir(g.sortDir);
    if (g.criterionValue) setCriteria((prev) => [{ ...prev[0], value: g.criterionValue! }, ...prev.slice(1)]);
    if (g.lookupValue) setLookupValue(g.lookupValue);
  }

  function updateCriterion(index: number, patch: Partial<Criterion>) {
    setCriteria((prev) => prev.map((c, i) => (i === index ? { ...c, ...patch } : c)));
  }
  function addCriterion() { setCriteria((prev) => [...prev, { col: 0, op: "=", value: "" }]); }
  function removeCriterion(index: number) { setCriteria((prev) => prev.filter((_, i) => i !== index)); }

  function run() {
    if (!table) return;
    let r: FunctionResult;
    switch (fnKey) {
      case "xlookup": r = xlookup(table, lookupValue, lookupCol, returnCol, ifNotFound); break;
      case "index-match": r = indexMatch(table, lookupValue, lookupCol, returnCol); break;
      case "vlookup": r = vlookup(table, lookupValue, lookupCol, returnCol); break;
      case "hlookup": r = hlookup(table, lookupValue, lookupCol, returnCol); break;
      case "offset-sum": r = offsetAggregate(table, valueCol, offsetStart, offsetHeight, offsetMode); break;
      case "sumifs": r = sumifs(table, sumCol, criteria); break;
      case "countifs": r = countifs(table, criteria); break;
      case "averageifs": r = averageifs(table, sumCol, criteria); break;
      case "correl": r = correl(table, valueCol, colB); break;
      case "stdev-s": r = stdevS(table, valueCol); break;
      case "var-s": r = varS(table, valueCol); break;
      case "forecast-linear": r = forecastLinear(table, valueCol, colB, targetX); break;
      case "forecast-ets": r = forecastETS(table, dateCol, valueCol, targetX, seasonality); break;
      case "maxifs": r = maxifs(table, valueCol, criteria); break;
      case "minifs": r = minifs(table, valueCol, criteria); break;
      case "sumproduct": r = sumproduct(table, valueCol, colB); break;
      case "round": r = roundColumn(table, valueCol, decimals, roundMode); break;
      case "aggregate": r = aggregateColumn(table, valueCol, aggFn); break;
      case "rank": r = rankColumn(table, valueCol, sortDir === "asc" ? "asc" : "desc"); break;
      case "percentile": r = percentileColumn(table, valueCol, percentileK); break;
      case "concat": r = concatColumns(table, valueCol, colB, sep); break;
      case "textsplit": r = textSplit(table, valueCol, delimiter); break;
      case "textjoin": r = textJoin(table, valueCol, delimiter); break;
      case "left-right-mid": r = substringColumn(table, valueCol, substrMode, charCount, midStart); break;
      case "trim-clean": r = trimCleanColumn(table, valueCol); break;
      case "upper-lower-proper": r = caseColumn(table, valueCol, caseMode); break;
      case "substitute": r = substituteColumn(table, valueCol, oldText, newText); break;
      case "len": r = lenColumn(table, valueCol); break;
      case "textbefore-after": r = textBeforeAfter(table, valueCol, delimiter, beforeAfter); break;
      case "datedif": r = datedif(table, valueCol, endDateCol, dateUnit); break;
      case "year-month-day": r = datePart(table, valueCol, datePartMode); break;
      case "weekday": r = weekdayColumn(table, valueCol); break;
      case "networkdays": r = networkdays(table, valueCol, endDateCol); break;
      case "eomonth": r = eomonth(table, valueCol, monthsOffset); break;
      case "text-date": r = formatDateColumn(table, valueCol, dateFormat); break;
      case "datevalue": r = dateValue(table, valueCol); break;
      case "timevalue": r = timeValue(table, valueCol); break;
      case "regexreplace": r = regexReplaceColumn(table, valueCol, regexPattern, regexReplacement); break;
      case "regextest": r = regexTestColumn(table, valueCol, regexPattern); break;
      case "regexextract": r = regexExtractColumn(table, valueCol, regexPattern); break;
      case "formulatext": r = formulaTextColumn(table, valueCol); break;
      case "error-type": r = errorTypeColumn(table, valueCol); break;
      case "if": r = ifColumn(table, valueCol, ifOp, ifValue, ifTrue, ifFalse); break;
      case "ifs": r = ifsBanding(table, valueCol, bands); break;
      case "and-or": r = andOrColumn(table, criteria, andOrMode); break;
      case "iferror": r = iferrorColumn(table, valueCol, fallback); break;
      case "switch": r = switchColumn(table, valueCol, switchMaps); break;
      case "is-type": r = typeCheckColumn(table, valueCol, typeCheckMode); break;
      case "indirect": r = indirectColumn(table, valueCol); break;
      case "sequence": r = sequenceColumn(table, sequenceStart, sequenceStep); break;
      case "filter": r = filterTableByCriterion(table, criteria[0]); break;
      case "unique": r = uniqueValues(table, valueCol); break;
      case "choosecols": r = chooseCols(table, otherColumns.split(",").map(v => Number(v.trim())).filter(Number.isFinite)); break;
      case "vstack": r = vstackTables(table, secondTable); break;
      case "tocol": r = tocol(table, valueCol, true); break;
      case "sort": r = sortTableByColumn(table, valueCol, sortDir); break;
      case "let": r = letCalc(table, valueCol, letAgg, letOp, letValue); break;
      case "xnpv": r = xnpv(table, valueCol, dateCol, rate); break;
      case "npv": r = npv(table, valueCol, rate); break;
      case "xirr": r = xirr(table, valueCol, dateCol); break;
      case "irr": r = irr(table, valueCol); break;
      case "pmt": r = pmt(rate, periods, presentValue); break;
      case "ipmt": r = ipmt(rate, period, periods, presentValue); break;
      default: r = { ok: false, message: "Unknown function." };
    }
    setResult(r);
  }

  const ColSelect = ({ value, onChange, label }: { value: number; onChange: (v: number) => void; label: string }) => (
    <label className="text-xs font-medium uppercase tracking-wide text-ink-3">
      {label}
      <select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal normal-case text-ink" value={value} onChange={(e) => onChange(Number(e.target.value))}>
        {headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
      </select>
    </label>
  );

  const CriteriaEditor = (
    <div className="rounded-xl border border-line p-3">
      <p className="text-xs font-medium uppercase tracking-wide text-ink-3">Criteria (all must match)</p>
      <div className="mt-2 space-y-2">
        {criteria.map((c, i) => (
          <div key={i} className="flex flex-wrap items-center gap-2 text-sm">
            <select className="rounded border border-line px-2 py-1" value={c.col} onChange={(e) => updateCriterion(i, { col: Number(e.target.value) })}>
              {headers.map((h, hi) => <option key={hi} value={hi}>{h}</option>)}
            </select>
            <select className="rounded border border-line px-2 py-1" value={c.op} onChange={(e) => updateCriterion(i, { op: e.target.value as CriteriaOp })}>
              {["=", "!=", ">", "<", ">=", "<=", "contains"].map((op) => <option key={op} value={op}>{op}</option>)}
            </select>
            <input className="min-w-32 flex-1 rounded border border-line px-2 py-1" placeholder="value" value={c.value} onChange={(e) => updateCriterion(i, { value: e.target.value })} />
            {criteria.length > 1 ? <button className="text-xs text-error" onClick={() => removeCriterion(i)}>Remove</button> : null}
          </div>
        ))}
      </div>
      {(fnKey !== "filter") ? <button className="mt-2 rounded-md border border-line px-3 py-1 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={addCriterion}>Add criterion</button> : null}
    </div>
  );

  return (
    <Panel centered title="Advanced Excel Functions" description="Transform a spreadsheet into a data analysis engine. Upload a file, then pick from 50+ Excel functions across lookup, conditional maths, statistics, text, date, logical, dynamic arrays, and finance, or type what you want in plain English. Results appear in seconds and each shows the equivalent Excel formula.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-error">{up.error}</p> : null}

      {table ? (
        <div className="mt-4 space-y-4">
          <div className="rounded-xl border border-line bg-canvas p-5">
            <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Describe what you want (natural language)</label>
            <div className="mt-1 flex flex-wrap gap-2">
              <input className="min-w-64 flex-1 rounded border border-line px-3 py-2 text-sm" placeholder='e.g. sum all amounts where region is "A", or find unique clients, or loan payment' value={instruction} onChange={(e) => setInstruction(e.target.value)} />
              <button className="rounded-md bg-ink px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={runInterpret}>Detect function</button>
            </div>
            {note ? <p className="mt-2 text-xs text-accent">{note}</p> : null}
          </div>

          <div className="rounded-xl border border-line bg-canvas p-5">
            <div className="flex flex-wrap items-center gap-3">
              <label className="text-sm font-medium">Function</label>
              <select className="rounded-lg border border-line px-3 py-1.5 text-sm" value={fnKey} onChange={(e) => { setFnKey(e.target.value as FunctionKey); setResult(null); }}>
                {(["Lookup & Reference", "Conditional Mathematics", "Statistical & Math", "Text", "Date & Time", "Logical", "Dynamic Arrays", "Advanced Logic", "Financial & Forecasting"] as const).map((cat) => (
                  <optgroup key={cat} label={cat}>
                    {EXCEL_FUNCTIONS.filter((f) => f.category === cat).map((f) => <option key={f.key} value={f.key}>{f.label}</option>)}
                  </optgroup>
                ))}
              </select>
            </div>
            <p className="mt-2 text-sm text-ink-2">{currentDef.description}</p>
            <p className="mt-1 font-mono text-xs text-ink-3">{currentDef.syntax}</p>

            <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {(fnKey === "xlookup" || fnKey === "index-match") ? (
                <>
                  <ColSelect label="Lookup column" value={lookupCol} onChange={setLookupCol} />
                  <ColSelect label="Return column" value={returnCol} onChange={setReturnCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Lookup value<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal normal-case" value={lookupValue} onChange={(e) => setLookupValue(e.target.value)} /></label>
                  {fnKey === "xlookup" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">If not found<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal normal-case" value={ifNotFound} onChange={(e) => setIfNotFound(e.target.value)} /></label> : null}
                </>
              ) : null}

              {fnKey === "offset-sum" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Start row (0-based)<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={offsetStart} onChange={(e) => setOffsetStart(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Height (rows)<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={offsetHeight} onChange={(e) => setOffsetHeight(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Mode<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={offsetMode} onChange={(e) => setOffsetMode(e.target.value as "sum" | "average")}><option value="sum">SUM</option><option value="average">AVERAGE</option></select></label>
                </>
              ) : null}

              {fnKey === "sumifs" ? <ColSelect label="Sum column" value={sumCol} onChange={setSumCol} /> : null}
              {fnKey === "sumproduct" ? (<><ColSelect label="Column A" value={valueCol} onChange={setValueCol} /><ColSelect label="Column B" value={colB} onChange={setColB} /></>) : null}
              {(fnKey === "unique" || fnKey === "sort" || fnKey === "stdev-s" || fnKey === "var-s" || fnKey === "datevalue" || fnKey === "timevalue" || fnKey === "regexreplace" || fnKey === "regextest" || fnKey === "regexextract" || fnKey === "formulatext" || fnKey === "error-type" || fnKey === "tocol") ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "correl" ? (<><ColSelect label="Column A" value={valueCol} onChange={setValueCol} /><ColSelect label="Column B" value={colB} onChange={setColB} /></>) : null}
              {fnKey === "forecast-linear" ? (<><ColSelect label="Known X" value={valueCol} onChange={setValueCol} /><ColSelect label="Known Y" value={colB} onChange={setColB} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Forecast X<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={targetX} onChange={e=>setTargetX(Number(e.target.value))} /></label></>) : null}
              {fnKey === "forecast-ets" ? (<><ColSelect label="Timeline" value={dateCol} onChange={setDateCol} /><ColSelect label="Values" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Target period<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={targetX} onChange={e=>setTargetX(Number(e.target.value))} /></label><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Seasonality<input type="number" min="1" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={seasonality} onChange={e=>setSeasonality(Number(e.target.value))} /></label></>) : null}
              {fnKey === "choosecols" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Columns (1-based, comma-separated)<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={otherColumns} onChange={e=>setOtherColumns(e.target.value)} /></label> : null}
              {(fnKey === "regexreplace" || fnKey === "regextest" || fnKey === "regexextract") ? (<><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Regex pattern<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal font-mono" value={regexPattern} onChange={e=>setRegexPattern(e.target.value)} /></label>{fnKey === "regexreplace" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Replacement<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={regexReplacement} onChange={e=>setRegexReplacement(e.target.value)} /></label> : null}</>) : null}
              {fnKey === "sort" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Direction<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={sortDir} onChange={(e) => setSortDir(e.target.value as "asc" | "desc")}><option value="asc">Ascending</option><option value="desc">Descending</option></select></label> : null}

              {fnKey === "let" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Aggregate<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={letAgg} onChange={(e) => setLetAgg(e.target.value as typeof letAgg)}><option value="sum">SUM</option><option value="average">AVERAGE</option><option value="count">COUNT</option><option value="max">MAX</option><option value="min">MIN</option></select></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Compare<div className="mt-1 flex gap-1"><select className="rounded border border-line px-1 py-1.5 text-sm font-normal" value={letOp} onChange={(e) => setLetOp(e.target.value as CriteriaOp)}>{[">", "<", ">=", "<=", "=", "!="].map((op) => <option key={op} value={op}>{op}</option>)}</select><input className="w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={letValue} onChange={(e) => setLetValue(e.target.value)} /></div></label>
                </>
              ) : null}

              {(fnKey === "xnpv" || fnKey === "xirr") ? (<><ColSelect label="Values column" value={valueCol} onChange={setValueCol} /><ColSelect label="Dates column" value={dateCol} onChange={setDateCol} /></>) : null}
              {(fnKey === "npv" || fnKey === "irr") ? <ColSelect label="Cash flow column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "vstack" ? (<div className="sm:col-span-2 lg:col-span-3 rounded-xl border border-line bg-surface p-3"><input ref={secondUploadRef} type="file" className="sr-only" accept=".csv,.xlsx,.xls" onChange={e=>{const file=e.target.files?.[0]; if(file) void onSecondFile(file); if(secondUploadRef.current) secondUploadRef.current.value="";}} /><button className="rounded-md border border-line px-3 py-1.5 text-xs font-medium" onClick={()=>secondUploadRef.current?.click()}>{secondTable ? `Second table: ${secondTable.rows.length} rows` : "Import second table for VSTACK"}</button></div>) : null}
              {fnKey === "xnpv" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Discount rate %<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={rate} onChange={(e) => setRate(Number(e.target.value))} /></label> : null}

              {(fnKey === "pmt" || fnKey === "ipmt") ? (
                <>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Annual rate %<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={rate} onChange={(e) => setRate(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Total periods<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={periods} onChange={(e) => setPeriods(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Present value<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={presentValue} onChange={(e) => setPresentValue(Number(e.target.value))} /></label>
                  {fnKey === "ipmt" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Payment period<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={period} onChange={(e) => setPeriod(Number(e.target.value))} /></label> : null}
                </>
              ) : null}

              {(fnKey === "vlookup" || fnKey === "hlookup") ? (
                <>
                  <ColSelect label="Lookup column" value={lookupCol} onChange={setLookupCol} />
                  <ColSelect label="Return column" value={returnCol} onChange={setReturnCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Lookup value<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal normal-case" value={lookupValue} onChange={(e) => setLookupValue(e.target.value)} /></label>
                </>
              ) : null}

              {fnKey === "averageifs" ? <ColSelect label="Average column" value={sumCol} onChange={setSumCol} /> : null}
              {(fnKey === "maxifs" || fnKey === "minifs") ? <ColSelect label={fnKey === "maxifs" ? "Max column" : "Min column"} value={valueCol} onChange={setValueCol} /> : null}

              {fnKey === "round" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Decimals<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={decimals} onChange={(e) => setDecimals(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Mode<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={roundMode} onChange={(e) => setRoundMode(e.target.value as RoundMode)}><option value="round">ROUND</option><option value="up">ROUNDUP</option><option value="down">ROUNDDOWN</option></select></label>
                </>
              ) : null}

              {fnKey === "aggregate" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Function<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={aggFn} onChange={(e) => setAggFn(e.target.value as AggFn)}><option value="sum">SUM</option><option value="average">AVERAGE</option><option value="min">MIN</option><option value="max">MAX</option><option value="median">MEDIAN</option><option value="count">COUNT</option><option value="stdev">STDEV</option><option value="var">VAR</option></select></label>
                </>
              ) : null}

              {fnKey === "rank" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Order<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={sortDir} onChange={(e) => setSortDir(e.target.value as "asc" | "desc")}><option value="desc">Highest first</option><option value="asc">Lowest first</option></select></label></>) : null}

              {fnKey === "percentile" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Percentile (0-1)<input type="number" step="0.05" min="0" max="1" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={percentileK} onChange={(e) => setPercentileK(Number(e.target.value))} /></label></>) : null}

              {fnKey === "concat" ? (<><ColSelect label="First column" value={valueCol} onChange={setValueCol} /><ColSelect label="Second column" value={colB} onChange={setColB} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Separator<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={sep} onChange={(e) => setSep(e.target.value)} /></label></>) : null}

              {(fnKey === "textsplit" || fnKey === "textjoin" || fnKey === "textbefore-after") ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Delimiter<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={delimiter} onChange={(e) => setDelimiter(e.target.value)} /></label></>) : null}
              {fnKey === "textbefore-after" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Before or after<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={beforeAfter} onChange={(e) => setBeforeAfter(e.target.value as "before" | "after")}><option value="before">Before</option><option value="after">After</option></select></label> : null}

              {fnKey === "left-right-mid" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Mode<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={substrMode} onChange={(e) => setSubstrMode(e.target.value as SubstringMode)}><option value="left">LEFT</option><option value="right">RIGHT</option><option value="mid">MID</option></select></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Characters<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={charCount} onChange={(e) => setCharCount(Number(e.target.value))} /></label>
                  {substrMode === "mid" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Start position<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={midStart} onChange={(e) => setMidStart(Number(e.target.value))} /></label> : null}
                </>
              ) : null}

              {(fnKey === "trim-clean" || fnKey === "len") ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "upper-lower-proper" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Case<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={caseMode} onChange={(e) => setCaseMode(e.target.value as CaseMode)}><option value="proper">Proper Case</option><option value="upper">UPPERCASE</option><option value="lower">lowercase</option></select></label></>) : null}
              {fnKey === "substitute" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Find<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={oldText} onChange={(e) => setOldText(e.target.value)} /></label><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Replace with<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={newText} onChange={(e) => setNewText(e.target.value)} /></label></>) : null}

              {(fnKey === "year-month-day") ? (<><ColSelect label="Date column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Part<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={datePartMode} onChange={(e) => setDatePartMode(e.target.value as "year" | "month" | "day")}><option value="year">Year</option><option value="month">Month</option><option value="day">Day</option></select></label></>) : null}
              {(fnKey === "weekday" || fnKey === "text-date" || fnKey === "eomonth") ? <ColSelect label="Date column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "text-date" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Format<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={dateFormat} onChange={(e) => setDateFormat(e.target.value as "iso" | "us" | "long")}><option value="iso">yyyy-mm-dd</option><option value="us">mm/dd/yyyy</option><option value="long">d Month yyyy</option></select></label> : null}
              {fnKey === "eomonth" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Months offset<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={monthsOffset} onChange={(e) => setMonthsOffset(Number(e.target.value))} /></label> : null}
              {(fnKey === "datedif" || fnKey === "networkdays") ? (<><ColSelect label="Start date column" value={valueCol} onChange={setValueCol} /><ColSelect label="End date column" value={endDateCol} onChange={setEndDateCol} /></>) : null}
              {fnKey === "datedif" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Unit<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={dateUnit} onChange={(e) => setDateUnit(e.target.value as "days" | "months" | "years")}><option value="days">Days</option><option value="months">Months</option><option value="years">Years</option></select></label> : null}

              {fnKey === "if" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Condition<div className="mt-1 flex gap-1"><select className="rounded border border-line px-1 py-1.5 text-sm font-normal" value={ifOp} onChange={(e) => setIfOp(e.target.value as CriteriaOp)}>{[">", "<", ">=", "<=", "=", "!=", "contains"].map((op) => <option key={op} value={op}>{op}</option>)}</select><input className="w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={ifValue} onChange={(e) => setIfValue(e.target.value)} /></div></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">If true<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={ifTrue} onChange={(e) => setIfTrue(e.target.value)} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">If false<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={ifFalse} onChange={(e) => setIfFalse(e.target.value)} /></label>
                </>
              ) : null}

              {fnKey === "ifs" ? <ColSelect label="Numeric column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "and-or" ? <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Logic<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={andOrMode} onChange={(e) => setAndOrMode(e.target.value as "and" | "or")}><option value="and">AND (all)</option><option value="or">OR (any)</option></select></label> : null}
              {fnKey === "iferror" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-medium uppercase tracking-wide text-ink-3">Fallback<input className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={fallback} onChange={(e) => setFallback(e.target.value)} /></label></>) : null}
              {fnKey === "switch" ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "is-type" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Check<select className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={typeCheckMode} onChange={(e) => setTypeCheckMode(e.target.value as TypeCheckMode)}><option value="blank">ISBLANK</option><option value="number">ISNUMBER</option><option value="text">ISTEXT</option></select></label>
                </>
              ) : null}
              {fnKey === "indirect" ? <ColSelect label="Selector column (holds a column name)" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "sequence" ? (
                <>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Start<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={sequenceStart} onChange={(e) => setSequenceStart(Number(e.target.value))} /></label>
                  <label className="text-xs font-medium uppercase tracking-wide text-ink-3">Step<input type="number" className="mt-1 block w-full rounded border border-line px-2 py-1.5 text-sm font-normal" value={sequenceStep} onChange={(e) => setSequenceStep(Number(e.target.value))} /></label>
                </>
              ) : null}
            </div>

            {fnKey === "ifs" ? (
              <div className="mt-4 rounded-xl border border-line p-3">
                <p className="text-xs font-medium uppercase tracking-wide text-ink-3">Bands (value greater than or equal to threshold gets the label)</p>
                <div className="mt-2 space-y-2">
                  {bands.map((b, i) => (
                    <div key={i} className="flex items-center gap-2 text-sm">
                      <input type="number" className="w-28 rounded border border-line px-2 py-1" value={b.threshold} onChange={(e) => setBands((prev) => prev.map((x, xi) => (xi === i ? { ...x, threshold: Number(e.target.value) } : x)))} />
                      <input className="flex-1 rounded border border-line px-2 py-1" value={b.label} onChange={(e) => setBands((prev) => prev.map((x, xi) => (xi === i ? { ...x, label: e.target.value } : x)))} />
                      {bands.length > 1 ? <button className="text-xs text-error" onClick={() => setBands((prev) => prev.filter((_, xi) => xi !== i))}>Remove</button> : null}
                    </div>
                  ))}
                </div>
                <button className="mt-2 rounded-md border border-line px-3 py-1 text-xs font-medium text-ink-2" onClick={() => setBands((prev) => [...prev, { threshold: 0, label: "" }])}>Add band</button>
              </div>
            ) : null}

            {fnKey === "switch" ? (
              <div className="mt-4 rounded-xl border border-line p-3">
                <p className="text-xs font-medium uppercase tracking-wide text-ink-3">Value mapping (from -&gt; to)</p>
                <div className="mt-2 space-y-2">
                  {switchMaps.map((m, i) => (
                    <div key={i} className="flex items-center gap-2 text-sm">
                      <input className="flex-1 rounded border border-line px-2 py-1" placeholder="from" value={m.from} onChange={(e) => setSwitchMaps((prev) => prev.map((x, xi) => (xi === i ? { ...x, from: e.target.value } : x)))} />
                      <input className="flex-1 rounded border border-line px-2 py-1" placeholder="to" value={m.to} onChange={(e) => setSwitchMaps((prev) => prev.map((x, xi) => (xi === i ? { ...x, to: e.target.value } : x)))} />
                      {switchMaps.length > 1 ? <button className="text-xs text-error" onClick={() => setSwitchMaps((prev) => prev.filter((_, xi) => xi !== i))}>Remove</button> : null}
                    </div>
                  ))}
                </div>
                <button className="mt-2 rounded-md border border-line px-3 py-1 text-xs font-medium text-ink-2" onClick={() => setSwitchMaps((prev) => [...prev, { from: "", to: "" }])}>Add mapping</button>
              </div>
            ) : null}

            {(fnKey === "sumifs" || fnKey === "countifs" || fnKey === "averageifs" || fnKey === "filter" || fnKey === "and-or") ? <div className="mt-4">{CriteriaEditor}</div> : null}

            <button className="mt-5 rounded-md bg-ink px-6 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint" onClick={run}>Run function</button>
          </div>

          {result ? (
            <div className={`rounded-xl border p-5 ${result.ok ? "border-line bg-canvas" : "border-line bg-surface"}`}>
              {result.scalar !== undefined ? <p className="text-3xl font-medium text-ink">{result.scalar}</p> : null}
              <p className={`mt-1 text-sm ${result.ok ? "text-ink-2" : "text-error"}`}>{result.message}</p>
              {result.formula ? <p className="mt-2 rounded-lg bg-slate-900 px-3 py-2 font-mono text-xs text-accent">{result.formula}</p> : null}
              {result.table ? (<><TablePreview table={result.table} /><DownloadButtons table={result.table} name="result" /></>) : null}
            </div>
          ) : null}
        </div>
      ) : null}
    </Panel>
  );
}

function TimesheetTool() {
  const [entries, setEntries] = useState<TimesheetEntry[]>([
    { date: new Date().toISOString().slice(0, 10), project: "", task: "", hours: 0, billable: true },
  ]);

  function update(index: number, patch: Partial<TimesheetEntry>) {
    setEntries((prev) => prev.map((e, i) => (i === index ? { ...e, ...patch } : e)));
  }
  function addRow() { setEntries((prev) => [...prev, { date: new Date().toISOString().slice(0, 10), project: "", task: "", hours: 0, billable: true }]); }
  function removeRow(index: number) { setEntries((prev) => prev.filter((_, i) => i !== index)); }

  const summary = useMemo(() => timesheetSummary(entries), [entries]);

  return (
    <Panel title="Timesheet Builder" description="Track hours across projects and generate a billable time report. Export a clean Excel timesheet for invoicing and analysis.">
      <div className="rounded-xl border border-line bg-canvas p-5">
        <div className="overflow-x-auto">
          <table className="min-w-full text-left text-sm">
            <thead className="text-xs uppercase text-ink-3"><tr><th className="px-2 py-2">Date</th><th className="px-2 py-2">Project</th><th className="px-2 py-2">Task</th><th className="px-2 py-2">Hours</th><th className="px-2 py-2">Billable</th><th /></tr></thead>
            <tbody>
              {entries.map((e, i) => (
                <tr key={i}>
                  <td className="px-2 py-1"><input type="date" className="rounded border border-line px-2 py-1" value={e.date} onChange={(ev) => update(i, { date: ev.target.value })} /></td>
                  <td className="px-2 py-1"><input className="rounded border border-line px-2 py-1" value={e.project} onChange={(ev) => update(i, { project: ev.target.value })} placeholder="Project" /></td>
                  <td className="px-2 py-1"><input className="rounded border border-line px-2 py-1" value={e.task} onChange={(ev) => update(i, { task: ev.target.value })} placeholder="Task" /></td>
                  <td className="px-2 py-1"><input type="number" step="0.25" className="w-20 rounded border border-line px-2 py-1" value={e.hours} onChange={(ev) => update(i, { hours: Number(ev.target.value) })} /></td>
                  <td className="px-2 py-1"><input type="checkbox" checked={e.billable} onChange={(ev) => update(i, { billable: ev.target.checked })} /></td>
                  <td className="px-2 py-1"><button className="text-xs text-error" onClick={() => removeRow(i)}>Remove</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <button className="mt-3 rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-slate-950" onClick={addRow}>Add row</button>
      </div>

      <div className="mt-4 grid gap-4 sm:grid-cols-3">
        <div className="rounded-xl border border-line bg-canvas p-5"><p className="text-3xl font-medium">{summary.totalHours}</p><p className="text-xs uppercase text-ink-3">Total hours</p></div>
        <div className="rounded-xl border border-line bg-canvas p-5"><p className="text-3xl font-medium text-ink-2">{summary.billableHours}</p><p className="text-xs uppercase text-ink-3">Billable</p></div>
        <div className="rounded-xl border border-line bg-canvas p-5"><p className="text-3xl font-medium text-ink-3">{summary.nonBillableHours}</p><p className="text-xs uppercase text-ink-3">Non-billable</p></div>
      </div>

      <button className="mt-4 rounded-md bg-accent-tint px-5 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => downloadMatrixExcel(timesheetToMatrix(entries), "timesheet.xlsx", "Timesheet")}>Download Excel timesheet</button>
    </Panel>
  );
}
