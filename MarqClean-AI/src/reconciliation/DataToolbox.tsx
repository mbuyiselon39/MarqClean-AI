import { useMemo, useRef, useState } from "react";
import { extractToTable, type DataTable } from "./engine";
import { GlobalHeader, ToolLaunchpad } from "./WorkspaceShell";
import {
  EXCEL_FUNCTIONS,
  aggregateColumn,
  andOrColumn,
  averageifs,
  caseColumn,
  concatColumns,
  countifs,
  datedif,
  datePart,
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
  uniqueValues,
  vlookup,
  weekdayColumn,
  xirr,
  xlookup,
  xnpv,
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

type ToolKey = "functions" | "merge" | "dedupe" | "fuzzy" | "clarity" | "explorer" | "predict" | "extract" | "timesheet";

const TOOLS: Array<{ key: ToolKey; label: string; blurb: string; icon: string }> = [
  { key: "functions", label: "Advanced Excel Functions", blurb: "Lookup, text, date, logical, stats, and finance functions, automated.", icon: "📊" },
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
          <span className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-[#5B4FFF]/20 to-[#8B5CF6]/20 text-xl ring-1 ring-inset ring-white/10">🧰</span>
          <div>
            <h2 className="mc-display text-xl font-bold text-white">Data Toolbox</h2>
            <p className="text-xs text-[#b0bacb]">Advanced Excel functions, cleansing, matching and transformation. All tools free, browser-based.</p>
          </div>
        </div>
        <button className="ws-btn-secondary rounded-full px-4 py-2 text-sm font-medium" onClick={onExit}>Back to MarqClean AI</button>
      </div>

      <div className="relative mx-auto max-w-[110rem] px-5 pb-16 pt-4 lg:px-8">
        <div className="mb-2 flex items-center justify-between">
          <p className="mc-mono text-[11px] uppercase tracking-[0.2em] text-[#8a94a8]">Tools Available ({TOOLS.length})</p>
        </div>
        <ToolLaunchpad tools={TOOLS} active={tool} onSelect={setTool} />

        <main id="main-content" className="mt-6 min-w-0">
          {tool === "functions" ? <FunctionsTool /> : null}
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
        <h1 className="text-3xl font-semibold tracking-[-0.03em]">{title}</h1>
        <p className={centered ? "mt-2 text-justify leading-7 text-slate-600 [hyphens:auto]" : "mt-2 max-w-3xl text-slate-600"}>{description}</p>
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
      className="rounded-2xl border border-dashed border-slate-300 bg-white p-5 text-center transition hover:border-indigo-500"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) onFile(f); }}
    >
      <input ref={refEl} className="sr-only" type="file" accept=".pdf,.xlsx,.xls,.csv,.txt,.docx,.doc" onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); if (refEl.current) refEl.current.value = ""; }} />
      <h3 className="text-sm font-semibold text-slate-900">{label}</h3>
      <p className="mt-1 text-xs text-slate-500">{table ? `${table.sourceName} - ${table.rows.length} rows` : "Drop or choose a CSV, Excel, PDF, or Word file."}</p>
      <button className="mt-3 rounded-full bg-slate-950 px-4 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={() => refEl.current?.click()}>
        {table ? "Replace" : "Choose file"}
      </button>
    </div>
  );
}

function TablePreview({ table, max = 8 }: { table: DataTable; max?: number }) {
  return (
    <div className="mt-4 overflow-auto rounded-xl border border-slate-200 bg-white" style={{ maxHeight: "22rem" }}>
      <table className="min-w-full text-left text-sm">
        <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
          <tr>{table.headers.map((h, i) => <th key={i} className="whitespace-nowrap px-3 py-2 font-semibold">{h}</th>)}</tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {table.rows.slice(0, max).map((row, r) => (
            <tr key={r}>{table.headers.map((_, c) => <td key={c} className="max-w-48 truncate px-3 py-2 text-slate-700">{row[c] || "-"}</td>)}</tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function DownloadButtons({ table, name }: { table: DataTable; name: string }) {
  return (
    <div className="mt-4 flex flex-wrap gap-2">
      <button className="rounded-full bg-indigo-500 px-4 py-2 text-xs font-semibold text-slate-950 transition hover:bg-violet-300" onClick={() => downloadTableExcel(table, `${name}.xlsx`, name)}>Download Excel</button>
      <button className="rounded-full border border-slate-300 px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => downloadTableCsv(table, `${name}.csv`)}>Download CSV</button>
    </div>
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
      {a.error || b.error ? <p className="mt-3 text-sm text-red-700">{a.error || b.error}</p> : null}
      {a.table && b.table ? (
        <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <label className="font-semibold">Mode</label>
            <select className="rounded border border-slate-300 px-3 py-1.5" value={mode} onChange={(e) => setMode(e.target.value as "append" | "join")}>
              <option value="append">Append (stack rows)</option>
              <option value="join">Join (match by key)</option>
            </select>
            {mode === "join" ? (
              <>
                <label>First key</label>
                <select className="rounded border border-slate-300 px-2 py-1" value={leftKey} onChange={(e) => setLeftKey(Number(e.target.value))}>
                  {a.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
                <label>Second key</label>
                <select className="rounded border border-slate-300 px-2 py-1" value={rightKey} onChange={(e) => setRightKey(Number(e.target.value))}>
                  {b.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
              </>
            ) : null}
            <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={run}>Merge</button>
          </div>
        </div>
      ) : null}
      {result ? (<><p className="mt-4 text-sm text-slate-600">{result.rows.length} rows, {result.headers.length} columns.</p><TablePreview table={result} /><DownloadButtons table={result} name="merged" /></>) : null}
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
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-5">
          <h3 className="text-sm font-semibold">Exact duplicate columns (leave all unchecked for whole-row match)</h3>
          <div className="mt-2 flex flex-wrap gap-2 text-sm">
            {up.table.headers.map((h, i) => (
              <label key={i} className="flex items-center gap-1 rounded border border-slate-200 px-2 py-1"><input type="checkbox" checked={selected.has(i)} onChange={() => toggle(i)} />{h}</label>
            ))}
          </div>
          <button className="mt-3 rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={() => setResult(removeDuplicatesByColumns(up.table!, Array.from(selected)))}>Remove duplicates</button>

          <div className="mt-6 border-t border-slate-200 pt-4">
            <h3 className="text-sm font-semibold">Fuzzy duplicate finder</h3>
            <div className="mt-2 flex flex-wrap items-center gap-3 text-sm">
              <select className="rounded border border-slate-300 px-2 py-1" value={fuzzyCol} onChange={(e) => setFuzzyCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <label>Similarity</label>
              <input type="range" min={0.6} max={0.98} step={0.01} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
              <span>{Math.round(threshold * 100)}%</span>
              <button className="rounded-full border border-slate-300 px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => setFuzzyPairs(findFuzzyDuplicates(up.table!, fuzzyCol, threshold))}>Find near-duplicates</button>
            </div>
            {fuzzyPairs ? (
              <div className="mt-3 max-h-64 overflow-auto text-sm">
                <p className="text-slate-600">{fuzzyPairs.length} likely duplicate pairs</p>
                <table className="mt-2 min-w-full text-left"><tbody className="divide-y divide-slate-100">
                  {fuzzyPairs.map((p, i) => <tr key={i}><td className="py-1 pr-4">{p.valueA}</td><td className="py-1 pr-4">{p.valueB}</td><td className="py-1 text-indigo-700">{Math.round(p.score * 100)}%</td></tr>)}
                </tbody></table>
              </div>
            ) : null}
          </div>
        </div>
      ) : null}
      {result ? (<><p className="mt-4 text-sm text-slate-600">Removed {result.removed} duplicates. {result.table.rows.length} rows remain.</p><TablePreview table={result.table} /><DownloadButtons table={result.table} name="deduplicated" /></>) : null}
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
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <label>Search column</label>
            <select className="rounded border border-slate-300 px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <input className="min-w-48 flex-1 rounded border border-slate-300 px-3 py-1.5" placeholder="Search name or address" value={query} onChange={(e) => setQuery(e.target.value)} />
            <label>Min similarity</label>
            <input type="range" min={0.4} max={0.95} step={0.01} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
            <span>{Math.round(threshold * 100)}%</span>
            <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={() => setResults(fuzzySearch(up.table!, col, query, threshold))}>Search</button>
          </div>
          {results ? (
            <div className="mt-3 max-h-72 overflow-auto text-sm">
              <p className="text-slate-600">{results.length} matches</p>
              <table className="mt-2 min-w-full text-left"><tbody className="divide-y divide-slate-100">
                {results.map((r, i) => <tr key={i}><td className="py-1 pr-4">{r.value}</td><td className="py-1 text-indigo-700">{Math.round(r.score * 100)}%</td></tr>)}
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
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 space-y-4">
          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={() => setReport(computeClarityScore(up.table!))}>Compute clarity score</button>
            {report ? (
              <div className="mt-4">
                <div className="flex items-center gap-4">
                  <div className="flex h-20 w-20 items-center justify-center rounded-full bg-indigo-100 text-2xl font-bold text-indigo-800">{report.score}</div>
                  <div className="text-sm text-slate-600">
                    <p>{report.totalRows} rows, {report.filledCells}/{report.totalCells} cells filled</p>
                    <p>{report.duplicateRows} duplicate rows</p>
                  </div>
                </div>
                {report.issues.length ? (
                  <div className="mt-4 max-h-48 overflow-auto text-sm">
                    <table className="min-w-full text-left"><thead className="text-xs uppercase text-slate-500"><tr><th className="py-1 pr-4">Column</th><th className="py-1 pr-4">Issue</th><th className="py-1">Count</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">{report.issues.map((iss, i) => <tr key={i}><td className="py-1 pr-4">{iss.column}</td><td className="py-1 pr-4 text-amber-700">{iss.issue}</td><td className="py-1">{iss.count}</td></tr>)}</tbody></table>
                  </div>
                ) : null}
              </div>
            ) : null}
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <h3 className="text-sm font-semibold">Auto clean actions</h3>
            <div className="mt-2 flex flex-wrap items-center gap-3 text-sm">
              <select className="rounded border border-slate-300 px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => { const r = smartFill(up.table!, col); recompute(r.table); }}>Fill gaps (mode)</button>
              <select className="rounded border border-slate-300 px-2 py-1" value={fmt} onChange={(e) => setFmt(e.target.value as FormatMode)}>
                <option value="title">Title Case</option><option value="upper">UPPERCASE</option><option value="lower">lowercase</option><option value="trim">Trim</option>
              </select>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => recompute(autoFormatColumn(up.table!, col, fmt))}>Auto format</button>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => { const r = removeDuplicatesByColumns(up.table!, []); recompute(r.table); }}>Remove duplicate rows</button>
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
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 grid gap-4 lg:grid-cols-2">
          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <h3 className="text-sm font-semibold">Filter</h3>
            <div className="mt-2 flex flex-wrap items-center gap-2 text-sm">
              <select className="rounded border border-slate-300 px-2 py-1" value={col} onChange={(e) => setCol(Number(e.target.value))}>
                {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
              </select>
              <select className="rounded border border-slate-300 px-2 py-1" value={op} onChange={(e) => setOp(e.target.value as FilterOp)}>
                <option value="contains">contains</option><option value="equals">equals</option><option value="gt">greater than</option><option value="lt">less than</option><option value="nonempty">is not empty</option><option value="empty">is empty</option>
              </select>
              <input className="rounded border border-slate-300 px-2 py-1" value={value} onChange={(e) => setValue(e.target.value)} placeholder="value" />
              <button className="rounded-full bg-slate-950 px-4 py-1.5 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={() => setView(filterTable(up.table!, col, op, value))}>Apply</button>
            </div>
            {view ? (<><p className="mt-3 text-sm text-slate-600">{view.rows.length} matching rows</p><DownloadButtons table={view} name="filtered" /></>) : null}
          </div>
          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <h3 className="text-sm font-semibold">Distribution of {up.table.headers[col]}</h3>
            <div className="mt-3 space-y-2">
              {dist.map((d) => (
                <div key={d.label}>
                  <div className="flex justify-between text-xs text-slate-600"><span className="truncate pr-2">{d.label}</span><span className="font-semibold">{d.count}</span></div>
                  <div className="mt-1 h-2 rounded-full bg-slate-100"><div className="h-2 rounded-full bg-indigo-500" style={{ width: `${Math.max((d.count / distMax) * 100, 2)}%` }} /></div>
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
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}
      {up.table ? (
        <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-5">
          <div className="flex flex-wrap items-center gap-3 text-sm">
            <select className="rounded border border-slate-300 px-2 py-1" value={mode} onChange={(e) => setMode(e.target.value as "regression" | "classify")}>
              <option value="regression">Predict number (regression)</option>
              <option value="classify">Predict category (classification)</option>
            </select>
            <label>Feature</label>
            <select className="rounded border border-slate-300 px-2 py-1" value={feature} onChange={(e) => setFeature(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <label>Target</label>
            <select className="rounded border border-slate-300 px-2 py-1" value={target} onChange={(e) => setTarget(Number(e.target.value))}>
              {up.table.headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
            </select>
            <input className="w-28 rounded border border-slate-300 px-2 py-1" placeholder="feature value" value={input} onChange={(e) => setInput(e.target.value)} />
            <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={run}>Predict</button>
          </div>
          {output ? <p className="mt-4 rounded-xl bg-indigo-50 px-4 py-3 text-sm font-medium text-indigo-800">{output}</p> : null}
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
      <textarea className="h-40 w-full rounded-2xl border border-slate-300 bg-white px-4 py-3 text-sm" placeholder="Paste HTML containing a <table> or list here" value={html} onChange={(e) => setHtml(e.target.value)} />
      <div className="mt-3 flex gap-2">
        <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={run}>Extract table</button>
      </div>
      {error ? <p className="mt-3 text-sm text-red-700">{error}</p> : null}
      {table ? (<><p className="mt-4 text-sm text-slate-600">{table.rows.length} rows, {table.headers.length} columns.</p><TablePreview table={table} /><DownloadButtons table={table} name="extracted" /></>) : null}
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

  const table = up.table;
  const headers = table?.headers ?? [];
  const currentDef = EXCEL_FUNCTIONS.find((f) => f.key === fnKey)!;

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
      case "if": r = ifColumn(table, valueCol, ifOp, ifValue, ifTrue, ifFalse); break;
      case "ifs": r = ifsBanding(table, valueCol, bands); break;
      case "and-or": r = andOrColumn(table, criteria, andOrMode); break;
      case "iferror": r = iferrorColumn(table, valueCol, fallback); break;
      case "switch": r = switchColumn(table, valueCol, switchMaps); break;
      case "is-type": r = typeCheckColumn(table, valueCol, typeCheckMode); break;
      case "filter": r = filterTableByCriterion(table, criteria[0]); break;
      case "unique": r = uniqueValues(table, valueCol); break;
      case "sort": r = sortTableByColumn(table, valueCol, sortDir); break;
      case "let": r = letCalc(table, valueCol, letAgg, letOp, letValue); break;
      case "xnpv": r = xnpv(table, valueCol, dateCol, rate); break;
      case "xirr": r = xirr(table, valueCol, dateCol); break;
      case "pmt": r = pmt(rate, periods, presentValue); break;
      case "ipmt": r = ipmt(rate, period, periods, presentValue); break;
      default: r = { ok: false, message: "Unknown function." };
    }
    setResult(r);
  }

  const ColSelect = ({ value, onChange, label }: { value: number; onChange: (v: number) => void; label: string }) => (
    <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">
      {label}
      <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal normal-case text-slate-800" value={value} onChange={(e) => onChange(Number(e.target.value))}>
        {headers.map((h, i) => <option key={i} value={i}>{h}</option>)}
      </select>
    </label>
  );

  const CriteriaEditor = (
    <div className="rounded-xl border border-slate-200 p-3">
      <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Criteria (all must match)</p>
      <div className="mt-2 space-y-2">
        {criteria.map((c, i) => (
          <div key={i} className="flex flex-wrap items-center gap-2 text-sm">
            <select className="rounded border border-slate-300 px-2 py-1" value={c.col} onChange={(e) => updateCriterion(i, { col: Number(e.target.value) })}>
              {headers.map((h, hi) => <option key={hi} value={hi}>{h}</option>)}
            </select>
            <select className="rounded border border-slate-300 px-2 py-1" value={c.op} onChange={(e) => updateCriterion(i, { op: e.target.value as CriteriaOp })}>
              {["=", "!=", ">", "<", ">=", "<=", "contains"].map((op) => <option key={op} value={op}>{op}</option>)}
            </select>
            <input className="min-w-32 flex-1 rounded border border-slate-300 px-2 py-1" placeholder="value" value={c.value} onChange={(e) => updateCriterion(i, { value: e.target.value })} />
            {criteria.length > 1 ? <button className="text-xs text-red-600" onClick={() => removeCriterion(i)}>Remove</button> : null}
          </div>
        ))}
      </div>
      {(fnKey !== "filter") ? <button className="mt-2 rounded-full border border-slate-300 px-3 py-1 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={addCriterion}>Add criterion</button> : null}
    </div>
  );

  return (
    <Panel centered title="Advanced Excel Functions" description="Transform a spreadsheet into a data analysis engine. Upload a file, then pick from 40+ Excel functions across lookup, conditional maths, statistics, text, date, logical, dynamic arrays, and finance, or type what you want in plain English. Results appear in seconds and each shows the equivalent Excel formula.">
      <UploadBox label="Upload file" table={up.table} refEl={up.ref} onFile={up.onFile} />
      {up.error ? <p className="mt-3 text-sm text-red-700">{up.error}</p> : null}

      {table ? (
        <div className="mt-4 space-y-4">
          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Describe what you want (natural language)</label>
            <div className="mt-1 flex flex-wrap gap-2">
              <input className="min-w-64 flex-1 rounded border border-slate-300 px-3 py-2 text-sm" placeholder='e.g. sum all amounts where region is "A", or find unique clients, or loan payment' value={instruction} onChange={(e) => setInstruction(e.target.value)} />
              <button className="rounded-full bg-slate-950 px-5 py-2 text-xs font-semibold text-white transition hover:bg-indigo-700" onClick={runInterpret}>Detect function</button>
            </div>
            {note ? <p className="mt-2 text-xs text-indigo-700">{note}</p> : null}
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <div className="flex flex-wrap items-center gap-3">
              <label className="text-sm font-semibold">Function</label>
              <select className="rounded-lg border border-slate-300 px-3 py-1.5 text-sm" value={fnKey} onChange={(e) => { setFnKey(e.target.value as FunctionKey); setResult(null); }}>
                {(["Lookup & Reference", "Conditional Mathematics", "Statistical & Math", "Text", "Date & Time", "Logical", "Dynamic Arrays", "Advanced Logic", "Financial & Forecasting"] as const).map((cat) => (
                  <optgroup key={cat} label={cat}>
                    {EXCEL_FUNCTIONS.filter((f) => f.category === cat).map((f) => <option key={f.key} value={f.key}>{f.label}</option>)}
                  </optgroup>
                ))}
              </select>
            </div>
            <p className="mt-2 text-sm text-slate-600">{currentDef.description}</p>
            <p className="mt-1 font-mono text-xs text-slate-400">{currentDef.syntax}</p>

            <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {(fnKey === "xlookup" || fnKey === "index-match") ? (
                <>
                  <ColSelect label="Lookup column" value={lookupCol} onChange={setLookupCol} />
                  <ColSelect label="Return column" value={returnCol} onChange={setReturnCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Lookup value<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal normal-case" value={lookupValue} onChange={(e) => setLookupValue(e.target.value)} /></label>
                  {fnKey === "xlookup" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">If not found<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal normal-case" value={ifNotFound} onChange={(e) => setIfNotFound(e.target.value)} /></label> : null}
                </>
              ) : null}

              {fnKey === "offset-sum" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Start row (0-based)<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={offsetStart} onChange={(e) => setOffsetStart(Number(e.target.value))} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Height (rows)<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={offsetHeight} onChange={(e) => setOffsetHeight(Number(e.target.value))} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Mode<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={offsetMode} onChange={(e) => setOffsetMode(e.target.value as "sum" | "average")}><option value="sum">SUM</option><option value="average">AVERAGE</option></select></label>
                </>
              ) : null}

              {fnKey === "sumifs" ? <ColSelect label="Sum column" value={sumCol} onChange={setSumCol} /> : null}
              {fnKey === "sumproduct" ? (<><ColSelect label="Column A" value={valueCol} onChange={setValueCol} /><ColSelect label="Column B" value={colB} onChange={setColB} /></>) : null}
              {(fnKey === "unique" || fnKey === "sort") ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "sort" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Direction<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={sortDir} onChange={(e) => setSortDir(e.target.value as "asc" | "desc")}><option value="asc">Ascending</option><option value="desc">Descending</option></select></label> : null}

              {fnKey === "let" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Aggregate<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={letAgg} onChange={(e) => setLetAgg(e.target.value as typeof letAgg)}><option value="sum">SUM</option><option value="average">AVERAGE</option><option value="count">COUNT</option><option value="max">MAX</option><option value="min">MIN</option></select></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Compare<div className="mt-1 flex gap-1"><select className="rounded border border-slate-300 px-1 py-1.5 text-sm font-normal" value={letOp} onChange={(e) => setLetOp(e.target.value as CriteriaOp)}>{[">", "<", ">=", "<=", "=", "!="].map((op) => <option key={op} value={op}>{op}</option>)}</select><input className="w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={letValue} onChange={(e) => setLetValue(e.target.value)} /></div></label>
                </>
              ) : null}

              {(fnKey === "xnpv" || fnKey === "xirr") ? (<><ColSelect label="Values column" value={valueCol} onChange={setValueCol} /><ColSelect label="Dates column" value={dateCol} onChange={setDateCol} /></>) : null}
              {fnKey === "xnpv" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Discount rate %<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={rate} onChange={(e) => setRate(Number(e.target.value))} /></label> : null}

              {(fnKey === "pmt" || fnKey === "ipmt") ? (
                <>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Annual rate %<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={rate} onChange={(e) => setRate(Number(e.target.value))} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Total periods<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={periods} onChange={(e) => setPeriods(Number(e.target.value))} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Present value<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={presentValue} onChange={(e) => setPresentValue(Number(e.target.value))} /></label>
                  {fnKey === "ipmt" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Payment period<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={period} onChange={(e) => setPeriod(Number(e.target.value))} /></label> : null}
                </>
              ) : null}

              {(fnKey === "vlookup" || fnKey === "hlookup") ? (
                <>
                  <ColSelect label="Lookup column" value={lookupCol} onChange={setLookupCol} />
                  <ColSelect label="Return column" value={returnCol} onChange={setReturnCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Lookup value<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal normal-case" value={lookupValue} onChange={(e) => setLookupValue(e.target.value)} /></label>
                </>
              ) : null}

              {fnKey === "averageifs" ? <ColSelect label="Average column" value={sumCol} onChange={setSumCol} /> : null}

              {fnKey === "round" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Decimals<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={decimals} onChange={(e) => setDecimals(Number(e.target.value))} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Mode<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={roundMode} onChange={(e) => setRoundMode(e.target.value as RoundMode)}><option value="round">ROUND</option><option value="up">ROUNDUP</option><option value="down">ROUNDDOWN</option></select></label>
                </>
              ) : null}

              {fnKey === "aggregate" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Function<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={aggFn} onChange={(e) => setAggFn(e.target.value as AggFn)}><option value="sum">SUM</option><option value="average">AVERAGE</option><option value="min">MIN</option><option value="max">MAX</option><option value="median">MEDIAN</option><option value="count">COUNT</option><option value="stdev">STDEV</option><option value="var">VAR</option></select></label>
                </>
              ) : null}

              {fnKey === "rank" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Order<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={sortDir} onChange={(e) => setSortDir(e.target.value as "asc" | "desc")}><option value="desc">Highest first</option><option value="asc">Lowest first</option></select></label></>) : null}

              {fnKey === "percentile" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Percentile (0-1)<input type="number" step="0.05" min="0" max="1" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={percentileK} onChange={(e) => setPercentileK(Number(e.target.value))} /></label></>) : null}

              {fnKey === "concat" ? (<><ColSelect label="First column" value={valueCol} onChange={setValueCol} /><ColSelect label="Second column" value={colB} onChange={setColB} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Separator<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={sep} onChange={(e) => setSep(e.target.value)} /></label></>) : null}

              {(fnKey === "textsplit" || fnKey === "textjoin" || fnKey === "textbefore-after") ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Delimiter<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={delimiter} onChange={(e) => setDelimiter(e.target.value)} /></label></>) : null}
              {fnKey === "textbefore-after" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Before or after<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={beforeAfter} onChange={(e) => setBeforeAfter(e.target.value as "before" | "after")}><option value="before">Before</option><option value="after">After</option></select></label> : null}

              {fnKey === "left-right-mid" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Mode<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={substrMode} onChange={(e) => setSubstrMode(e.target.value as SubstringMode)}><option value="left">LEFT</option><option value="right">RIGHT</option><option value="mid">MID</option></select></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Characters<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={charCount} onChange={(e) => setCharCount(Number(e.target.value))} /></label>
                  {substrMode === "mid" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Start position<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={midStart} onChange={(e) => setMidStart(Number(e.target.value))} /></label> : null}
                </>
              ) : null}

              {(fnKey === "trim-clean" || fnKey === "len") ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "upper-lower-proper" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Case<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={caseMode} onChange={(e) => setCaseMode(e.target.value as CaseMode)}><option value="proper">Proper Case</option><option value="upper">UPPERCASE</option><option value="lower">lowercase</option></select></label></>) : null}
              {fnKey === "substitute" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Find<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={oldText} onChange={(e) => setOldText(e.target.value)} /></label><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Replace with<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={newText} onChange={(e) => setNewText(e.target.value)} /></label></>) : null}

              {(fnKey === "year-month-day") ? (<><ColSelect label="Date column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Part<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={datePartMode} onChange={(e) => setDatePartMode(e.target.value as "year" | "month" | "day")}><option value="year">Year</option><option value="month">Month</option><option value="day">Day</option></select></label></>) : null}
              {(fnKey === "weekday" || fnKey === "text-date" || fnKey === "eomonth") ? <ColSelect label="Date column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "text-date" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Format<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={dateFormat} onChange={(e) => setDateFormat(e.target.value as "iso" | "us" | "long")}><option value="iso">yyyy-mm-dd</option><option value="us">mm/dd/yyyy</option><option value="long">d Month yyyy</option></select></label> : null}
              {fnKey === "eomonth" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Months offset<input type="number" className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={monthsOffset} onChange={(e) => setMonthsOffset(Number(e.target.value))} /></label> : null}
              {(fnKey === "datedif" || fnKey === "networkdays") ? (<><ColSelect label="Start date column" value={valueCol} onChange={setValueCol} /><ColSelect label="End date column" value={endDateCol} onChange={setEndDateCol} /></>) : null}
              {fnKey === "datedif" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Unit<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={dateUnit} onChange={(e) => setDateUnit(e.target.value as "days" | "months" | "years")}><option value="days">Days</option><option value="months">Months</option><option value="years">Years</option></select></label> : null}

              {fnKey === "if" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Condition<div className="mt-1 flex gap-1"><select className="rounded border border-slate-300 px-1 py-1.5 text-sm font-normal" value={ifOp} onChange={(e) => setIfOp(e.target.value as CriteriaOp)}>{[">", "<", ">=", "<=", "=", "!=", "contains"].map((op) => <option key={op} value={op}>{op}</option>)}</select><input className="w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={ifValue} onChange={(e) => setIfValue(e.target.value)} /></div></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">If true<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={ifTrue} onChange={(e) => setIfTrue(e.target.value)} /></label>
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">If false<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={ifFalse} onChange={(e) => setIfFalse(e.target.value)} /></label>
                </>
              ) : null}

              {fnKey === "ifs" ? <ColSelect label="Numeric column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "and-or" ? <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Logic<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={andOrMode} onChange={(e) => setAndOrMode(e.target.value as "and" | "or")}><option value="and">AND (all)</option><option value="or">OR (any)</option></select></label> : null}
              {fnKey === "iferror" ? (<><ColSelect label="Column" value={valueCol} onChange={setValueCol} /><label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Fallback<input className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={fallback} onChange={(e) => setFallback(e.target.value)} /></label></>) : null}
              {fnKey === "switch" ? <ColSelect label="Column" value={valueCol} onChange={setValueCol} /> : null}
              {fnKey === "is-type" ? (
                <>
                  <ColSelect label="Column" value={valueCol} onChange={setValueCol} />
                  <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Check<select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={typeCheckMode} onChange={(e) => setTypeCheckMode(e.target.value as TypeCheckMode)}><option value="blank">ISBLANK</option><option value="number">ISNUMBER</option><option value="text">ISTEXT</option></select></label>
                </>
              ) : null}
            </div>

            {fnKey === "ifs" ? (
              <div className="mt-4 rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Bands (value greater than or equal to threshold gets the label)</p>
                <div className="mt-2 space-y-2">
                  {bands.map((b, i) => (
                    <div key={i} className="flex items-center gap-2 text-sm">
                      <input type="number" className="w-28 rounded border border-slate-300 px-2 py-1" value={b.threshold} onChange={(e) => setBands((prev) => prev.map((x, xi) => (xi === i ? { ...x, threshold: Number(e.target.value) } : x)))} />
                      <input className="flex-1 rounded border border-slate-300 px-2 py-1" value={b.label} onChange={(e) => setBands((prev) => prev.map((x, xi) => (xi === i ? { ...x, label: e.target.value } : x)))} />
                      {bands.length > 1 ? <button className="text-xs text-red-600" onClick={() => setBands((prev) => prev.filter((_, xi) => xi !== i))}>Remove</button> : null}
                    </div>
                  ))}
                </div>
                <button className="mt-2 rounded-full border border-slate-300 px-3 py-1 text-xs font-semibold text-slate-700" onClick={() => setBands((prev) => [...prev, { threshold: 0, label: "" }])}>Add band</button>
              </div>
            ) : null}

            {fnKey === "switch" ? (
              <div className="mt-4 rounded-xl border border-slate-200 p-3">
                <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">Value mapping (from -&gt; to)</p>
                <div className="mt-2 space-y-2">
                  {switchMaps.map((m, i) => (
                    <div key={i} className="flex items-center gap-2 text-sm">
                      <input className="flex-1 rounded border border-slate-300 px-2 py-1" placeholder="from" value={m.from} onChange={(e) => setSwitchMaps((prev) => prev.map((x, xi) => (xi === i ? { ...x, from: e.target.value } : x)))} />
                      <input className="flex-1 rounded border border-slate-300 px-2 py-1" placeholder="to" value={m.to} onChange={(e) => setSwitchMaps((prev) => prev.map((x, xi) => (xi === i ? { ...x, to: e.target.value } : x)))} />
                      {switchMaps.length > 1 ? <button className="text-xs text-red-600" onClick={() => setSwitchMaps((prev) => prev.filter((_, xi) => xi !== i))}>Remove</button> : null}
                    </div>
                  ))}
                </div>
                <button className="mt-2 rounded-full border border-slate-300 px-3 py-1 text-xs font-semibold text-slate-700" onClick={() => setSwitchMaps((prev) => [...prev, { from: "", to: "" }])}>Add mapping</button>
              </div>
            ) : null}

            {(fnKey === "sumifs" || fnKey === "countifs" || fnKey === "averageifs" || fnKey === "filter" || fnKey === "and-or") ? <div className="mt-4">{CriteriaEditor}</div> : null}

            <button className="mt-5 rounded-full bg-slate-950 px-6 py-3 text-sm font-semibold text-white transition hover:bg-indigo-700" onClick={run}>Run function</button>
          </div>

          {result ? (
            <div className={`rounded-2xl border p-5 ${result.ok ? "border-slate-200 bg-white" : "border-red-200 bg-red-50"}`}>
              {result.scalar !== undefined ? <p className="text-3xl font-semibold text-slate-950">{result.scalar}</p> : null}
              <p className={`mt-1 text-sm ${result.ok ? "text-slate-600" : "text-red-700"}`}>{result.message}</p>
              {result.formula ? <p className="mt-2 rounded-lg bg-slate-900 px-3 py-2 font-mono text-xs text-indigo-100">{result.formula}</p> : null}
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
      <div className="rounded-2xl border border-slate-200 bg-white p-5">
        <div className="overflow-x-auto">
          <table className="min-w-full text-left text-sm">
            <thead className="text-xs uppercase text-slate-500"><tr><th className="px-2 py-2">Date</th><th className="px-2 py-2">Project</th><th className="px-2 py-2">Task</th><th className="px-2 py-2">Hours</th><th className="px-2 py-2">Billable</th><th /></tr></thead>
            <tbody>
              {entries.map((e, i) => (
                <tr key={i}>
                  <td className="px-2 py-1"><input type="date" className="rounded border border-slate-300 px-2 py-1" value={e.date} onChange={(ev) => update(i, { date: ev.target.value })} /></td>
                  <td className="px-2 py-1"><input className="rounded border border-slate-300 px-2 py-1" value={e.project} onChange={(ev) => update(i, { project: ev.target.value })} placeholder="Project" /></td>
                  <td className="px-2 py-1"><input className="rounded border border-slate-300 px-2 py-1" value={e.task} onChange={(ev) => update(i, { task: ev.target.value })} placeholder="Task" /></td>
                  <td className="px-2 py-1"><input type="number" step="0.25" className="w-20 rounded border border-slate-300 px-2 py-1" value={e.hours} onChange={(ev) => update(i, { hours: Number(ev.target.value) })} /></td>
                  <td className="px-2 py-1"><input type="checkbox" checked={e.billable} onChange={(ev) => update(i, { billable: ev.target.checked })} /></td>
                  <td className="px-2 py-1"><button className="text-xs text-red-600" onClick={() => removeRow(i)}>Remove</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <button className="mt-3 rounded-full border border-slate-300 px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={addRow}>Add row</button>
      </div>

      <div className="mt-4 grid gap-4 sm:grid-cols-3">
        <div className="rounded-2xl border border-slate-200 bg-white p-5"><p className="text-3xl font-semibold">{summary.totalHours}</p><p className="text-xs uppercase text-slate-500">Total hours</p></div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5"><p className="text-3xl font-semibold text-emerald-700">{summary.billableHours}</p><p className="text-xs uppercase text-slate-500">Billable</p></div>
        <div className="rounded-2xl border border-slate-200 bg-white p-5"><p className="text-3xl font-semibold text-slate-500">{summary.nonBillableHours}</p><p className="text-xs uppercase text-slate-500">Non-billable</p></div>
      </div>

      <button className="mt-4 rounded-full bg-indigo-500 px-5 py-2 text-xs font-semibold text-slate-950 transition hover:bg-violet-300" onClick={() => downloadMatrixExcel(timesheetToMatrix(entries), "timesheet.xlsx", "Timesheet")}>Download Excel timesheet</button>
    </Panel>
  );
}
