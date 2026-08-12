import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  analyzeWorkbook,
  buildFilteredWorkbook,
  buildRemovedSheetsLog,
  renameSheet,
  reorderSheets,
  setVisibility,
  type FilterResult,
  type SheetInfo,
  type WorkbookAnalysis,
} from "./sheet-manager";
import { downloadBlob } from "./engine";

type RenameMode = "prefix" | "suffix" | "find-replace" | "sequential";

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function StatCard({ label, value, tone }: { label: string; value: number | string; tone?: "warn" | "danger" | "default" }) {
  const toneClass = tone === "danger" ? "text-rose-600" : tone === "warn" ? "text-amber-600" : "text-slate-950";
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <p className={`text-2xl font-semibold ${toneClass}`}>{value}</p>
      <p className="mt-1 text-xs uppercase tracking-wide text-slate-500">{label}</p>
    </div>
  );
}

function VisibilityBadge({ visibility }: { visibility: SheetInfo["visibility"] }) {
  if (visibility === "visible") return <span className="rounded-full bg-emerald-100 px-2 py-0.5 text-[11px] font-medium text-emerald-700">Visible</span>;
  if (visibility === "hidden") return <span className="rounded-full bg-amber-100 px-2 py-0.5 text-[11px] font-medium text-amber-700">Hidden</span>;
  return <span className="rounded-full bg-rose-100 px-2 py-0.5 text-[11px] font-medium text-rose-700">Very Hidden</span>;
}

export default function SmartSheetManager() {
  const inputRef = useRef<HTMLInputElement>(null);
  const workbookRef = useRef<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [analysis, setAnalysis] = useState<WorkbookAnalysis | null>(null);
  const [order, setOrder] = useState<string[]>([]);
  const [keep, setKeep] = useState<Set<string>>(new Set());
  const [search, setSearch] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<FilterResult | null>(null);
  const [firstN, setFirstN] = useState(5);
  const [lastN, setLastN] = useState(5);
  const [keyword, setKeyword] = useState("");
  const [renameMode, setRenameMode] = useState<RenameMode>("prefix");
  const [renameText, setRenameText] = useState("");
  const [renameFind, setRenameFind] = useState("");
  const [renameReplace, setRenameReplace] = useState("");
  const [nonce, setNonce] = useState(0); // bumps to force table re-render after in-place workbook mutation

  const analyzedAt = useRef(0);

  async function handleFile(file: File) {
    setError("");
    setBusy(true);
    setResult(null);
    try {
      const buffer = await file.arrayBuffer();
      const start = performance.now();
      const workbook = XLSX.read(buffer, { type: "array", bookVBA: true, cellFormula: true, cellStyles: true });
      workbookRef.current = workbook;
      const info = analyzeWorkbook(workbook, buffer.byteLength);
      analyzedAt.current = performance.now() - start;
      setAnalysis(info);
      setOrder(workbook.SheetNames.slice());
      // default selection: keep everything except sheets the heuristic scores very low
      setKeep(new Set(info.sheets.filter((s) => s.keepScore >= 20).map((s) => s.name)));
      setFileName(file.name);
    } catch (err) {
      setError(err instanceof Error ? `Could not read this workbook: ${err.message}` : "Could not read this workbook.");
    } finally {
      setBusy(false);
    }
  }

  const visibleSheets = useMemo(() => {
    if (!analysis) return [];
    const q = search.trim().toLowerCase();
    return order
      .map((name) => analysis.sheets.find((s) => s.name === name))
      .filter((s): s is SheetInfo => Boolean(s))
      .filter((s) => !q || s.name.toLowerCase().includes(q));
  }, [analysis, order, search]);

  const keepCount = keep.size;
  const removeCount = (analysis?.totalSheets ?? 0) - keepCount;

  function toggleKeep(name: string) {
    setKeep((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }

  function selectAll() {
    setKeep(new Set(order));
  }
  function deselectAll() {
    setKeep(new Set());
  }
  function applyKeepFirstN() {
    setKeep(new Set(order.slice(0, Math.max(0, firstN))));
  }
  function applyKeepLastN() {
    setKeep(new Set(order.slice(Math.max(0, order.length - lastN))));
  }
  function applyKeepKeyword() {
    if (!keyword.trim()) return;
    const q = keyword.trim().toLowerCase();
    setKeep(new Set(order.filter((name) => name.toLowerCase().includes(q))));
  }
  function applyDeleteKeyword() {
    if (!keyword.trim()) return;
    const q = keyword.trim().toLowerCase();
    setKeep((prev) => new Set(Array.from(prev).filter((name) => !name.toLowerCase().includes(q))));
  }
  function applyDeleteEmpty() {
    if (!analysis) return;
    const emptyNames = new Set(analysis.sheets.filter((s) => s.isEmpty).map((s) => s.name));
    setKeep((prev) => new Set(Array.from(prev).filter((name) => !emptyNames.has(name))));
  }
  function applyDeleteHidden() {
    if (!analysis) return;
    const hiddenNames = new Set(analysis.sheets.filter((s) => s.visibility !== "visible").map((s) => s.name));
    setKeep((prev) => new Set(Array.from(prev).filter((name) => !hiddenNames.has(name))));
  }
  function applyDeleteDuplicates() {
    if (!analysis) return;
    const dupeNames = new Set(analysis.sheets.filter((s) => s.duplicateOf).map((s) => s.name));
    setKeep((prev) => new Set(Array.from(prev).filter((name) => !dupeNames.has(name))));
  }
  function applySuggested() {
    if (!analysis) return;
    setKeep(new Set(analysis.sheets.filter((s) => s.keepScore >= 60).map((s) => s.name)));
  }

  function applySort(mode: "az" | "za") {
    setOrder((prev) => {
      const next = prev.slice().sort((a, b) => (mode === "az" ? a.localeCompare(b) : b.localeCompare(a)));
      return next;
    });
  }
  function moveSheet(name: string, direction: -1 | 1) {
    setOrder((prev) => {
      const idx = prev.indexOf(name);
      const target = idx + direction;
      if (idx === -1 || target < 0 || target >= prev.length) return prev;
      const next = prev.slice();
      [next[idx], next[target]] = [next[target], next[idx]];
      return next;
    });
  }

  function applyRename() {
    const wb = workbookRef.current;
    if (!wb || !analysis) return;
    const targets = Array.from(keep);
    let sequence = 1;
    targets.forEach((name) => {
      let next = name;
      if (renameMode === "prefix" && renameText) next = `${renameText}${name}`;
      if (renameMode === "suffix" && renameText) next = `${name}${renameText}`;
      if (renameMode === "find-replace" && renameFind) next = name.split(renameFind).join(renameReplace);
      if (renameMode === "sequential" && renameText) next = `${renameText} ${sequence++}`;
      if (next !== name) {
        renameSheet(wb, name, next);
        setOrder((prevOrder) => prevOrder.map((n) => (n === name ? next : n)));
        setKeep((prevKeep) => {
          if (!prevKeep.has(name)) return prevKeep;
          const nextKeep = new Set(prevKeep);
          nextKeep.delete(name);
          nextKeep.add(next);
          return nextKeep;
        });
      }
    });
    const refreshed = analyzeWorkbook(wb, analysis.originalSizeBytes);
    setAnalysis(refreshed);
    setNonce((n) => n + 1);
  }

  function applyVisibility(visibility: SheetInfo["visibility"]) {
    const wb = workbookRef.current;
    if (!wb || !analysis) return;
    keep.forEach((name) => setVisibility(wb, name, visibility));
    const refreshed = analyzeWorkbook(wb, analysis.originalSizeBytes);
    setAnalysis(refreshed);
    setNonce((n) => n + 1);
  }

  function buildWorkbook() {
    const wb = workbookRef.current;
    if (!wb || !analysis) return;
    setError("");
    if (keep.size === 0) {
      setError("Select at least one sheet to keep — a workbook cannot have zero worksheets.");
      return;
    }
    try {
      reorderSheets(wb, order);
      const keepOrdered = order.filter((name) => keep.has(name));
      const outcome = buildFilteredWorkbook(wb, keepOrdered);
      setResult(outcome);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not build the cleaned workbook.");
    }
  }

  function downloadCleaned() {
    if (!result) return;
    const base = fileName.replace(/\.(xlsx|xlsm|xls)$/i, "") || "workbook";
    const ext = result.usedXlsm ? "xlsm" : "xlsx";
    downloadBlob(result.bytes, `${base}-cleaned.${ext}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  function downloadLog() {
    if (!result || !analysis) return;
    const log = buildRemovedSheetsLog(analysis, result.removedSheets);
    downloadBlob(log, "smart-sheet-manager-removed-sheets.csv", "text/csv");
  }

  function reset() {
    workbookRef.current = null;
    setAnalysis(null);
    setOrder([]);
    setKeep(new Set());
    setResult(null);
    setFileName("");
    setError("");
  }

  return (
    <section key={nonce}>
      <h1 className="text-3xl font-semibold tracking-[-0.03em]">Smart Sheet Manager</h1>
      <p className="mt-2 max-w-3xl text-slate-600">
        Upload an Excel workbook with any number of worksheets. Smart Sheet Manager analyses every sheet — visibility,
        formulas, cross-sheet references, empty sheets, and duplicates — then lets you keep exactly the sheets you need
        and rebuild a clean workbook in seconds.
      </p>

      <div className="mt-4 rounded-xl border border-teal-200 bg-teal-50 p-4 text-sm text-teal-900">
        <p className="font-semibold">What is genuinely preserved when the workbook is rebuilt</p>
        <p className="mt-1 text-teal-800">
          Cell values, formulas, number formats, merged cells, column widths, and defined names are preserved. If your
          file already contains a VBA project (.xlsm), it is carried through unchanged. <strong>Charts and PivotTables
          are not preserved on rebuild</strong> — this is a limitation of the underlying open-source engine, not
          something Smart Sheet Manager silently drops without telling you.
        </p>
      </div>

      {error ? <p className="mt-4 rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">{error}</p> : null}

      {!analysis ? (
        <div
          className="mt-6 rounded-2xl border border-dashed border-slate-300 bg-white p-10 text-center transition hover:border-teal-500"
          onDragOver={(event) => event.preventDefault()}
          onDrop={(event) => {
            event.preventDefault();
            const file = event.dataTransfer.files?.[0];
            if (file) handleFile(file);
          }}
        >
          <input
            ref={inputRef}
            type="file"
            className="sr-only"
            accept=".xlsx,.xlsm,.xls"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) handleFile(file);
              if (inputRef.current) inputRef.current.value = "";
            }}
          />
          <p className="text-base font-semibold text-slate-900">{busy ? "Analysing workbook..." : "Drop an Excel workbook, or choose a file"}</p>
          <p className="mt-2 text-sm text-slate-500">Supports .xlsx and .xlsm, from 1 to thousands of worksheets.</p>
          <button className="mt-4 rounded-full bg-slate-950 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-teal-700" onClick={() => inputRef.current?.click()} disabled={busy}>
            Choose file
          </button>
        </div>
      ) : (
        <div className="mt-6 space-y-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <p className="text-sm text-slate-600">
              <span className="font-semibold text-slate-900">{fileName}</span> — analysed in {analyzedAt.current.toFixed(0)}ms
            </p>
            <button className="rounded-full border border-slate-300 px-4 py-1.5 text-xs font-semibold text-slate-600 transition hover:border-teal-500 hover:text-teal-700" onClick={reset}>
              Upload a different file
            </button>
          </div>

          <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
            <StatCard label="Total sheets" value={analysis.totalSheets} />
            <StatCard label="Visible" value={analysis.visibleCount} />
            <StatCard label="Hidden" value={analysis.hiddenCount} tone="warn" />
            <StatCard label="Very hidden" value={analysis.veryHiddenCount} tone="warn" />
            <StatCard label="Empty" value={analysis.emptyCount} tone="danger" />
            <StatCard label="Duplicate groups" value={analysis.duplicateGroups.length} tone="danger" />
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <p className="text-sm font-semibold text-slate-900">Bulk actions</p>
            <div className="mt-3 flex flex-wrap items-center gap-2 text-xs">
              <button className="rounded-full bg-slate-100 px-3 py-1.5 font-semibold text-slate-700 hover:bg-slate-200" onClick={selectAll}>Select all</button>
              <button className="rounded-full bg-slate-100 px-3 py-1.5 font-semibold text-slate-700 hover:bg-slate-200" onClick={deselectAll}>Deselect all</button>
              <button className="rounded-full bg-teal-100 px-3 py-1.5 font-semibold text-teal-700 hover:bg-teal-200" onClick={applySuggested}>Apply suggested keep list</button>
              <button className="rounded-full bg-rose-100 px-3 py-1.5 font-semibold text-rose-700 hover:bg-rose-200" onClick={applyDeleteEmpty}>Delete empty sheets</button>
              <button className="rounded-full bg-rose-100 px-3 py-1.5 font-semibold text-rose-700 hover:bg-rose-200" onClick={applyDeleteHidden}>Delete hidden sheets</button>
              <button className="rounded-full bg-rose-100 px-3 py-1.5 font-semibold text-rose-700 hover:bg-rose-200" onClick={applyDeleteDuplicates}>Delete duplicate sheets</button>
            </div>
            <div className="mt-4 flex flex-wrap items-end gap-3 text-xs">
              <label className="flex items-center gap-2">Keep first
                <input type="number" min={0} className="w-16 rounded border border-slate-300 px-2 py-1" value={firstN} onChange={(e) => setFirstN(Number(e.target.value))} />
                <button className="rounded-full bg-slate-950 px-3 py-1.5 font-semibold text-white" onClick={applyKeepFirstN}>Apply</button>
              </label>
              <label className="flex items-center gap-2">Keep last
                <input type="number" min={0} className="w-16 rounded border border-slate-300 px-2 py-1" value={lastN} onChange={(e) => setLastN(Number(e.target.value))} />
                <button className="rounded-full bg-slate-950 px-3 py-1.5 font-semibold text-white" onClick={applyKeepLastN}>Apply</button>
              </label>
              <label className="flex items-center gap-2">Keyword
                <input type="text" placeholder="e.g. Summary" className="w-40 rounded border border-slate-300 px-2 py-1" value={keyword} onChange={(e) => setKeyword(e.target.value)} />
                <button className="rounded-full bg-slate-950 px-3 py-1.5 font-semibold text-white" onClick={applyKeepKeyword}>Keep matching</button>
                <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={applyDeleteKeyword}>Delete matching</button>
              </label>
            </div>
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <p className="text-sm font-semibold text-slate-900">Rename &amp; sort</p>
            <div className="mt-3 flex flex-wrap items-end gap-3 text-xs">
              <select className="rounded border border-slate-300 px-2 py-1.5" value={renameMode} onChange={(e) => setRenameMode(e.target.value as RenameMode)}>
                <option value="prefix">Add prefix</option>
                <option value="suffix">Add suffix</option>
                <option value="find-replace">Find &amp; replace</option>
                <option value="sequential">Sequential naming</option>
              </select>
              {renameMode === "find-replace" ? (
                <>
                  <input placeholder="Find" className="w-32 rounded border border-slate-300 px-2 py-1" value={renameFind} onChange={(e) => setRenameFind(e.target.value)} />
                  <input placeholder="Replace with" className="w-32 rounded border border-slate-300 px-2 py-1" value={renameReplace} onChange={(e) => setRenameReplace(e.target.value)} />
                </>
              ) : (
                <input placeholder={renameMode === "sequential" ? "Base name, e.g. Month" : "Text"} className="w-40 rounded border border-slate-300 px-2 py-1" value={renameText} onChange={(e) => setRenameText(e.target.value)} />
              )}
              <button className="rounded-full bg-slate-950 px-3 py-1.5 font-semibold text-white" onClick={applyRename}>Apply to selected ({keepCount})</button>
              <span className="mx-2 h-4 w-px bg-slate-200" />
              <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={() => applySort("az")}>Sort A-Z</button>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={() => applySort("za")}>Sort Z-A</button>
              <span className="mx-2 h-4 w-px bg-slate-200" />
              <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={() => applyVisibility("hidden")}>Hide selected</button>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={() => applyVisibility("visible")}>Unhide selected</button>
              <button className="rounded-full border border-slate-300 px-3 py-1.5 font-semibold text-slate-700" onClick={() => applyVisibility("veryHidden")}>Very-hide selected</button>
            </div>
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 p-4">
              <input
                type="text"
                placeholder="Search sheet names..."
                className="w-full max-w-xs rounded-lg border border-slate-300 px-3 py-2 text-sm"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <p className="text-sm font-semibold text-slate-700">
                <span className="text-emerald-700">{keepCount} sheet{keepCount === 1 ? "" : "s"} will remain</span>
                {" · "}
                <span className="text-rose-700">{removeCount} sheet{removeCount === 1 ? "" : "s"} will be removed</span>
              </p>
            </div>
            <div className="max-h-[28rem] overflow-y-auto">
              <table className="min-w-full text-left text-sm">
                <thead className="sticky top-0 bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
                  <tr>
                    <th className="w-10 py-2 pl-4"></th>
                    <th className="py-2 pr-4">Sheet name</th>
                    <th className="py-2 pr-4">Status</th>
                    <th className="py-2 pr-4">Size</th>
                    <th className="py-2 pr-4">Formulas</th>
                    <th className="py-2 pr-4">Suggested keep score</th>
                    <th className="py-2 pr-4">Reorder</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {visibleSheets.map((sheet) => (
                    <tr key={sheet.name} className={keep.has(sheet.name) ? "" : "bg-slate-50/60 opacity-60"}>
                      <td className="py-2 pl-4">
                        <input type="checkbox" checked={keep.has(sheet.name)} onChange={() => toggleKeep(sheet.name)} />
                      </td>
                      <td className="py-2 pr-4 font-medium text-slate-900">
                        {sheet.name}
                        {sheet.duplicateOf ? <span className="ml-2 rounded bg-rose-100 px-1.5 py-0.5 text-[10px] font-semibold text-rose-700">Duplicate of {sheet.duplicateOf}</span> : null}
                        {sheet.isEmpty ? <span className="ml-2 rounded bg-slate-200 px-1.5 py-0.5 text-[10px] font-semibold text-slate-600">Empty</span> : null}
                      </td>
                      <td className="py-2 pr-4"><VisibilityBadge visibility={sheet.visibility} /></td>
                      <td className="py-2 pr-4 text-slate-500">{sheet.rowCount}×{sheet.colCount}</td>
                      <td className="py-2 pr-4 text-slate-500">{sheet.hasFormulas ? "Yes" : "—"}{sheet.referencedByFormulaCount > 0 ? ` · referenced ×${sheet.referencedByFormulaCount}` : ""}</td>
                      <td className="py-2 pr-4">
                        <div className="flex items-center gap-2" title={sheet.keepReasons.join("; ") || "No strong signals either way"}>
                          <div className="h-1.5 w-16 rounded-full bg-slate-200">
                            <div className={`h-1.5 rounded-full ${sheet.keepScore >= 60 ? "bg-emerald-500" : sheet.keepScore >= 30 ? "bg-amber-500" : "bg-rose-500"}`} style={{ width: `${sheet.keepScore}%` }} />
                          </div>
                          <span className="text-xs text-slate-500">{sheet.keepScore}</span>
                        </div>
                      </td>
                      <td className="py-2 pr-4">
                        <button className="rounded border border-slate-300 px-1.5 text-xs text-slate-500 hover:border-teal-500" onClick={() => moveSheet(sheet.name, -1)} aria-label={`Move ${sheet.name} up`}>↑</button>
                        <button className="ml-1 rounded border border-slate-300 px-1.5 text-xs text-slate-500 hover:border-teal-500" onClick={() => moveSheet(sheet.name, 1)} aria-label={`Move ${sheet.name} down`}>↓</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-5">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <p className="text-sm font-semibold text-slate-900">Build the cleaned workbook</p>
                <p className="mt-1 text-xs text-slate-500">
                  {keepCount === 0
                    ? "Select at least one sheet to keep before building."
                    : `${keepCount} sheet${keepCount === 1 ? "" : "s"} will be kept, ${removeCount} will be removed.`}
                </p>
              </div>
              <button
                className="rounded-full bg-teal-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-teal-700 disabled:cursor-not-allowed disabled:bg-slate-300"
                onClick={buildWorkbook}
                disabled={keepCount === 0}
              >
                Build cleaned workbook
              </button>
            </div>

            {result ? (
              <div className="mt-5 rounded-xl border border-emerald-200 bg-emerald-50 p-4">
                <p className="text-sm font-semibold text-emerald-900">Workbook rebuilt successfully</p>
                <div className="mt-2 grid grid-cols-2 gap-3 text-sm text-emerald-800 sm:grid-cols-4">
                  <div><p className="text-xs uppercase text-emerald-600">Original size</p><p className="font-semibold">{formatBytes(analysis.originalSizeBytes)}</p></div>
                  <div><p className="text-xs uppercase text-emerald-600">New size</p><p className="font-semibold">{formatBytes(result.newSizeBytes)}</p></div>
                  <div><p className="text-xs uppercase text-emerald-600">Size reduced</p><p className="font-semibold">{Math.max(0, Math.round((1 - result.newSizeBytes / Math.max(1, analysis.originalSizeBytes)) * 100))}%</p></div>
                  <div><p className="text-xs uppercase text-emerald-600">Processing time</p><p className="font-semibold">{result.elapsedMs.toFixed(1)}ms</p></div>
                </div>
                <div className="mt-4 flex flex-wrap gap-2">
                  <button className="rounded-full bg-emerald-700 px-4 py-2 text-xs font-semibold text-white hover:bg-emerald-800" onClick={downloadCleaned}>Download cleaned workbook</button>
                  <button className="rounded-full border border-emerald-300 px-4 py-2 text-xs font-semibold text-emerald-800 hover:border-emerald-500" onClick={downloadLog}>Download removed-sheets log</button>
                </div>
              </div>
            ) : null}
          </div>
        </div>
      )}
    </section>
  );
}
