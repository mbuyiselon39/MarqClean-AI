import { useMemo, useRef, useState } from "react";
import {
  analyzeSheet,
  autoDetect,
  buildPreview,
  formatAmount,
  generateAndValidateWorkbook,
  generateTestWorkbook,
  readWorkbook,
  validateProcessing,
  type ColumnMap,
  type PreviewResult,
  type ProcessOptions,
  type ProcessResult,
  type RawSheet,
  type WorkbookData,
  type WorkbookValidation,
} from "./client-funds";
import { GlobalHeader } from "./WorkspaceShell";

export default function ClientFunds({ onExit }: { onExit: () => void }) {
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [startRow, setStartRow] = useState(101);
  const [columns, setColumns] = useState<ColumnMap>({ dateCol: 0, clientCol: 1, srnCol: 2, amountCol: 3 });
  const [ignoreSubtotals, setIgnoreSubtotals] = useState(true);
  const [addSubtotals, setAddSubtotals] = useState(true);
  const [addBlankRow, setAddBlankRow] = useState(true);
  const [createFormulaGuide, setCreateFormulaGuide] = useState(true);
  const [createErrorReport, setCreateErrorReport] = useState(true);

  const [preview, setPreview] = useState<PreviewResult | null>(null);
  const [processed, setProcessed] = useState<ProcessResult | null>(null);
  const [downloadInfo, setDownloadInfo] = useState<{ bytes: Uint8Array; fileName: string } | null>(null);
  const [validation, setValidation] = useState<WorkbookValidation | null>(null);
  const [showErrors, setShowErrors] = useState(false);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [testResult, setTestResult] = useState<{ workbook: { bytes: Uint8Array; fileName: string }; validation: WorkbookValidation } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const activeSheet: RawSheet | null = workbook && sheetName ? workbook.sheets[sheetName] ?? null : null;

  const options: ProcessOptions = useMemo(
    () => ({ sheetName, headerRow, startRow, columns, ignoreSubtotals, addSubtotals, addBlankRow, createFormulaGuide, createErrorReport }),
    [sheetName, headerRow, startRow, columns, ignoreSubtotals, addSubtotals, addBlankRow, createFormulaGuide, createErrorReport]
  );

  async function handleFile(file: File) {
    setError("");
    setBusy(true);
    setPreview(null);
    setProcessed(null);
    setDownloadInfo(null);
    try {
      const ext = file.name.split(".").pop()?.toLowerCase();
      if (ext !== "xlsx" && ext !== "xls") throw new Error("Please upload an .xlsx or .xls workbook.");
      const wb = await readWorkbook(file);
      if (!wb.sheetNames.length) throw new Error("No worksheets found in the workbook.");
      const first = wb.sheetNames[0];
      const detected = autoDetect(wb.sheets[first]);
      setWorkbook(wb);
      setSheetName(first);
      setHeaderRow(detected.headerRow);
      setStartRow(detected.startRow);
      setColumns(detected.columns);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "The file could not be read.");
      setWorkbook(null);
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function reautoDetect() {
    if (!workbook || !sheetName) return;
    const detected = autoDetect(workbook.sheets[sheetName]);
    setHeaderRow(detected.headerRow);
    setStartRow(detected.startRow);
    setColumns(detected.columns);
  }

  function runPreview() {
    if (!activeSheet) return;
    setError("");
    setPreview(buildPreview(activeSheet, options));
  }

  async function execute() {
    if (!activeSheet || !workbook) return;
    setError("");
    setBusy(true);
    setDownloadInfo(null);
    setValidation(null);
    try {
      const result = analyzeSheet(activeSheet, options);
      if (!result.details.length) throw new Error("No valid detail rows were detected. Adjust the start row or column mapping and preview again.");
      const issues = validateProcessing(result, options);
      if (issues.length) throw new Error(`Validation failed: ${issues.join(" ")}`);

      // Generate with SheetJS and re-read the produced binary before allowing download
      const { workbook: output, validation: check } = await generateAndValidateWorkbook(workbook, result, options);
      setProcessed(result);
      setValidation(check);
      if (check.ok) {
        setDownloadInfo(output);
      } else {
        setDownloadInfo(null);
        setError("Workbook validation failed. The generated Excel file is invalid.");
      }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Processing failed.");
    } finally {
      setBusy(false);
    }
  }

  function downloadBytes(bytes: Uint8Array, fileName: string) {
    const blob = new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 250);
  }

  function download() {
    if (!downloadInfo || !validation?.ok) return;
    downloadBytes(downloadInfo.bytes, downloadInfo.fileName);
  }

  async function runTestWorkbook() {
    setError("");
    setBusy(true);
    try {
      const result = await generateTestWorkbook();
      setTestResult(result);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Test workbook failed.");
    } finally {
      setBusy(false);
    }
  }

  function reset() {
    setWorkbook(null); setSheetName(""); setPreview(null); setProcessed(null); setDownloadInfo(null); setValidation(null); setError(""); setShowErrors(false);
  }

  const colOptions = activeSheet ? Array.from({ length: activeSheet.maxCols }, (_, i) => i) : [];
  const colLabel = (i: number) => String.fromCharCode(65 + i);

  return (
    <div className="ws-root ws-scope relative min-h-screen">
      <GlobalHeader />
      <div className="ws-grid pointer-events-none absolute inset-x-0 top-0 h-[38rem]" aria-hidden="true" />

      <main id="main-content" className="relative mx-auto max-w-6xl px-5 pb-20 pt-24 lg:px-8">
        <div className="ws-surface mb-6 flex flex-wrap items-start justify-between gap-4 rounded-2xl p-6">
          <div className="flex items-start gap-4">
            <span className="flex h-12 w-12 items-center justify-center rounded-xl bg-gradient-to-br from-[#0D9488]/20 to-[#14B8A6]/20 text-2xl ring-1 ring-inset ring-white/10">📊</span>
            <div>
              <h1 className="mc-display text-2xl font-bold tracking-tight text-white sm:text-3xl">Excel Automation</h1>
              <p className="mt-1 max-w-2xl text-sm text-[#b0bacb]">Automate Excel sorting, subtotals, formulas and workbook generation. Includes the Sort &amp; Subtotal Engine, number formatting and error reporting.</p>
              <div className="mt-2 flex flex-wrap gap-2">
                <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-semibold uppercase">XLSX</span>
                <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-semibold uppercase">XLS</span>
                <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-semibold text-[#5EEAD4]">Engine: MarqClean AI Structure Engine</span>
              </div>
            </div>
          </div>
          <button className="ws-btn-secondary rounded-full px-4 py-2 text-sm font-medium" onClick={onExit}>Back to MarqClean AI</button>
        </div>
        <p className="mt-3 max-w-3xl text-slate-600">
          Upload an Excel workbook, preview the detected data, then click Execute Sorting. The tool groups related
          records by the root SRN or reference before the first slash, subtotals each group, inserts a blank spacer row,
          and produces a new workbook with a Sorted_Output sheet, a Formula_Guide, and an Error_Report when needed.
          Amounts are stored as real numbers and formatted as #,##0.00, and European-style values such as 182,06 or
          1.000,00 are converted automatically. The original worksheet is preserved. Everything runs in your browser.
        </p>

        {error ? <p className="mt-5 rounded-2xl bg-red-50 p-4 text-sm font-medium text-red-700" role="alert">{error}</p> : null}

        {/* Step 1: Upload */}
        <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
          <h2 className="text-lg font-semibold">1. Upload Excel File</h2>
          <div
            className="mt-3 rounded-xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center transition hover:border-teal-500"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) void handleFile(f); }}
          >
            <input ref={inputRef} className="sr-only" type="file" accept=".xlsx,.xls" onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleFile(f); }} />
            {workbook ? (
              <p className="text-sm text-slate-600">{workbook.fileName} loaded</p>
            ) : (
              <div className="text-sm text-slate-600">
                <p className="font-medium text-white">Drop Excel, CSV, PDF or Word files</p>
                <div className="mt-2 flex flex-wrap justify-center gap-2 text-xs">
                  <span className="ws-badge rounded-md px-2 py-0.5 font-semibold uppercase">XLSX</span>
                  <span className="ws-badge rounded-md px-2 py-0.5 font-semibold uppercase">XLS</span>
                </div>
                <p className="mt-2 text-xs text-[#8a94a8]">Maximum file size: 2GB &middot; Browser-based secure processing</p>
              </div>
            )}
            <div className="mt-4 flex flex-wrap justify-center gap-2">
              <button className="rounded-full bg-slate-950 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-teal-700" onClick={() => inputRef.current?.click()} disabled={busy}>
                {busy ? "Reading..." : workbook ? "Replace file" : "Upload Excel File"}
              </button>
              {workbook ? <button className="rounded-full border border-slate-300 px-5 py-2.5 text-sm font-semibold text-slate-700 transition hover:border-slate-950" onClick={reset}>Reset Upload</button> : null}
            </div>
          </div>
        </section>

        {/* Step 2: Configure + Preview */}
        {workbook && activeSheet ? (
          <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <h2 className="text-lg font-semibold">2. Configure &amp; Preview Data</h2>
              <button className="rounded-full border border-slate-300 px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-slate-950" onClick={reautoDetect}>Auto detect</button>
            </div>

            <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Worksheet
                <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={sheetName} onChange={(e) => { setSheetName(e.target.value); const d = autoDetect(workbook.sheets[e.target.value]); setHeaderRow(d.headerRow); setStartRow(d.startRow); setColumns(d.columns); }}>
                  {workbook.sheetNames.map((n) => <option key={n} value={n}>{n}</option>)}
                </select>
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Header row
                <input type="number" min={1} className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={headerRow} onChange={(e) => setHeaderRow(Math.max(1, Number(e.target.value)))} />
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Data start row
                <input type="number" min={1} className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={startRow} onChange={(e) => setStartRow(Math.max(1, Number(e.target.value)))} />
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Date / Note column
                <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={columns.dateCol} onChange={(e) => setColumns({ ...columns, dateCol: Number(e.target.value) })}>{colOptions.map((i) => <option key={i} value={i}>{colLabel(i)}</option>)}</select>
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Client Name / Reference column
                <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={columns.clientCol} onChange={(e) => setColumns({ ...columns, clientCol: Number(e.target.value) })}>{colOptions.map((i) => <option key={i} value={i}>{colLabel(i)}</option>)}</select>
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">SRN / Account Reference column
                <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={columns.srnCol} onChange={(e) => setColumns({ ...columns, srnCol: Number(e.target.value) })}>{colOptions.map((i) => <option key={i} value={i}>{colLabel(i)}</option>)}</select>
              </label>
              <label className="text-xs font-semibold uppercase tracking-wide text-slate-500">Amount column
                <select className="mt-1 block w-full rounded border border-slate-300 px-2 py-1.5 text-sm font-normal" value={columns.amountCol} onChange={(e) => setColumns({ ...columns, amountCol: Number(e.target.value) })}>{colOptions.map((i) => <option key={i} value={i}>{colLabel(i)}</option>)}</select>
              </label>
            </div>

            <div className="mt-4 flex flex-wrap gap-4 text-sm text-slate-700">
              <label className="flex items-center gap-2"><input type="checkbox" checked={ignoreSubtotals} onChange={(e) => setIgnoreSubtotals(e.target.checked)} />Ignore existing subtotal rows</label>
              <label className="flex items-center gap-2"><input type="checkbox" checked={addSubtotals} onChange={(e) => setAddSubtotals(e.target.checked)} />Add subtotal rows</label>
              <label className="flex items-center gap-2"><input type="checkbox" checked={addBlankRow} onChange={(e) => setAddBlankRow(e.target.checked)} />Add blank row between groups</label>
              <label className="flex items-center gap-2"><input type="checkbox" checked={createFormulaGuide} onChange={(e) => setCreateFormulaGuide(e.target.checked)} />Create Formula_Guide sheet</label>
              <label className="flex items-center gap-2"><input type="checkbox" checked={createErrorReport} onChange={(e) => setCreateErrorReport(e.target.checked)} />Create Error_Report sheet</label>
            </div>

            <button className="mt-5 rounded-full bg-slate-950 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-teal-700" onClick={runPreview}>Preview Data</button>

            {preview ? (
              <div className="mt-5">
                <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                  <Stat label="Valid detail rows" value={preview.validCount} />
                  <Stat label="Blank rows ignored" value={preview.blankCount} />
                  <Stat label="Existing subtotal rows" value={preview.subtotalCount} />
                  <Stat label="Invalid amount rows" value={preview.invalidCount} />
                </div>
                <p className="mt-3 text-xs text-slate-500">Worksheet: {preview.sheetName} | Header row: {preview.headerRow} | Start row: {preview.startRow} | Detected columns: {preview.headers.join(", ")}</p>
                <div className="mt-3 overflow-auto rounded-xl border border-slate-200" style={{ maxHeight: "22rem" }}>
                  <table className="min-w-full text-left text-sm">
                    <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500"><tr><th className="px-3 py-2">Date / Note</th><th className="px-3 py-2">Client Name</th><th className="px-3 py-2">SRN</th><th className="px-3 py-2">Amount</th><th className="px-3 py-2">Grouping key</th></tr></thead>
                    <tbody className="divide-y divide-slate-100">
                      {preview.sampleRows.map((row, i) => (
                        <tr key={i}><td className="px-3 py-1.5 text-slate-700">{row.date || "-"}</td><td className="px-3 py-1.5 text-slate-700">{row.client || "-"}</td><td className="px-3 py-1.5 text-slate-700">{row.srn || "-"}</td><td className="px-3 py-1.5 text-slate-700">{row.amount}</td><td className="px-3 py-1.5 font-medium text-teal-700">{row.key}</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : null}
          </section>
        ) : null}

        {/* Diagnostics: Test Workbook */}
        <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
          <h2 className="text-lg font-semibold">Diagnostics</h2>
          <p className="mt-1 text-sm text-slate-600">Generate a minimal test workbook to confirm the Excel export mechanism produces valid files that open in Microsoft Excel.</p>
          <button className="mt-3 rounded-full border border-slate-300 px-5 py-2.5 text-sm font-semibold text-slate-700 transition hover:border-slate-950" onClick={runTestWorkbook} disabled={busy}>Generate Test Workbook</button>
          {testResult ? (
            <div className={`mt-3 rounded-xl p-3 text-sm ${testResult.validation.ok ? "bg-emerald-50 text-emerald-800" : "bg-red-50 text-red-700"}`}>
              <p className="font-medium">{testResult.validation.ok ? "Test workbook is valid and reopened successfully." : "Test workbook failed validation."}</p>
              <p className="mt-1 text-xs">Size: {testResult.validation.size} bytes | Sheets: {testResult.validation.sheetSummaries.map((s) => s.name).join(", ") || "none"}</p>
              {testResult.validation.issues.length ? <p className="mt-1 text-xs">{testResult.validation.issues.join(" ")}</p> : null}
              {testResult.validation.ok ? <button className="mt-2 rounded-full bg-slate-950 px-4 py-1.5 text-xs font-semibold text-white transition hover:bg-teal-700" onClick={() => downloadBytes(testResult.workbook.bytes, testResult.workbook.fileName)}>Download Test Workbook</button> : null}
            </div>
          ) : null}
        </section>

        {/* Step 3: Execute */}
        {preview ? (
          <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
            <h2 className="text-lg font-semibold">3. Execute Sorting</h2>
            <p className="mt-1 text-sm text-slate-600">Runs grouping, sorting, subtotaling, workbook generation, and client-side validation. The download only unlocks after the generated file is re-opened successfully.</p>
            <button className="mt-4 rounded-full bg-teal-500 px-6 py-3 text-sm font-semibold text-slate-950 transition hover:bg-teal-300" onClick={execute} disabled={busy}>
              {busy ? "Processing..." : "Execute Sorting"}
            </button>
          </section>
        ) : null}

        {/* Validation failure panel */}
        {validation && !validation.ok ? (
          <section className="mt-6 rounded-2xl border border-red-200 bg-red-50 p-5" role="alert">
            <p className="text-lg font-semibold text-red-700">Workbook validation failed. The generated Excel file is invalid.</p>
            <ul className="mt-3 list-disc space-y-1 pl-5 text-sm text-red-700">
              {validation.issues.map((iss, i) => <li key={i}>{iss}</li>)}
            </ul>
            <p className="mt-3 text-xs text-red-600">Sheets attempted: {validation.sheetSummaries.map((s) => `${s.name} (${s.rows} rows)`).join(", ") || "none"} | Binary size: {validation.size} bytes. Download is blocked until validation passes.</p>
          </section>
        ) : null}

        {/* Step 4: Results */}
        {processed && downloadInfo && validation?.ok ? (
          <section className="mt-6 rounded-2xl border border-slate-200 bg-white p-5">
            <h2 className="text-lg font-semibold">4. Download Sorted Workbook</h2>
            <div className="mt-3 grid gap-3 sm:grid-cols-3">
              <Stat label="Groups created" value={processed.groups.length} />
              <Stat label="Detail rows sorted" value={processed.details.length} />
              <Stat label="Rows flagged" value={processed.errors.length} />
            </div>
            <div className="mt-4 flex flex-wrap gap-2">
              <button className="rounded-full bg-slate-950 px-6 py-3 text-sm font-semibold text-white transition hover:bg-teal-700" onClick={download}>Download Sorted Workbook</button>
              {processed.errors.length ? <button className="rounded-full border border-slate-300 px-5 py-3 text-sm font-semibold text-slate-700 transition hover:border-slate-950" onClick={() => setShowErrors((s) => !s)}>{showErrors ? "Hide" : "View"} Error Report ({processed.errors.length})</button> : null}
            </div>
            <p className="mt-2 text-xs text-emerald-700">Validated: the generated .xlsx was re-opened successfully ({validation.size} bytes, sheets: {validation.sheetSummaries.map((s) => s.name).join(", ")}). Safe to download.</p>
            <p className="mt-1 text-xs text-slate-500">Output file: {downloadInfo.fileName}. Contains the original worksheet, Sorted_Output{createFormulaGuide ? ", Formula_Guide" : ""}{processed.errors.length && createErrorReport ? ", Error_Report" : ""}.</p>

            {showErrors && processed.errors.length ? (
              <div className="mt-4 overflow-auto rounded-xl border border-amber-200" style={{ maxHeight: "22rem" }}>
                <table className="min-w-full text-left text-sm">
                  <thead className="bg-amber-50 text-xs uppercase tracking-wide text-amber-700"><tr><th className="px-3 py-2">Row</th><th className="px-3 py-2">A</th><th className="px-3 py-2">B</th><th className="px-3 py-2">C</th><th className="px-3 py-2">D</th><th className="px-3 py-2">Reason</th></tr></thead>
                  <tbody className="divide-y divide-amber-100">
                    {processed.errors.map((e, i) => (
                      <tr key={i}><td className="px-3 py-1.5">{e.sourceRow}</td><td className="px-3 py-1.5">{e.a}</td><td className="px-3 py-1.5">{e.b}</td><td className="px-3 py-1.5">{e.c}</td><td className="px-3 py-1.5">{e.d}</td><td className="px-3 py-1.5 text-amber-700">{e.reason}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : null}

            {/* Preview of grouped output */}
            <div className="mt-5">
              <h3 className="text-sm font-semibold text-slate-900">Sorted output preview</h3>
              <div className="mt-2 overflow-auto rounded-xl border border-slate-200" style={{ maxHeight: "26rem" }}>
                <table className="min-w-full text-left text-sm">
                  <thead className="bg-[#1a1b3a] text-xs uppercase tracking-wide text-white"><tr><th className="px-3 py-2">Date / Note</th><th className="px-3 py-2">Client Name</th><th className="px-3 py-2">SRN</th><th className="px-3 py-2 text-right">Amount</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">
                    {processed.groups.slice(0, 30).flatMap((g, gi) => [
                      ...g.rows.map((row, ri) => (
                        <tr key={`${gi}-${ri}`}><td className="px-3 py-1.5 text-slate-700">{row.date || ""}</td><td className="px-3 py-1.5 text-slate-700">{row.client}</td><td className="px-3 py-1.5 text-slate-700">{row.srn}</td><td className="px-3 py-1.5 text-right text-slate-700">{formatAmount(row.amount)}</td></tr>
                      )),
                      <tr key={`${gi}-sub`} className="bg-slate-50"><td className="px-3 py-1.5" /><td className="px-3 py-1.5" /><td className="px-3 py-1.5" /><td className="px-3 py-1.5 text-right font-bold text-slate-900">{formatAmount(g.subtotal)}</td></tr>,
                      <tr key={`${gi}-blank`}><td className="px-3 py-1.5" colSpan={4}>&nbsp;</td></tr>,
                    ])}
                  </tbody>
                </table>
              </div>
              {processed.groups.length > 30 ? <p className="mt-2 text-xs text-slate-500">Showing first 30 groups. The downloaded workbook contains all {processed.groups.length} groups.</p> : null}
            </div>
          </section>
        ) : null}
      </main>
    </div>
  );
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <p className="text-2xl font-semibold text-slate-950">{value}</p>
      <p className="mt-1 text-xs uppercase tracking-wide text-slate-500">{label}</p>
    </div>
  );
}
