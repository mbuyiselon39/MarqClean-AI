import { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

export type SmartRecommendation = {
  id: "cleaner" | "reconciliation" | "converter" | "formulas" | "bank" | "data-engine";
  label: string;
  reason: string;
};

export type SmartPipelineAction = {
  id: "trim" | "group-reference" | "subtotals" | "export-xlsx";
  label: string;
};

type Diagnostics = {
  rows: number;
  columns: number;
  corruptedRows: number;
  nonStandardDates: number;
  headers: string[];
  recommendations: SmartRecommendation[];
};

const ACTIONS: SmartPipelineAction[] = [
  { id: "trim", label: "Trim whitespaces" },
  { id: "group-reference", label: "Group rows by reference" },
  { id: "subtotals", label: "Inject SUM subtotals" },
  { id: "export-xlsx", label: "Export as Clean XLSX" },
];

const normalise = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, "");

function detectDate(value: string) {
  const v = value.trim();
  if (!v) return false;
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(v)) return true;
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(v)) return true;
  return false;
}

function analyseMatrix(matrix: unknown[][]): Diagnostics {
  const rows = matrix.filter((row) => row.some((v) => String(v ?? "").trim() !== ""));
  const headers = (rows[0] ?? []).map((v) => String(v ?? "").trim());
  const expected = headers.length;
  let corruptedRows = 0;
  let nonStandardDates = 0;

  for (const row of rows.slice(1)) {
    if (row.length !== expected) corruptedRows += 1;
    for (const value of row) {
      const text = String(value ?? "");
      if (/\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/.test(text) && !/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(text.trim())) {
        nonStandardDates += 1;
      }
    }
  }

  const names = headers.map(normalise);
  const hasTransaction = names.some((h) => /reference|transaction|amount|debit|credit|balance/.test(h));
  const hasLead = names.some((h) => /email|phone|company|firstname|lastname|lead/.test(h));
  const hasNumeric = rows.slice(1, 50).some((row) => row.some((v) => /^[-+]?\d[\d,. ]*$/.test(String(v ?? "").trim())));

  const recommendations: SmartRecommendation[] = [];
  if (hasTransaction) recommendations.push({ id: "reconciliation", label: "Reconciliation Hub", reason: "Transaction/reference/amount fields detected." });
  if (hasLead) recommendations.push({ id: "cleaner", label: "Quick Data & CSV Cleaner", reason: "Lead/contact fields detected." });
  if (hasNumeric) recommendations.push({ id: "formulas", label: "Excel Automation", reason: "Numeric fields detected; formulas and totals may help." });
  if (!recommendations.length) recommendations.push({ id: "data-engine", label: "Local Data Engine", reason: "Inspect the schema and choose a workflow from the local engine." });
  return { rows: Math.max(0, rows.length - 1), columns: expected, corruptedRows, nonStandardDates, headers, recommendations: recommendations.slice(0, 3) };
}

export default function SmartDropzone({
  onFile,
  onRecommendation,
}: {
  onFile: (file: File, actions: SmartPipelineAction[]) => void;
  onRecommendation?: (recommendation: SmartRecommendation) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const [diagnostics, setDiagnostics] = useState<Diagnostics | null>(null);
  const [fileName, setFileName] = useState("");
  const [fileType, setFileType] = useState("");
  const [message, setMessage] = useState("");
  const [selectedActions, setSelectedActions] = useState<string[]>(["trim", "export-xlsx"]);

  const largeWarning = useMemo(() => diagnostics && diagnostics.rows > 25000, [diagnostics]);

  async function inspect(file: File) {
    setMessage("");
    const extension = file.name.split(".").pop()?.toLowerCase();
    if (!extension || !["csv", "xlsx", "pdf"].includes(extension)) {
      setMessage("Supported files: CSV, XLSX and PDF.");
      return;
    }
    setFileName(file.name);
    setFileType(extension.toUpperCase());

    if (extension === "pdf") {
      setDiagnostics({
        rows: 0, columns: 0, corruptedRows: 0, nonStandardDates: 0, headers: [],
        recommendations: [
          { id: "bank", label: "Bank Ledger X", reason: "PDF detected. Bank Ledger X can be used for statement workflows." },
          { id: "reconciliation", label: "Reconciliation Hub", reason: "Use reconciliation after transaction extraction." },
          { id: "data-engine", label: "Local Data Engine", reason: "Inspect extracted data locally after conversion." },
        ],
      });
      return;
    }

    try {
      if (extension === "csv") {
        const text = await file.text();
        const parsed = Papa.parse<unknown[]>(text, { skipEmptyLines: true }).data;
        setDiagnostics(analyseMatrix(parsed));
      } else {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", dense: true, cellDates: true, raw: false });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const matrix = sheet ? XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false }) : [];
        setDiagnostics(analyseMatrix(matrix));
      }
    } catch {
      setDiagnostics(null);
      setMessage("The file could not be inspected in this browser. You can still try the relevant workspace.");
    }
  }

  function chooseFile(file: File) {
    void inspect(file);
  }

  function toggleAction(id: string) {
    setSelectedActions((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id]);
  }

  function run() {
    if (!fileName || !inputRef.current?.files?.[0]) return;
    const actions = ACTIONS.filter((action) => selectedActions.includes(action.id));
    onFile(inputRef.current.files[0], actions);
  }

  return (
    <section className="smart-dropzone mx-auto max-w-6xl rounded-[28px] border border-line bg-canvas p-4 sm:p-6" aria-label="Smart file ingestion">
      <div
        className={`smart-dropzone__target ${dragging ? "is-dragging" : ""}`}
        onDragEnter={(e) => { e.preventDefault(); setDragging(true); }}
        onDragOver={(e) => e.preventDefault()}
        onDragLeave={(e) => { if (e.currentTarget === e.target) setDragging(false); }}
        onDrop={(e) => { e.preventDefault(); setDragging(false); const file = e.dataTransfer.files[0]; if (file) { if (inputRef.current) { const dt = new DataTransfer(); dt.items.add(file); inputRef.current.files = dt.files; } chooseFile(file); } }}
      >
        <input ref={inputRef} className="sr-only" type="file" accept=".csv,.xlsx,.pdf,text/csv,application/pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(e) => { const file = e.target.files?.[0]; if (file) chooseFile(file); }} />
        <div className="flex flex-col items-center text-center">
          <span className="smart-dropzone__icon" aria-hidden="true">↥</span>
          <p className="mt-4 text-xs font-medium uppercase tracking-[.2em] text-[rgb(var(--accent))]">Smart Drop</p>
          <h2 className="mt-2 text-2xl font-medium tracking-tight text-ink sm:text-3xl">Drop your data. MarqClean AI finds the right workflow.</h2>
          <p className="mt-3 max-w-2xl text-sm leading-6 text-ink-2">Upload CSV, XLSX or PDF. The first scan happens locally in your browser before any workflow is started.</p>
          <div className="mt-5 flex flex-wrap justify-center gap-2">
            {["CSV", "XLSX", "PDF"].map((format) => <span key={format} className="rounded-md bg-surface px-3 py-1 text-xs font-medium text-ink-2">{format}</span>)}
          </div>
          <button type="button" className="mt-6 rounded-md bg-[rgb(var(--accent))] px-6 py-3 text-sm font-medium text-white shadow-blue-200 transition hover:bg-[rgb(var(--accent-hover))]" onClick={() => inputRef.current?.click()}>Choose a file</button>
          <div className="mt-5 inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-4 py-2 text-xs font-medium text-emerald-800">
            <span aria-hidden="true">🔒</span> 100% Secure Client-Side Execution — your files never leave this browser.
          </div>
        </div>
      </div>

      {message ? <p className="mt-4 rounded-lg bg-red-50 p-4 text-sm font-medium text-red-700">{message}</p> : null}

      {diagnostics ? (
        <div className="mt-5 grid gap-4 lg:grid-cols-[1.15fr_.85fr]">
          <div className="rounded-lg border border-line bg-surface p-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div>
                <p className="text-xs font-medium uppercase tracking-[.18em] text-ink-3">Smart Scan Diagnostics</p>
                <h3 className="mt-1 text-lg font-medium text-ink">{fileName}</h3>
              </div>
              <span className="rounded-md bg-canvas px-3 py-1 text-xs font-medium text-ink-2">{fileType}</span>
            </div>
            <div className="mt-4 grid gap-3 sm:grid-cols-3">
              <div className="rounded-lg bg-canvas p-4"><p className="text-2xl font-medium text-ink">{diagnostics.corruptedRows}</p><p className="text-xs text-ink-3">Corrupted/inconsistent rows</p></div>
              <div className="rounded-lg bg-canvas p-4"><p className="text-2xl font-medium text-ink">{diagnostics.nonStandardDates}</p><p className="text-xs text-ink-3">Non-standard date values</p></div>
              <div className="rounded-lg bg-canvas p-4"><p className="text-2xl font-medium text-ink">{diagnostics.rows.toLocaleString()}</p><p className="text-xs text-ink-3">Data rows detected</p></div>
            </div>
            {largeWarning ? <div className="mt-4 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm font-medium text-amber-900">Large dataset detected. Processing locally may take roughly ~15 seconds depending on your computer and browser capacity.</div> : null}
            {diagnostics.headers.length ? <p className="mt-4 text-xs leading-5 text-ink-3"><strong>Detected columns:</strong> {diagnostics.headers.slice(0, 12).join(" · ")}{diagnostics.headers.length > 12 ? " · …" : ""}</p> : null}
          </div>

          <div className="rounded-lg border border-line bg-canvas p-5">
            <p className="text-xs font-medium uppercase tracking-[.18em] text-ink-3">Recommended Automations</p>
            <div className="mt-3 space-y-2">
              {diagnostics.recommendations.map((item) => (
                <button key={item.id} type="button" onClick={() => onRecommendation?.(item)} className="w-full rounded-lg border border-line bg-surface p-3 text-left transition hover:border-blue-200 hover:bg-blue-50">
                  <span className="block text-sm font-medium text-ink">▶ {item.label}</span>
                  <span className="mt-1 block text-xs leading-5 text-ink-3">{item.reason}</span>
                </button>
              ))}
            </div>
            <p className="mt-5 text-xs font-medium uppercase tracking-[.18em] text-ink-3">Workflow Macro</p>
            <div className="mt-3 space-y-2">
              {ACTIONS.map((action) => (
                <label key={action.id} className="flex cursor-pointer items-center gap-3 rounded-lg border border-line px-3 py-2.5 text-sm text-ink-2">
                  <input type="checkbox" checked={selectedActions.includes(action.id)} onChange={() => toggleAction(action.id)} />
                  <span>{action.label}</span>
                </label>
              ))}
            </div>
            <button type="button" disabled={fileType === "PDF"} onClick={run} className="mt-4 w-full rounded-md bg-slate-950 px-5 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-40">Run workflow locally</button>
            {fileType === "PDF" ? <p className="mt-2 text-xs text-ink-3">Open Bank Ledger X for PDF extraction before running a data workflow.</p> : null}
          </div>
        </div>
      ) : null}
    </section>
  );
}
