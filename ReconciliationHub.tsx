import { AnimatePresence, motion } from "framer-motion";
import { useMemo, useRef, useState } from "react";
import { GlobalHeader, ToolLaunchpad } from "./WorkspaceShell";
import SmartSheetManager from "./SmartSheetManager";
import {
  autoDetectMapping,
  buildReconciliationWorkbook,
  detectFileKind,
  downloadBlob,
  extractToTable,
  MACRO_LIBRARY,
  PLATFORM_FIELDS,
  runReconciliation,
  type CompareFieldConfig,
  type DataTable,
  type MacroDefinition,
  type ReconResult,
} from "./engine";
import {
  autoDetectMailingMapping,
  buildMailingWorkbook,
  extractMailingRecordsFromPdf,
  extractMailingRecordsFromTable,
  MAILING_FIELDS,
  runMailingVerification,
  type ExtractedMailingRecord,
  type MailingFieldKey,
  type MailingMapping,
  type MailingResult,
} from "./mailing";

const MAX_FILE_SIZE_BYTES = 2 * 1024 * 1024 * 1024;

type NavKey =
  | "dashboard"
  | "bank"
  | "mailing"
  | "comparison"
  | "cleansing"
  | "reconciliation"
  | "compliance"
  | "documents"
  | "reporting"
  | "administration"
  | "sheet-manager";

const NAV_ITEMS: Array<{ key: NavKey; label: string; blurb: string; icon: string }> = [
  { key: "dashboard", label: "Dashboard", blurb: "Processing statistics, exception summary, and history.", icon: "📊" },
  { key: "bank", label: "Bank Ledger X", blurb: "Bank statement and ledger reconciliation module.", icon: "🏦" },
  { key: "mailing", label: "Data Verification", blurb: "Verify PDF client mailing data against Excel master data.", icon: "🔎" },
  { key: "comparison", label: "File Comparison Centre", blurb: "Excel, CSV, PDF, and Word comparison in every mode.", icon: "🔁" },
  { key: "cleansing", label: "Data Cleansing Centre", blurb: "Duplicates, names, addresses, phones, and formatting.", icon: "🧹" },
  { key: "reconciliation", label: "Reconciliation Services", blurb: "Client, investor, account, banking, and register recon.", icon: "🔄" },
  { key: "compliance", label: "Compliance Validation", blurb: "KYC/FICA, CRS/FATCA, PEP, trust, and director checks.", icon: "✅" },
  { key: "documents", label: "Document Validation Centre", blurb: "Application, mandate, banking, and identity validation.", icon: "📄" },
  { key: "reporting", label: "Reporting Centre", blurb: "Reconciliation, exception, audit, and quality reports.", icon: "📈" },
  { key: "administration", label: "Administration", blurb: "Templates, field mapping, audit logs, and settings.", icon: "⚙️" },
  { key: "sheet-manager", label: "Smart Sheet Manager", blurb: "Analyse, keep, delete, rename, and sort worksheets in any workbook.", icon: "📑" },
];

const STATUS_COLORS: Record<string, string> = {
  "Exact Match": "bg-surface text-ink-2",
  "Partial Match": "bg-accent-tint text-accent",
  "Mismatch": "bg-surface text-error",
  "Missing Value": "bg-surface text-ink-2",
  "Unmatched Source": "bg-surface text-ink-2",
  "Unmatched Target": "bg-surface text-ink-2",
  "Duplicate": "bg-surface text-ink-2",
};

type HistoryEntry = {
  time: string;
  source: string;
  target: string;
  records: number;
  exceptions: number;
  matchPct: number;
};

export default function ReconciliationHub({ onExit }: { onExit: () => void }) {
  const [nav, setNav] = useState<NavKey>("mailing");
  const [sourceTable, setSourceTable] = useState<DataTable | null>(null);
  const [targetTable, setTargetTable] = useState<DataTable | null>(null);
  const [sourceMapping, setSourceMapping] = useState<Record<number, string>>({});
  const [targetMapping, setTargetMapping] = useState<Record<number, string>>({});
  const [selectedFields, setSelectedFields] = useState<Set<string>>(new Set());
  const [keyField, setKeyField] = useState<string>("");
  const [result, setResult] = useState<ReconResult | null>(null);
  const [error, setError] = useState("");
  const [processing, setProcessing] = useState(false);
  const [macroSearch, setMacroSearch] = useState("");
  const [selectedMacro, setSelectedMacro] = useState<string>("");
  const [history, setHistory] = useState<HistoryEntry[]>([]);

  const sourceInputRef = useRef<HTMLInputElement>(null);
  const targetInputRef = useRef<HTMLInputElement>(null);

  // Mailing verification state
  const [mailingRecords, setMailingRecords] = useState<ExtractedMailingRecord[] | null>(null);
  const [mailingPdfName, setMailingPdfName] = useState("");
  const [mailingExcel, setMailingExcel] = useState<DataTable | null>(null);
  const [mailingMapping, setMailingMapping] = useState<MailingMapping>({});
  const [mailingFields, setMailingFields] = useState<Set<MailingFieldKey>>(new Set(["fullName", "addressLine1", "postalCode"]));
  const [mailingKey, setMailingKey] = useState<MailingFieldKey>("fullName");
  const [mailingResult, setMailingResult] = useState<MailingResult | null>(null);
  const mailingPdfRef = useRef<HTMLInputElement>(null);
  const mailingExcelRef = useRef<HTMLInputElement>(null);

  async function loadMailingPdf(file: File) {
    setError("");
    setProcessing(true);
    try {
      const kind = detectFileKind(file.name);
      let records: ExtractedMailingRecord[];
      if (kind === "pdf") {
        records = await extractMailingRecordsFromPdf(file);
      } else if (kind === "csv" || kind === "text" || kind === "word" || kind === "excel") {
        records = extractMailingRecordsFromTable(await extractToTable(file));
      } else {
        throw new Error("Upload a PDF, CSV, Word, or Excel file for extraction.");
      }
      if (!records.length) throw new Error("No mailing records could be extracted. Try a text-based PDF or paste the data.");
      setMailingRecords(records);
      setMailingPdfName(file.name);
      setMailingResult(null);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Extraction failed.");
    } finally {
      setProcessing(false);
    }
  }

  async function loadMailingExcel(file: File) {
    setError("");
    setProcessing(true);
    try {
      const table = await extractToTable(file);
      if (!table.headers.length) throw new Error("No readable Excel data found.");
      setMailingExcel(table);
      setMailingMapping(autoDetectMailingMapping(table));
      setMailingResult(null);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "The Excel master could not be read.");
    } finally {
      setProcessing(false);
    }
  }

  function toggleMailingField(key: MailingFieldKey) {
    setMailingFields((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  function runMailing() {
    setError("");
    if (!mailingRecords || !mailingExcel) {
      setError("Upload both a PDF (or source file) and an Excel master file.");
      return;
    }
    const fields = Array.from(mailingFields);
    if (!fields.length) {
      setError("Select at least one field to verify.");
      return;
    }
    const mappedKeys = new Set(Object.values(mailingMapping).filter(Boolean));
    const usableFields = fields.filter((key) => mappedKeys.has(key));
    if (!usableFields.length) {
      setError("Map the Excel columns to mailing fields first (Administration or Auto Detect).");
      return;
    }
    const effectiveKey = usableFields.includes(mailingKey) ? mailingKey : usableFields[0];
    setProcessing(true);
    window.setTimeout(() => {
      try {
        const result = runMailingVerification(mailingRecords, mailingExcel, mailingMapping, usableFields, effectiveKey, mailingPdfName);
        setMailingResult(result);
        setHistory((prev) => [
          { time: new Date().toLocaleString(), source: mailingPdfName, target: mailingExcel.sourceName, records: result.summary.pdfRecords, exceptions: result.summary.exceptions, matchPct: result.summary.matchPercentage },
          ...prev,
        ].slice(0, 20));
      } catch (caughtError) {
        setError(caughtError instanceof Error ? caughtError.message : "Verification failed.");
      } finally {
        setProcessing(false);
      }
    }, 30);
  }

  function downloadMailingReport() {
    if (!mailingResult || !mailingExcel) return;
    const workbook = buildMailingWorkbook(mailingResult, mailingPdfName, mailingExcel.sourceName);
    downloadBlob(workbook, "mailing-verification-report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  const mailingMappedKeys = useMemo(() => new Set(Object.values(mailingMapping).filter(Boolean)), [mailingMapping]);

  const commonMappedFields = useMemo(() => {
    const sourceKeys = new Set(Object.values(sourceMapping).filter(Boolean));
    const targetKeys = new Set(Object.values(targetMapping).filter(Boolean));
    return PLATFORM_FIELDS.filter((f) => sourceKeys.has(f.key) && targetKeys.has(f.key));
  }, [sourceMapping, targetMapping]);

  const filteredMacros = useMemo(() => {
    const q = macroSearch.trim().toLowerCase();
    if (!q) return MACRO_LIBRARY;
    return MACRO_LIBRARY.filter(
      (m) => m.name.toLowerCase().includes(q) || m.category.toLowerCase().includes(q) || m.description.toLowerCase().includes(q)
    );
  }, [macroSearch]);

  async function loadFile(file: File, target: "source" | "target") {
    setError("");
    setProcessing(true);
    try {
      if (detectFileKind(file.name) === "unknown") throw new Error("Unsupported file. Upload PDF, Excel, CSV, or Word.");
      if (file.size > MAX_FILE_SIZE_BYTES) throw new Error("File exceeds the 2 GB limit.");

      const table = await extractToTable(file);
      if (!table.headers.length) throw new Error("No readable data found in the file.");

      const mapping = autoDetectMapping(table);
      if (target === "source") {
        setSourceTable(table);
        setSourceMapping(mapping);
      } else {
        setTargetTable(table);
        setTargetMapping(mapping);
      }
      setResult(null);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "The file could not be processed.");
    } finally {
      setProcessing(false);
    }
  }

  function applyMacro(macro: MacroDefinition) {
    setSelectedMacro(macro.id);
    setSelectedFields(new Set(macro.fieldKeys));
    setKeyField(macro.keyField);
  }

  function toggleField(key: string) {
    setSelectedFields((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  function columnForField(mapping: Record<number, string>, fieldKey: string): number {
    const entry = Object.entries(mapping).find(([, key]) => key === fieldKey);
    return entry ? Number(entry[0]) : -1;
  }

  function runComparison() {
    setError("");
    if (!sourceTable || !targetTable) {
      setError("Upload both a source file and a target file first.");
      return;
    }
    const fields = Array.from(selectedFields).filter((k) => k);
    if (!fields.length) {
      setError("Select at least one field to compare, or choose a macro.");
      return;
    }

    const configs: CompareFieldConfig[] = [];
    fields.forEach((fieldKey) => {
      const sourceColumn = columnForField(sourceMapping, fieldKey);
      const targetColumn = columnForField(targetMapping, fieldKey);
      if (sourceColumn >= 0 && targetColumn >= 0) {
        configs.push({ fieldKey, sourceColumn, targetColumn });
      }
    });

    if (!configs.length) {
      setError("The selected fields are not mapped in both files. Map columns to platform fields first.");
      return;
    }

    const effectiveKey = configs.find((c) => c.fieldKey === keyField) ? keyField : configs[0].fieldKey;
    setProcessing(true);
    window.setTimeout(() => {
      try {
        const reconResult = runReconciliation(sourceTable, targetTable, configs, effectiveKey);
        setResult(reconResult);
        setHistory((prev) => [
          {
            time: new Date().toLocaleString(),
            source: sourceTable.sourceName,
            target: targetTable.sourceName,
            records: reconResult.summary.totalRecords,
            exceptions: reconResult.summary.exceptions,
            matchPct: reconResult.summary.matchPercentage,
          },
          ...prev,
        ].slice(0, 20));
        setNav("reporting");
      } catch (caughtError) {
        setError(caughtError instanceof Error ? caughtError.message : "Reconciliation failed.");
      } finally {
        setProcessing(false);
      }
    }, 30);
  }

  function downloadReport() {
    if (!result || !sourceTable || !targetTable) return;
    const workbook = buildReconciliationWorkbook(result, sourceTable.sourceName, targetTable.sourceName);
    downloadBlob(workbook, "reconciliation-report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  return (
    <div className="ws-root ws-scope relative min-h-screen">
      <GlobalHeader />
      <div className="ws-grid pointer-events-none absolute inset-x-0 top-0 h-[38rem]" aria-hidden="true" />

      <div className="relative mx-auto mb-2 flex max-w-[110rem] flex-wrap items-center justify-between gap-3 px-5 pt-24 lg:px-8">
        <div className="flex items-center gap-3">
          <span className="flex h-10 w-10 items-center justify-center rounded-xl bg-canvas from-[rgb(var(--accent))]/20 to-[rgb(var(--accent))]/20 text-xl ring-1 ring-inset ring-white/10">🔄</span>
          <div>
            <h2 className="mc-display text-xl font-medium text-ink">Reconciliation Hub</h2>
            <p className="text-xs text-[#b0bacb]">AI-powered reconciliation, validation and record matching. Runs 100% in your browser.</p>
          </div>
        </div>
        <button className="ws-btn-secondary rounded-md px-4 py-2 text-sm font-medium" onClick={onExit}>Back to MarqClean AI</button>
      </div>

      <div className="relative mx-auto max-w-[110rem] px-5 pb-16 pt-4 lg:px-8">
        <div className="mb-2">
          <p className="mc-mono text-[11px] uppercase tracking-[0.2em] text-[#8a94a8]">Modules Available ({NAV_ITEMS.length})</p>
        </div>
        <ToolLaunchpad tools={NAV_ITEMS} active={nav} onSelect={setNav} />

        <main id="main-content" className="mt-6 min-w-0">
          {error ? <p className="mb-5 rounded-xl bg-surface p-4 text-sm font-medium text-error" role="alert">{error}</p> : null}

          <AnimatePresence>
            {processing ? (
              <motion.div className="mb-5 overflow-hidden rounded-xl bg-canvas shadow" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
                <div className="p-4 text-sm font-medium text-ink-2" role="status">Processing files and reconciling records...</div>
                <motion.div className="h-1 bg-accent-tint" initial={{ width: "10%" }} animate={{ width: "100%" }} transition={{ duration: 1, repeat: Infinity, repeatType: "reverse" }} />
              </motion.div>
            ) : null}
          </AnimatePresence>

          {nav === "bank" ? (
            <BankLedgerView />
          ) : nav === "mailing" ? (
            <MailingView
              records={mailingRecords}
              pdfName={mailingPdfName}
              excel={mailingExcel}
              mapping={mailingMapping}
              setMapping={setMailingMapping}
              fields={mailingFields}
              toggleField={toggleMailingField}
              matchKey={mailingKey}
              setMatchKey={setMailingKey}
              result={mailingResult}
              mappedKeys={mailingMappedKeys}
              pdfRef={mailingPdfRef}
              excelRef={mailingExcelRef}
              onPdf={loadMailingPdf}
              onExcel={loadMailingExcel}
              onRun={runMailing}
              onDownload={downloadMailingReport}
              processing={processing}
            />
          ) : nav === "dashboard" ? (
            <section>
              <h1 className="text-3xl font-medium tracking-[-0.03em]">Processing Dashboard</h1>
              <p className="mt-2 max-w-2xl text-ink-2">Reconciliation statistics, exception summary, data quality metrics, and processing history for this browser session.</p>
              <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                <DashCard label="Runs this session" value={history.length} />
                <DashCard label="Last match %" value={history[0]?.matchPct ?? 0} suffix="%" />
                <DashCard label="Last exceptions" value={history[0]?.exceptions ?? 0} />
                <DashCard label="Last records" value={history[0]?.records ?? 0} />
              </div>
              <div className="mt-8 rounded-xl border border-line bg-canvas p-5">
                <h2 className="text-lg font-medium">Processing history</h2>
                {history.length ? (
                  <div className="mt-4 overflow-x-auto">
                    <table className="min-w-full text-left text-sm">
                      <thead className="text-xs uppercase tracking-wide text-ink-3">
                        <tr><th className="py-2 pr-4">Time</th><th className="py-2 pr-4">Source</th><th className="py-2 pr-4">Target</th><th className="py-2 pr-4">Records</th><th className="py-2 pr-4">Exceptions</th><th className="py-2">Match %</th></tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {history.map((entry, index) => (
                          <tr key={index}>
                            <td className="py-2 pr-4 text-ink-2">{entry.time}</td>
                            <td className="py-2 pr-4">{entry.source}</td>
                            <td className="py-2 pr-4">{entry.target}</td>
                            <td className="py-2 pr-4">{entry.records}</td>
                            <td className="py-2 pr-4">{entry.exceptions}</td>
                            <td className="py-2">{entry.matchPct}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <p className="mt-3 text-sm text-ink-3">No runs yet. Start in the File Comparison Centre.</p>
                )}
              </div>
            </section>
          ) : nav === "reporting" && result ? (
            <ReportView result={result} onDownload={downloadReport} />
          ) : nav === "administration" ? (
            <AdministrationView sourceTable={sourceTable} targetTable={targetTable} sourceMapping={sourceMapping} targetMapping={targetMapping} setSourceMapping={setSourceMapping} setTargetMapping={setTargetMapping} />
          ) : nav === "sheet-manager" ? (
            <SmartSheetManager />
          ) : (
            <section>
              <h1 className="text-3xl font-medium tracking-[-0.03em]">
                {NAV_ITEMS.find((i) => i.key === nav)?.label}
              </h1>
              <p className="mt-2 max-w-3xl text-ink-2">
                Upload a source and target file, auto-detect the structure, choose a macro or select fields, then run a colour-coded reconciliation. Every mode (Excel, CSV, PDF, Word, and multi-file) uses the same engine.
              </p>

              {/* Macro selector */}
              <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <h2 className="text-lg font-medium">Reconciliation macros</h2>
                  <input
                    className="w-64 max-w-full rounded-md border border-line px-4 py-2 text-sm"
                    placeholder="Search macros (e.g. banking, KYC, register)"
                    value={macroSearch}
                    onChange={(event) => setMacroSearch(event.target.value)}
                    aria-label="Search reconciliation macros"
                  />
                </div>
                <div className="mt-3 flex flex-wrap items-center gap-3">
                  <label className="text-sm text-ink-2">Quick select</label>
                  <select
                    className="rounded-lg border border-line px-3 py-2 text-sm"
                    value={selectedMacro}
                    onChange={(event) => {
                      const macro = MACRO_LIBRARY.find((m) => m.id === event.target.value);
                      if (macro) applyMacro(macro);
                    }}
                  >
                    <option value="">Choose a macro...</option>
                    {MACRO_LIBRARY.map((macro) => (
                      <option key={macro.id} value={macro.id}>{macro.name}</option>
                    ))}
                  </select>
                </div>
                <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
                  {filteredMacros.map((macro) => (
                    <button
                      key={macro.id}
                      className={`rounded-xl border p-4 text-left transition ${selectedMacro === macro.id ? "border-line bg-accent-tint" : "border-line hover:border-slate-400"}`}
                      onClick={() => applyMacro(macro)}
                    >
                      <span className="block text-xs uppercase tracking-wide text-ink-3">{macro.category}</span>
                      <span className="mt-1 block text-sm font-medium text-ink">{macro.name}</span>
                      <span className="mt-1 block text-xs text-ink-3">{macro.description}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Upload */}
              <div className="mt-6 grid gap-4 md:grid-cols-2">
                <UploadCard title="Source file" table={sourceTable} inputRef={sourceInputRef} onPick={() => sourceInputRef.current?.click()} onFile={(f) => loadFile(f, "source")} />
                <UploadCard title="Target file" table={targetTable} inputRef={targetInputRef} onPick={() => targetInputRef.current?.click()} onFile={(f) => loadFile(f, "target")} />
              </div>

              {/* Field selection */}
              {sourceTable && targetTable ? (
                <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <h2 className="text-lg font-medium">Fields to compare</h2>
                    <div className="flex items-center gap-2 text-sm">
                      <label className="text-ink-2">Match key</label>
                      <select className="rounded-lg border border-line px-3 py-1.5" value={keyField} onChange={(event) => setKeyField(event.target.value)}>
                        <option value="">Auto (first field)</option>
                        {commonMappedFields.filter((f) => selectedFields.has(f.key)).map((f) => (
                          <option key={f.key} value={f.key}>{f.label}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <p className="mt-1 text-xs text-ink-3">Only fields mapped in both files can be compared. Use Administration to adjust mapping.</p>
                  <div className="mt-4 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                    {commonMappedFields.length ? commonMappedFields.map((field) => (
                      <label key={field.key} className="flex items-center gap-2 rounded-lg border border-line px-3 py-2 text-sm">
                        <input type="checkbox" checked={selectedFields.has(field.key)} onChange={() => toggleField(field.key)} />
                        <span>{field.label}</span>
                        <span className="ml-auto text-xs text-ink-3">{field.group}</span>
                      </label>
                    )) : (
                      <p className="text-sm text-ink-3">No overlapping mapped fields detected. Open Administration to map columns manually.</p>
                    )}
                  </div>
                  <button
                    className="mt-5 rounded-md bg-ink px-6 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint"
                    onClick={runComparison}
                    disabled={processing}
                  >
                    Run reconciliation
                  </button>
                </div>
              ) : null}

            </section>
          )}

          {nav === "reporting" && !result ? (
            <p className="mt-6 rounded-xl border border-line bg-canvas p-6 text-sm text-ink-2">Run a comparison to generate reconciliation, exception, audit, and data quality reports.</p>
          ) : null}
        </main>
      </div>
    </div>
  );
}

function BankLedgerView() {
  function goHomeBankTab() {
    window.history.pushState({}, "", "/");
    window.dispatchEvent(new PopStateEvent("popstate"));
    window.setTimeout(() => {
      window.dispatchEvent(new CustomEvent("marqclean:open-bank"));
      document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" });
    }, 80);
  }
  return (
    <section>
      <h1 className="text-3xl font-medium tracking-[-0.03em]">Bank Ledger X</h1>
      <p className="mt-2 max-w-3xl text-[#b0bacb]">
        Bank statement and ledger reconciliation module within Reconciliation Hub. Convert PDF bank statements into clean
        Excel, CSV and QIF, then reconcile transactions against ledger records with AI-powered matching and exception
        reporting.
      </p>
      <div className="mt-4 flex flex-wrap gap-2">
        <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-medium uppercase">PDF</span>
        <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-medium uppercase">XLSX</span>
        <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-medium uppercase">CSV</span>
        <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-medium uppercase">QIF</span>
      </div>

      <div className="mt-6 grid gap-4 lg:grid-cols-2">
        <div className="ws-surface rounded-xl p-6">
          <h2 className="text-lg font-medium text-ink">Convert a bank statement</h2>
          <p className="mt-2 text-sm text-[#b0bacb]">Turn a PDF statement into a structured, import-ready spreadsheet with normalized dates and amounts.</p>
          <button className="ws-btn-primary mt-4 rounded-md px-5 py-2.5 text-sm font-medium" onClick={goHomeBankTab}>Open Bank Statement Converter</button>
        </div>
        <div className="ws-surface rounded-xl p-6">
          <h2 className="text-lg font-medium text-ink">Reconcile against ledger</h2>
          <p className="mt-2 text-sm text-[#b0bacb]">Use the reconciliation modules to match statement transactions against ledger and account records, and export exceptions.</p>
          <div className="mt-3 flex flex-wrap gap-2 text-xs text-[#8a94a8]">
            <span>✓ Client Reconciliation</span>
            <span>✓ Investor Reconciliation</span>
            <span>✓ Register Reconciliation</span>
            <span>✓ Account Matching</span>
            <span>✓ Exception Reporting</span>
          </div>
        </div>
      </div>
    </section>
  );
}

function DashCard({ label, value, suffix }: { label: string; value: number; suffix?: string }) {
  return (
    <div className="rounded-xl border border-line bg-canvas p-5">
      <p className="text-3xl font-medium text-ink">{value}{suffix ?? ""}</p>
      <p className="mt-1 text-xs uppercase tracking-wide text-ink-3">{label}</p>
    </div>
  );
}

function UploadCard({
  title, table, inputRef, onPick, onFile,
}: {
  title: string;
  table: DataTable | null;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onPick: () => void;
  onFile: (file: File) => void;
}) {
  return (
    <div
      className="rounded-xl border border-dashed border-line bg-canvas p-5 text-center transition hover:border-line"
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const file = event.dataTransfer.files?.[0];
        if (file) onFile(file);
      }}
    >
      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept=".pdf,.xlsx,.xls,.csv,.txt,.docx,.doc"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
          if (inputRef.current) inputRef.current.value = "";
        }}
      />
      <h3 className="text-base font-medium text-ink">{title}</h3>
      {table ? (
        <div className="mt-2">
          <p className="text-sm text-ink-2">{table.sourceName}</p>
          <p className="text-xs text-ink-3">{table.rows.length} rows, {table.headers.length} columns</p>
        </div>
      ) : (
        <p className="mt-2 text-xs text-ink-3">Drop PDF, Excel, CSV, or Word, or choose a file.</p>
      )}
      <button className="mt-4 rounded-md bg-ink px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={onPick}>
        {table ? "Replace file" : "Choose file"}
      </button>
    </div>
  );
}

function ReportView({ result, onDownload }: { result: ReconResult; onDownload: () => void }) {
  const { summary } = result;
  const [showExceptionsOnly, setShowExceptionsOnly] = useState(false);

  const rows = useMemo(() => {
    const flattened: Array<{ key: string; field: string; source: string; target: string; status: string; comment: string }> = [];
    result.records.forEach((record) => {
      record.fields.forEach((field) => {
        if (showExceptionsOnly && field.status === "Exact Match") return;
        flattened.push({ key: record.key, field: field.fieldLabel, source: field.sourceValue, target: field.targetValue, status: field.status, comment: field.comment });
      });
    });
    return flattened.slice(0, 400);
  }, [result, showExceptionsOnly]);

  return (
    <section>
      <div className="flex flex-wrap items-center justify-between gap-3">
        <h1 className="text-3xl font-medium tracking-[-0.03em]">Reporting Centre</h1>
        <button className="rounded-md bg-accent-tint px-5 py-2.5 text-sm font-medium text-ink transition hover:bg-accent-tint" onClick={onDownload}>
          Download Excel report (4 sheets)
        </button>
      </div>

      <div className="mt-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <DashCard label="Total records" value={summary.totalRecords} />
        <DashCard label="Match %" value={summary.matchPercentage} suffix="%" />
        <DashCard label="Exceptions" value={summary.exceptions} />
        <DashCard label="Duplicates" value={summary.duplicates} />
      </div>

      <div className="mt-4 grid gap-3 sm:grid-cols-5">
        <Legend color="bg-surface text-ink-2" label={`Exact ${summary.matched}`} />
        <Legend color="bg-accent-tint text-accent" label={`Partial ${summary.partial}`} />
        <Legend color="bg-surface text-error" label={`Mismatch ${summary.mismatched}`} />
        <Legend color="bg-surface text-ink-2" label={`Missing ${summary.missing}`} />
        <Legend color="bg-surface text-ink-2" label={`Unmatched ${summary.unmatched}`} />
      </div>

      <div className="mt-6 rounded-xl border border-line bg-canvas">
        <div className="flex items-center justify-between border-b border-line px-5 py-3">
          <h2 className="text-lg font-medium">Detailed reconciliation</h2>
          <label className="flex items-center gap-2 text-sm text-ink-2">
            <input type="checkbox" checked={showExceptionsOnly} onChange={(event) => setShowExceptionsOnly(event.target.checked)} />
            Exceptions only
          </label>
        </div>
        <div className="max-h-[32rem] overflow-auto">
          <table className="min-w-full text-left text-sm">
            <thead className="sticky top-0 bg-surface text-xs uppercase tracking-wide text-ink-3">
              <tr>
                <th className="px-4 py-2">Record</th><th className="px-4 py-2">Field</th><th className="px-4 py-2">Source</th>
                <th className="px-4 py-2">Target</th><th className="px-4 py-2">Status</th><th className="px-4 py-2">Comment</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {rows.map((row, index) => (
                <tr key={index}>
                  <td className="max-w-40 truncate px-4 py-2 text-ink-2">{row.key}</td>
                  <td className="px-4 py-2 text-ink-2">{row.field}</td>
                  <td className="max-w-48 truncate px-4 py-2 text-ink-2">{row.source || "-"}</td>
                  <td className="max-w-48 truncate px-4 py-2 text-ink-2">{row.target || "-"}</td>
                  <td className="px-4 py-2"><span className={`rounded-md px-2 py-1 text-xs font-medium ${STATUS_COLORS[row.status] ?? "bg-surface text-ink-2"}`}>{row.status}</span></td>
                  <td className="max-w-48 truncate px-4 py-2 text-ink-3">{row.comment}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {result.quality.length ? (
        <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
          <h2 className="text-lg font-medium">Data quality report ({result.quality.length})</h2>
          <div className="mt-3 max-h-64 overflow-auto">
            <table className="min-w-full text-left text-sm">
              <thead className="text-xs uppercase tracking-wide text-ink-3"><tr><th className="py-2 pr-4">Location</th><th className="py-2 pr-4">Field</th><th className="py-2">Issue</th></tr></thead>
              <tbody className="divide-y divide-slate-100">
                {result.quality.slice(0, 200).map((issue, index) => (
                  <tr key={index}><td className="py-2 pr-4 text-ink-2">{issue.rowRef}</td><td className="py-2 pr-4">{issue.field}</td><td className="py-2 text-ink-2">{issue.issue}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : null}
    </section>
  );
}

function Legend({ color, label }: { color: string; label: string }) {
  return <div className={`rounded-lg px-3 py-2 text-center text-xs font-medium ${color}`}>{label}</div>;
}

function AdministrationView({
  sourceTable, targetTable, sourceMapping, targetMapping, setSourceMapping, setTargetMapping,
}: {
  sourceTable: DataTable | null;
  targetTable: DataTable | null;
  sourceMapping: Record<number, string>;
  targetMapping: Record<number, string>;
  setSourceMapping: (m: Record<number, string>) => void;
  setTargetMapping: (m: Record<number, string>) => void;
}) {
  return (
    <section>
      <h1 className="text-3xl font-medium tracking-[-0.03em]">Administration</h1>
      <p className="mt-2 max-w-3xl text-ink-2">Field mapping, templates, and configuration. Map each uploaded column to a platform field. Auto Detect Mapping is applied on upload; adjust it here as needed.</p>
      <div className="mt-6 grid gap-6 lg:grid-cols-2">
        <MappingTable title="Source mapping" table={sourceTable} mapping={sourceMapping} setMapping={setSourceMapping} />
        <MappingTable title="Target mapping" table={targetTable} mapping={targetMapping} setMapping={setTargetMapping} />
      </div>
    </section>
  );
}

function MappingTable({
  title, table, mapping, setMapping,
}: {
  title: string;
  table: DataTable | null;
  mapping: Record<number, string>;
  setMapping: (m: Record<number, string>) => void;
}) {
  if (!table) return <div className="rounded-xl border border-line bg-canvas p-5 text-sm text-ink-3">{title}: upload a file first.</div>;
  return (
    <div className="rounded-xl border border-line bg-canvas p-5">
      <h2 className="text-lg font-medium">{title}</h2>
      <p className="text-xs text-ink-3">{table.sourceName}</p>
      <div className="mt-3 max-h-96 space-y-2 overflow-auto">
        {table.headers.map((header, index) => (
          <div key={index} className="flex items-center gap-3">
            <span className="w-1/2 truncate text-sm text-ink-2" title={header}>{header || `Column ${index + 1}`}</span>
            <select
              className="w-1/2 rounded border border-line px-2 py-1 text-sm"
              value={mapping[index] ?? ""}
              onChange={(event) => setMapping({ ...mapping, [index]: event.target.value })}
            >
              <option value="">Not mapped</option>
              {PLATFORM_FIELDS.map((field) => (
                <option key={field.key} value={field.key}>{field.label}</option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
  );
}

function MailingUpload({
  title, subtitle, inputRef, onFile, filledLabel, onPick,
}: {
  title: string;
  subtitle: string;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onFile: (file: File) => void;
  filledLabel: string;
  onPick: () => void;
}) {
  return (
    <div
      className="rounded-xl border border-dashed border-line bg-canvas p-5 text-center transition hover:border-line"
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const file = event.dataTransfer.files?.[0];
        if (file) onFile(file);
      }}
    >
      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept=".pdf,.xlsx,.xls,.csv,.txt,.docx,.doc"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
          if (inputRef.current) inputRef.current.value = "";
        }}
      />
      <h3 className="text-base font-medium text-ink">{title}</h3>
      <p className="mt-1 text-xs text-ink-3">{filledLabel || subtitle}</p>
      <button className="mt-4 rounded-md bg-ink px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={onPick}>
        {filledLabel ? "Replace file" : "Choose file"}
      </button>
    </div>
  );
}

function MailingView(props: {
  records: ExtractedMailingRecord[] | null;
  pdfName: string;
  excel: DataTable | null;
  mapping: MailingMapping;
  setMapping: (m: MailingMapping) => void;
  fields: Set<MailingFieldKey>;
  toggleField: (key: MailingFieldKey) => void;
  matchKey: MailingFieldKey;
  setMatchKey: (key: MailingFieldKey) => void;
  result: MailingResult | null;
  mappedKeys: Set<string>;
  pdfRef: React.RefObject<HTMLInputElement | null>;
  excelRef: React.RefObject<HTMLInputElement | null>;
  onPdf: (file: File) => void;
  onExcel: (file: File) => void;
  onRun: () => void;
  onDownload: () => void;
  processing: boolean;
}) {
  const { records, excel, mapping, setMapping, fields, toggleField, matchKey, setMatchKey, result, mappedKeys } = props;
  const [showExceptionsOnly, setShowExceptionsOnly] = useState(false);

  const flatRows = useMemo(() => {
    if (!result) return [];
    const out: Array<{ record: number; row: string; field: string; pdf: string; excel: string; status: string; comment: string }> = [];
    result.records.forEach((record) => {
      record.fields.forEach((field) => {
        if (showExceptionsOnly && field.status === "Exact Match") return;
        out.push({ record: record.recordNumber, row: record.excelRow ? String(record.excelRow) : "-", field: field.field, pdf: field.pdfValue, excel: field.excelValue, status: field.status, comment: field.comments });
      });
    });
    return out.slice(0, 500);
  }, [result, showExceptionsOnly]);

  return (
    <section>
      <div className="mx-auto max-w-3xl text-center">
        <h1 className="text-3xl font-medium tracking-[-0.03em]">Data Verification</h1>
        <p className="mt-2 text-justify leading-7 text-ink-2 [hyphens:auto]">
          Upload an unstructured PDF client mailing list and a structured Excel master file. The platform extracts and
          structures the PDF into Full Name, Title, Address Lines, and Postal Code, maps them to the Excel columns, then
          verifies each record and produces a colour-coded report.
        </p>
      </div>

      <div className="mx-auto mt-6 grid max-w-3xl gap-4 md:grid-cols-2">
        <MailingUpload
          title="PDF mailing source"
          subtitle="Drop a PDF, CSV, Word, or Excel file to extract from."
          inputRef={props.pdfRef}
          onFile={props.onPdf}
          onPick={() => props.pdfRef.current?.click()}
          filledLabel={records ? `${props.pdfName} - ${records.length} records extracted` : ""}
        />
        <MailingUpload
          title="Excel master data"
          subtitle="Drop the structured Excel master file."
          inputRef={props.excelRef}
          onFile={props.onExcel}
          onPick={() => props.excelRef.current?.click()}
          filledLabel={excel ? `${excel.sourceName} - ${excel.rows.length} rows` : ""}
        />
      </div>

      {excel ? (
        <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h2 className="text-lg font-medium">Field mapping (Excel column to mailing field)</h2>
            <button
              className="rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-slate-950"
              onClick={() => setMapping(autoDetectMailingMapping(excel))}
            >
              Auto detect mapping
            </button>
          </div>
          <div className="mt-4 grid gap-2 sm:grid-cols-2">
            {excel.headers.map((header, index) => (
              <div key={index} className="flex items-center gap-2">
                <span className="w-1/2 truncate text-sm text-ink-2" title={header}>{header || `Column ${index + 1}`}</span>
                <select
                  className="w-1/2 rounded border border-line px-2 py-1 text-sm"
                  value={mapping[index] ?? ""}
                  onChange={(event) => setMapping({ ...mapping, [index]: event.target.value })}
                >
                  <option value="">Not mapped</option>
                  {MAILING_FIELDS.map((field) => (
                    <option key={field.key} value={field.key}>{field.label}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>
        </div>
      ) : null}

      {records && excel ? (
        <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h2 className="text-lg font-medium">Fields to verify</h2>
            <div className="flex items-center gap-2 text-sm">
              <label className="text-ink-2">Match key</label>
              <select className="rounded-lg border border-line px-3 py-1.5" value={matchKey} onChange={(event) => setMatchKey(event.target.value as MailingFieldKey)}>
                {MAILING_FIELDS.filter((f) => ["fullName", "shareholderNumber", "postalCode", "email", "cellPhone"].includes(f.key) && mappedKeys.has(f.key)).map((f) => (
                  <option key={f.key} value={f.key}>{f.label}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="mt-4 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
            {MAILING_FIELDS.filter((f) => ["fullName", "title", "addressLine1", "addressLine2", "addressLine3", "addressLine4", "postalCode", "email", "cellPhone"].includes(f.key)).map((field) => {
              const mapped = mappedKeys.has(field.key);
              return (
                <label key={field.key} className={`flex items-center gap-2 rounded-lg border px-3 py-2 text-sm ${mapped ? "border-line" : "border-slate-100 bg-surface text-ink-3"}`}>
                  <input type="checkbox" checked={fields.has(field.key)} disabled={!mapped} onChange={() => toggleField(field.key)} />
                  <span>{field.label}</span>
                  {!mapped ? <span className="ml-auto text-xs">not mapped</span> : null}
                </label>
              );
            })}
          </div>
          <button className="mt-5 rounded-md bg-ink px-6 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint" onClick={props.onRun} disabled={props.processing}>
            Run mailing verification
          </button>
        </div>
      ) : null}

      {result ? (
        <div className="mt-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <h2 className="text-2xl font-medium tracking-[-0.02em]">Results</h2>
            <button className="rounded-md bg-accent-tint px-5 py-2.5 text-sm font-medium text-ink transition hover:bg-accent-tint" onClick={props.onDownload}>
              Download Excel report (6 sheets)
            </button>
          </div>

          <div className="mt-4 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <DashCard label="PDF records" value={result.summary.pdfRecords} />
            <DashCard label="Match %" value={result.summary.matchPercentage} suffix="%" />
            <DashCard label="Exceptions" value={result.summary.exceptions} />
            <DashCard label="Missing" value={result.summary.missing} />
          </div>

          <div className="mt-4 grid gap-3 sm:grid-cols-4">
            <Legend color="bg-surface text-ink-2" label={`Exact ${result.summary.matches}`} />
            <Legend color="bg-accent-tint text-accent" label={`Partial ${result.summary.partial}`} />
            <Legend color="bg-surface text-error" label={`Mismatch ${result.summary.mismatches}`} />
            <Legend color="bg-surface text-ink-2" label={`Missing ${result.summary.missing}`} />
          </div>

          <div className="mt-6 rounded-xl border border-line bg-canvas">
            <div className="flex items-center justify-between border-b border-line px-5 py-3">
              <h3 className="text-lg font-medium">Detailed verification</h3>
              <label className="flex items-center gap-2 text-sm text-ink-2">
                <input type="checkbox" checked={showExceptionsOnly} onChange={(event) => setShowExceptionsOnly(event.target.checked)} />
                Exceptions only
              </label>
            </div>
            <div className="max-h-[30rem] overflow-auto">
              <table className="min-w-full text-left text-sm">
                <thead className="sticky top-0 bg-surface text-xs uppercase tracking-wide text-ink-3">
                  <tr>
                    <th className="px-4 py-2">Rec</th><th className="px-4 py-2">Excel Row</th><th className="px-4 py-2">Field</th>
                    <th className="px-4 py-2">PDF Value</th><th className="px-4 py-2">Excel Value</th><th className="px-4 py-2">Status</th><th className="px-4 py-2">Comment</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {flatRows.map((row, index) => (
                    <tr key={index}>
                      <td className="px-4 py-2 text-ink-2">{row.record}</td>
                      <td className="px-4 py-2 text-ink-2">{row.row}</td>
                      <td className="px-4 py-2 text-ink-2">{row.field}</td>
                      <td className="max-w-48 truncate px-4 py-2 text-ink-2">{row.pdf || "-"}</td>
                      <td className="max-w-48 truncate px-4 py-2 text-ink-2">{row.excel || "-"}</td>
                      <td className="px-4 py-2"><span className={`rounded-md px-2 py-1 text-xs font-medium ${STATUS_COLORS[row.status] ?? "bg-surface text-ink-2"}`}>{row.status}</span></td>
                      <td className="max-w-48 truncate px-4 py-2 text-ink-3">{row.comment}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {result.quality.length ? (
            <div className="mt-6 rounded-xl border border-line bg-canvas p-5">
              <h3 className="text-lg font-medium">Data quality issues ({result.quality.length})</h3>
              <div className="mt-3 max-h-64 overflow-auto">
                <table className="min-w-full text-left text-sm">
                  <thead className="text-xs uppercase tracking-wide text-ink-3"><tr><th className="py-2 pr-4">Location</th><th className="py-2">Issue</th></tr></thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.quality.slice(0, 200).map((issue, index) => (
                      <tr key={index}><td className="py-2 pr-4 text-ink-2">{issue.location}</td><td className="py-2 text-ink-2">{issue.issue}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : null}
        </div>
      ) : null}
    </section>
  );
}
