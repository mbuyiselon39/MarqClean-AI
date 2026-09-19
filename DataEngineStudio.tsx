import { useMemo, useState } from "react";
import Papa from "papaparse";
import { WorkspaceShell } from "./WorkspaceShell";
import { auditPhone, extractWebSignals, validateLeadRow } from "./localDataEngines";
import { queryCsv } from "./duckdbEngine";

type EngineTab = "engine" | "validate" | "phone" | "enrich";

const SAMPLE_SQL = `SELECT COUNT(*) AS records,
       COUNT(DISTINCT lower(trim(COALESCE("Email", '')))) AS unique_emails
FROM read_csv_auto(__CSV__, HEADER=TRUE, SAMPLE_SIZE=-1)`;

function download(name: string, body: string, type = "text/csv;charset=utf-8") {
  const url = URL.createObjectURL(new Blob([body], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

export default function DataEngineStudio({ onExit }: { onExit: () => void }) {
  const [tab, setTab] = useState<EngineTab>("engine");
  const [file, setFile] = useState<File | null>(null);
  const [sql, setSql] = useState(SAMPLE_SQL);
  const [queryRows, setQueryRows] = useState<Record<string, unknown>[]>([]);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("Ready — all processing stays in this browser.");
  const [phoneValue, setPhoneValue] = useState("");
  const [phoneCountry, setPhoneCountry] = useState("ZA");
  const [phoneResult, setPhoneResult] = useState<ReturnType<typeof auditPhone> | null>(null);
  const [schemaIssues, setSchemaIssues] = useState<string[]>([]);
  const [htmlInput, setHtmlInput] = useState("");
  const [urlInput, setUrlInput] = useState("");
  const [enrichment, setEnrichment] = useState<ReturnType<typeof extractWebSignals> | null>(null);

  const tabs = useMemo(() => [
    { key: "engine" as const, label: "SQL Data Engine", blurb: "DuckDB-WASM joins, grouping and deduplication" },
    { key: "validate" as const, label: "Schema Guard", blurb: "Zod-powered row validation and error reports" },
    { key: "phone" as const, label: "Phone Intelligence", blurb: "Country detection and international formatting" },
    { key: "enrich" as const, label: "Web Enrichment", blurb: "Extract signals from permitted local HTML" },
  ], []);

  async function runQuery() {
    if (!file) { setMessage("Choose a CSV first."); return; }
    setBusy(true);
    setMessage("Running the query in DuckDB-WASM…");
    try {
      const rows = await queryCsv(file, sql);
      setQueryRows(rows.slice(0, 100));
      setMessage(`Query complete. Showing ${Math.min(rows.length, 100)} result rows.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "DuckDB query failed.");
    } finally { setBusy(false); }
  }

  function validateCsv() {
    if (!file) { setMessage("Choose a CSV first."); return; }
    setBusy(true);
    setSchemaIssues([]);
    Papa.parse<Record<string, unknown>>(file, {
      header: true,
      worker: true,
      skipEmptyLines: "greedy",
      chunkSize: 1024 * 1024 * 4,
      chunk(results) {
        const rows = results.data as Record<string, unknown>[];
        const issues = rows.flatMap((row, index) => {
          const result = validateLeadRow(row);
          return result.valid ? [] : result.errors.map((error) => `Row ${index + 1}: ${error}`);
        });
        setSchemaIssues((current) => [...current, ...issues].slice(0, 500));
      },
      complete() {
        setBusy(false);
        setMessage("Validation complete. Invalid rows are listed below.");
      },
      error(error) {
        setBusy(false);
        setMessage(error.message);
      },
    });
  }

  function auditCurrentPhone() {
    setPhoneResult(auditPhone(phoneValue, phoneCountry as never));
  }

  async function enrichUrl() {
    const url = urlInput.trim();
    if (!url) { setMessage("Enter a URL first."); return; }
    setBusy(true);
    setMessage("Attempting a direct browser fetch. The target site must allow CORS.");
    try {
      const response = await fetch(url, { headers: { Accept: "text/html,application/xhtml+xml" } });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      setEnrichment(extractWebSignals(await response.text(), url));
      setMessage("Web signals extracted locally. No API or proxy was used.");
    } catch (error) {
      setMessage("Direct fetch was blocked or unavailable. Paste saved HTML below instead — this keeps MarqClean 100% serverless.");
    } finally { setBusy(false); }
  }

  function parseHtml() {
    if (!htmlInput.trim()) { setMessage("Paste HTML first."); return; }
    setEnrichment(extractWebSignals(htmlInput, urlInput.trim()));
    setMessage("HTML parsed locally.");
  }

  return (
    <WorkspaceShell
      icon="⚡"
      title="MarqClean Local Data Engine"
      description="A browser-native performance layer for large CSV workflows, validation, phone intelligence and permitted web enrichment. No API keys, server uploads or scraping proxy are required."
      engine="DuckDB-WASM + Papa Parse Workers + Zod + libphonenumber-js"
      formats={["CSV", "WASM", "LOCAL"]}
      status={{ tone: busy ? "processing" : "success", label: busy ? "Processing locally" : "Browser ready" }}
    >
      <div className="space-y-6">
        <div className="ws-surface rounded-2xl p-4">
          <div className="flex flex-wrap gap-2">
            {tabs.map((item) => (
              <button key={item.key} type="button" onClick={() => setTab(item.key)}
                className={`rounded-full border px-4 py-2 text-sm font-semibold transition ${tab === item.key ? "border-[#0877e8] bg-[#0877e8] text-white" : "border-slate-200 bg-white text-slate-600 hover:border-[#0877e8]"}`}>
                {item.label}
              </button>
            ))}
          </div>
        </div>

        {tab !== "enrich" ? (
          <div className="ws-surface rounded-2xl p-6">
            <label className="block text-sm font-semibold text-slate-800">Local CSV workspace</label>
            <p className="mt-1 text-sm text-slate-500">Files are read from your device. For multi-hundred-MB CSVs, the worker and DuckDB paths avoid pushing every row through React state.</p>
            <input className="mt-4 block w-full rounded-xl border p-3" type="file" accept=".csv,text/csv" onChange={(event) => setFile(event.target.files?.[0] ?? null)} />
            {file ? <p className="mt-2 text-xs font-medium text-slate-500">{file.name} · {(file.size / 1024 / 1024).toFixed(1)} MB</p> : null}
          </div>
        ) : null}

        {tab === "engine" ? (
          <div className="ws-surface rounded-2xl p-6">
            <div className="flex flex-col gap-3 lg:flex-row lg:items-end">
              <label className="flex-1 text-sm font-semibold text-slate-700">SQL</label>
              <button className="ws-btn-primary rounded-full px-5 py-2 text-sm" onClick={runQuery} disabled={busy}>Run locally</button>
            </div>
            <textarea className="mt-3 min-h-36 w-full rounded-xl border p-4 font-mono text-xs" value={sql} onChange={(e) => setSql(e.target.value)} />
            <p className="mt-3 text-xs text-slate-500">Use <code>__CSV__</code> as the registered local file. Example: GROUP BY, JOIN, DISTINCT, window functions and other DuckDB SQL can run without a server.</p>
            {queryRows.length ? <pre className="mt-4 max-h-72 overflow-auto rounded-xl bg-slate-950 p-4 text-xs text-slate-100">{JSON.stringify(queryRows, null, 2)}</pre> : null}
          </div>
        ) : null}

        {tab === "validate" ? (
          <div className="ws-surface rounded-2xl p-6">
            <div className="flex items-center justify-between gap-3">
              <div><h2 className="text-lg font-bold">Lead data schema</h2><p className="text-sm text-slate-500">Email, phone, company and website fields are checked locally with Zod.</p></div>
              <button className="ws-btn-primary rounded-full px-5 py-2 text-sm" onClick={validateCsv} disabled={busy}>Validate file</button>
            </div>
            {schemaIssues.length ? (
              <div className="mt-5">
                <p className="text-sm font-semibold text-red-600">{schemaIssues.length} issues captured</p>
                <pre className="mt-2 max-h-64 overflow-auto rounded-xl bg-red-50 p-4 text-xs text-red-800">{schemaIssues.join("\n")}</pre>
                <button className="mt-3 ws-btn-secondary rounded-full px-4 py-2 text-sm" onClick={() => download("MarqClean_Error_Report.csv", Papa.unparse(schemaIssues.map((issue) => ({ error: issue }))))}>Download Error Report</button>
              </div>
            ) : null}
          </div>
        ) : null}

        {tab === "phone" ? (
          <div className="ws-surface rounded-2xl p-6">
            <h2 className="text-lg font-bold">Phone intelligence</h2>
            <p className="mt-1 text-sm text-slate-500">Detect country, validate the number and output international/national formats without a network call.</p>
            <div className="mt-5 grid gap-3 sm:grid-cols-[1fr_160px_auto]">
              <input className="rounded-xl border p-3" placeholder="+27 11 123 4567" value={phoneValue} onChange={(e) => setPhoneValue(e.target.value)} />
              <select className="rounded-xl border p-3" value={phoneCountry} onChange={(e) => setPhoneCountry(e.target.value)}><option value="ZA">South Africa</option><option value="US">United States</option><option value="GB">United Kingdom</option><option value="AU">Australia</option><option value="IN">India</option></select>
              <button className="ws-btn-primary rounded-full px-5 py-3 text-sm" onClick={auditCurrentPhone}>Audit</button>
            </div>
            {phoneResult ? <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              {[
                ["Valid", phoneResult.valid ? "Yes" : "No"],
                ["Country", phoneResult.country ?? "—"],
                ["International", phoneResult.international ?? "—"],
                ["National", phoneResult.national ?? "—"],
              ].map(([label, value]) => <div key={label} className="rounded-xl border bg-slate-50 p-4"><p className="text-xs text-slate-500">{label}</p><p className="mt-1 font-semibold text-slate-900">{value}</p></div>)}
            </div> : null}
          </div>
        ) : null}

        {tab === "enrich" ? (
          <div className="grid gap-6 lg:grid-cols-2">
            <div className="ws-surface rounded-2xl p-6">
              <h2 className="text-lg font-bold">Web enrichment — no API</h2>
              <p className="mt-1 text-sm text-slate-500">Direct fetch works only when the website permits browser CORS. Otherwise paste saved page HTML. This extracts title, company/schema.org, description, emails, phones, social links and useful links locally.</p>
              <input className="mt-5 w-full rounded-xl border p-3" placeholder="https://example.com" value={urlInput} onChange={(e) => setUrlInput(e.target.value)} />
              <div className="mt-3 flex flex-wrap gap-2">
                <button className="ws-btn-primary rounded-full px-5 py-2 text-sm" onClick={enrichUrl} disabled={busy}>Try direct fetch</button>
                <label className="ws-btn-secondary cursor-pointer rounded-full px-5 py-2 text-sm">Load HTML file<input type="file" accept=".html,.htm,text/html" className="hidden" onChange={async (e) => { const f=e.target.files?.[0]; if(f){setHtmlInput(await f.text());setMessage("HTML file loaded locally.");} }} /></label>
              </div>
              <textarea className="mt-4 min-h-48 w-full rounded-xl border p-3 font-mono text-xs" placeholder="Or paste saved HTML here…" value={htmlInput} onChange={(e) => setHtmlInput(e.target.value)} />
              <button className="ws-btn-secondary rounded-full px-5 py-2 text-sm" onClick={parseHtml}>Extract signals</button>
            </div>
            <div className="ws-surface rounded-2xl p-6">
              <h2 className="text-lg font-bold">Enrichment preview</h2>
              {enrichment ? <div className="mt-4 space-y-3 text-sm">
                {[
                  ["Company", enrichment.company || "Not detected"],
                  ["Title", enrichment.title || "Not detected"],
                  ["Industry", enrichment.industry || "Not detected"],
                  ["Emails", enrichment.emails.join(", ") || "None detected"],
                  ["Phones", enrichment.phones.join(", ") || "None detected"],
                  ["Social", enrichment.social.length ? enrichment.social.join("\n") : "None detected"],
                ].map(([label, value]) => <div key={label} className="rounded-xl border bg-slate-50 p-4"><p className="text-xs font-semibold text-slate-500">{label}</p><p className="mt-1 whitespace-pre-wrap break-words text-slate-800">{value}</p></div>)}
              </div> : <div className="mt-6 rounded-xl border border-dashed p-8 text-center text-sm text-slate-500">No enrichment result yet.</div>}
            </div>
          </div>
        ) : null}

        <div className="ws-surface rounded-2xl p-6">
          <div className="grid gap-4 md:grid-cols-3">
            <div><p className="text-xs font-semibold uppercase tracking-wider text-[#0877e8]">Performance</p><h3 className="mt-1 font-bold">Worker-first parsing</h3><p className="mt-1 text-sm text-slate-500">Papa Parse supports worker and streaming modes for very large CSVs.</p></div>
            <div><p className="text-xs font-semibold uppercase tracking-wider text-[#0877e8]">Analytics</p><h3 className="mt-1 font-bold">SQL in the browser</h3><p className="mt-1 text-sm text-slate-500">DuckDB-WASM provides joins, dedupe, grouping and aggregations without uploading rows.</p></div>
            <div><p className="text-xs font-semibold uppercase tracking-wider text-[#0877e8]">Enrichment</p><h3 className="mt-1 font-bold">CORS-aware by design</h3><p className="mt-1 text-sm text-slate-500">No scraping API is hidden behind the UI; blocked sites can be handled from saved HTML.</p></div>
          </div>
          <p className="mt-5 rounded-xl bg-slate-50 p-4 text-xs text-slate-500">{message}</p>
        </div>

        <button type="button" className="ws-btn-secondary rounded-full px-5 py-2 text-sm" onClick={onExit}>← Back to MarqClean</button>
      </div>
    </WorkspaceShell>
  );
}
