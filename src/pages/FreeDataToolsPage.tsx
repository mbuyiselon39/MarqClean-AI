import React, { useState } from "react";
import {
  Sparkles,
  ArrowRight,
  Wand2,
  Table,
  FileCode2,
  Hash,
  Shield,
} from "lucide-react";
import Papa from "papaparse";
import { CopyDataButton } from "../components/CopyDataButton";

interface FreeDataToolsPageProps {
  onRequestAudit: () => void;
}

export const FreeDataToolsPage: React.FC<FreeDataToolsPageProps> = ({ onRequestAudit }) => {
  // Tool 1: Text & Email Normalizer
  const [quickInput, setQuickInput] = useState(
    "john doe,  CEO\nALEX.SMITH@ACME.CORP\n  maria garcia  \nALEX.SMITH@ACME.CORP\nsupport@fintech-ledger.org"
  );
  const [quickMode, setQuickMode] = useState<"proper" | "upper" | "lower" | "dedupe" | "email-clean">("proper");
  const [quickOutput, setQuickOutput] = useState("");

  // Tool 2: CSV / JSON / TSV Converter
  const [converterInput, setConverterInput] = useState(
    "id,name,role,department,salary\n1,Sarah Connor,Analyst,Operations,85000\n2,John Matrix,Engineer,Security,110000\n3,Ellen Ripley,Lead,Logistics,125000"
  );
  const [converterTarget, setConverterTarget] = useState<"json" | "tsv" | "pipe" | "semicolon">("json");
  const [converterOutput, setConverterOutput] = useState("");

  // Tool 3: Header & Field Sanitizer
  const [headerInput, setHeaderInput] = useState(
    "Customer Full Name (Primary), Total Revenue USD $, Transaction # ID, Date / Time Created"
  );
  const [headerCase, setHeaderCase] = useState<"snake" | "camel" | "pascal" | "kebab" | "clean_title">("snake");
  const [headerOutput, setHeaderOutput] = useState("");

  // Tool 4: Financial Number & Accounting Formatter
  const [numberInput, setNumberInput] = useState(
    "$1,245.50\n(3,450.00)\n€ 4.500,75\n- $12,900.20\n(850.50)"
  );
  const [numberOutput, setNumberOutput] = useState("");

  // Process Tool 1
  const handleProcessQuickText = () => {
    if (!quickInput.trim()) return;
    if (quickMode === "proper") {
      setQuickOutput(
        quickInput.replace(/\b\w+/g, (txt) => txt.charAt(0).toUpperCase() + txt.substring(1).toLowerCase())
      );
    } else if (quickMode === "upper") {
      setQuickOutput(quickInput.toUpperCase());
    } else if (quickMode === "lower") {
      setQuickOutput(quickInput.toLowerCase());
    } else if (quickMode === "dedupe") {
      const lines = quickInput.split("\n");
      const unique = Array.from(new Set(lines.map((l) => l.trim()))).filter(Boolean);
      setQuickOutput(unique.join("\n"));
    } else if (quickMode === "email-clean") {
      const emails = quickInput.match(/[\w.-]+@[\w.-]+\.\w+/g) || [];
      const clean = Array.from(new Set(emails.map((e) => e.toLowerCase().trim())));
      setQuickOutput(clean.join("\n"));
    }
  };

  // Process Tool 2
  const handleProcessConverter = () => {
    if (!converterInput.trim()) return;
    try {
      const parsed = Papa.parse(converterInput.trim(), { header: true, skipEmptyLines: true });
      if (converterTarget === "json") {
        setConverterOutput(JSON.stringify(parsed.data, null, 2));
      } else if (converterTarget === "tsv") {
        setConverterOutput(Papa.unparse(parsed.data, { delimiter: "\t" }));
      } else if (converterTarget === "pipe") {
        setConverterOutput(Papa.unparse(parsed.data, { delimiter: "|" }));
      } else if (converterTarget === "semicolon") {
        setConverterOutput(Papa.unparse(parsed.data, { delimiter: ";" }));
      }
    } catch {
      setConverterOutput("Error parsing tabular input. Verify CSV structure.");
    }
  };

  // Process Tool 3
  const handleProcessHeaders = () => {
    if (!headerInput.trim()) return;
    const rawItems = headerInput.split(/[\n,]+/).map((s) => s.trim()).filter(Boolean);
    const cleaned = rawItems.map((item) => {
      // remove special chars
      const sanitized = item.replace(/[^a-zA-Z0-9\s]/g, " ").trim();
      const words = sanitized.split(/\s+/).filter(Boolean);

      if (headerCase === "snake") {
        return words.map((w) => w.toLowerCase()).join("_");
      }
      if (headerCase === "kebab") {
        return words.map((w) => w.toLowerCase()).join("-");
      }
      if (headerCase === "camel") {
        return words
          .map((w, i) =>
            i === 0 ? w.toLowerCase() : w.charAt(0).toUpperCase() + w.substring(1).toLowerCase()
          )
          .join("");
      }
      if (headerCase === "pascal") {
        return words
          .map((w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase())
          .join("");
      }
      return words
        .map((w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase())
        .join(" ");
    });
    setHeaderOutput(cleaned.join(", "));
  };

  // Process Tool 4
  const handleProcessNumbers = () => {
    if (!numberInput.trim()) return;
    const lines = numberInput.split("\n");
    const cleaned = lines.map((line) => {
      const trimmed = line.trim();
      if (!trimmed) return "";
      let isNegative = false;
      if (trimmed.startsWith("(") && trimmed.endsWith(")")) {
        isNegative = true;
      }
      if (trimmed.includes("-")) {
        isNegative = true;
      }

      // Check if European format 1.234,56
      let numStr = trimmed.replace(/[^0-9.,]/g, "");
      if (numStr.includes(",") && numStr.includes(".")) {
        if (numStr.indexOf(".") < numStr.indexOf(",")) {
          // 1.234,56 -> replace . with '' and , with .
          numStr = numStr.replace(/\./g, "").replace(",", ".");
        } else {
          // 1,234.56 -> replace , with ''
          numStr = numStr.replace(/,/g, "");
        }
      } else if (numStr.includes(",")) {
        numStr = numStr.replace(/,/g, ".");
      }

      const val = parseFloat(numStr);
      if (isNaN(val)) return trimmed;
      const finalVal = isNegative ? -Math.abs(val) : Math.abs(val);
      return finalVal.toFixed(2);
    });
    setNumberOutput(cleaned.join("\n"));
  };

  return (
    <div className="min-h-screen bg-[#030611] text-slate-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-7xl space-y-12">
        {/* Page Header */}
        <div className="flex flex-col gap-4 border-b border-slate-800 pb-8 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <Sparkles className="h-3.5 w-3.5" />
              100% Free Browser Data Cleanroom
            </div>
            <h1 className="mt-3 text-3xl font-extrabold text-white sm:text-4xl">
              Free Data Transformation Toolbox
            </h1>
            <p className="mt-2 text-sm text-slate-400 max-w-2xl">
              Instant browser-native tools to sanitize, transform, normalize, and format messy spreadsheets, columns, and data objects with zero server uploads.
            </p>
          </div>

          <button
            onClick={onRequestAudit}
            className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 px-5 py-2.5 text-xs font-bold uppercase text-slate-950 hover:brightness-110 transition shadow-lg shadow-cyan-500/20"
          >
            <span>Request Diagnostic Audit</span>
            <ArrowRight className="h-3.5 w-3.5" />
          </button>
        </div>

        {/* Security & Privacy Banner */}
        <div className="rounded-2xl border border-emerald-500/25 bg-emerald-950/20 p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="rounded-xl bg-emerald-500/20 p-2 text-emerald-400">
              <Shield className="h-5 w-5" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-emerald-300">Zero-Egress In-Memory Cleanroom</h3>
              <p className="text-xs text-slate-400">
                All data transformations are calculated 100% locally in your browser memory. No records leave your machine.
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs font-mono text-emerald-400 font-bold">
            <span className="h-2 w-2 rounded-full bg-emerald-400 animate-ping" />
            AIR-GAPPED COMPLIANT
          </div>
        </div>

        {/* 1. Instant Text & List Normalizer Tool */}
        <section className="rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl shadow-2xl">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
            <div>
              <div className="inline-flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-cyan-400">
                <Wand2 className="h-4 w-4" />
                Tool 1 &bull; Text &amp; Email Normalizer
              </div>
              <h2 className="mt-1 text-2xl font-bold text-white">
                Capitalization, Deduplication &amp; RFC Email Filter
              </h2>
            </div>

            {/* Mode selector */}
            <div className="flex flex-wrap gap-1.5 rounded-xl border border-slate-800 bg-slate-900 p-1">
              {[
                { id: "proper", label: "Proper Case" },
                { id: "upper", label: "UPPERCASE" },
                { id: "lower", label: "lowercase" },
                { id: "dedupe", label: "Deduplicate" },
                { id: "email-clean", label: "Extract Emails" },
              ].map((m) => (
                <button
                  key={m.id}
                  onClick={() => setQuickMode(m.id as any)}
                  className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${
                    quickMode === m.id
                      ? "bg-cyan-500 text-slate-950 font-bold"
                      : "text-slate-400 hover:text-white"
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>
          </div>

          <div className="mt-6 grid gap-6 sm:grid-cols-2">
            <div>
              <label className="block text-xs font-semibold text-slate-300">Input Raw Text / Column Lines</label>
              <textarea
                rows={7}
                placeholder="Paste disordered text, names, or raw email lines here..."
                value={quickInput}
                onChange={(e) => setQuickInput(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
              />
              <button
                onClick={handleProcessQuickText}
                className="mt-3 rounded-xl bg-cyan-500 px-5 py-2 text-xs font-bold text-slate-950 hover:bg-cyan-400 transition shadow-md shadow-cyan-500/20"
              >
                Execute Transformation
              </button>
            </div>

            <div>
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-300">Cleaned &amp; Normalized Output</label>
                {quickOutput && (
                  <CopyDataButton
                    payload={{ type: "text", text: quickOutput }}
                    buttonTheme="cyan"
                    size="xs"
                  />
                )}
              </div>
              <textarea
                rows={7}
                readOnly
                placeholder="Cleaned output will appear here..."
                value={quickOutput}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-950 p-4 font-mono text-xs text-cyan-300 placeholder-slate-600 focus:outline-none"
              />
            </div>
          </div>
        </section>

        {/* 2. CSV / JSON / TSV Structure Converter */}
        <section className="rounded-3xl border border-blue-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl shadow-2xl">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
            <div>
              <div className="inline-flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-blue-400">
                <Table className="h-4 w-4" />
                Tool 2 &bull; CSV &amp; JSON Format Bridge
              </div>
              <h2 className="mt-1 text-2xl font-bold text-white">
                Convert CSV &bull; JSON &bull; TSV &bull; Semicolon
              </h2>
            </div>

            <div className="flex flex-wrap gap-1.5 rounded-xl border border-slate-800 bg-slate-900 p-1">
              {[
                { id: "json", label: "To JSON Array" },
                { id: "tsv", label: "To Excel TSV (Tabs)" },
                { id: "pipe", label: "To Pipe (|)" },
                { id: "semicolon", label: "To Semicolon (;)" },
              ].map((m) => (
                <button
                  key={m.id}
                  onClick={() => setConverterTarget(m.id as any)}
                  className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${
                    converterTarget === m.id
                      ? "bg-blue-500 text-slate-950 font-bold"
                      : "text-slate-400 hover:text-white"
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>
          </div>

          <div className="mt-6 grid gap-6 sm:grid-cols-2">
            <div>
              <label className="block text-xs font-semibold text-slate-300">Input CSV or Tabular Data</label>
              <textarea
                rows={7}
                placeholder="Paste CSV data with headers..."
                value={converterInput}
                onChange={(e) => setConverterInput(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
              />
              <button
                onClick={handleProcessConverter}
                className="mt-3 rounded-xl bg-blue-500 px-5 py-2 text-xs font-bold text-slate-950 hover:bg-blue-400 transition shadow-md shadow-blue-500/20"
              >
                Convert Structure
              </button>
            </div>

            <div>
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-300">Converted Data Representation</label>
                {converterOutput && (
                  <CopyDataButton
                    payload={{ type: "text", text: converterOutput }}
                    buttonTheme="dark"
                    size="xs"
                  />
                )}
              </div>
              <textarea
                rows={7}
                readOnly
                placeholder="Structured output will appear here..."
                value={converterOutput}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-950 p-4 font-mono text-xs text-blue-300 placeholder-slate-600 focus:outline-none"
              />
            </div>
          </div>
        </section>

        {/* 3. Column Header Normalizer & DB Identifier Sanitizer */}
        <section className="rounded-3xl border border-emerald-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl shadow-2xl">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
            <div>
              <div className="inline-flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-emerald-400">
                <FileCode2 className="h-4 w-4" />
                Tool 3 &bull; Header &amp; Schema Identifier Sanitizer
              </div>
              <h2 className="mt-1 text-2xl font-bold text-white">
                Database &amp; Code-Ready Column Names
              </h2>
            </div>

            <div className="flex flex-wrap gap-1.5 rounded-xl border border-slate-800 bg-slate-900 p-1">
              {[
                { id: "snake", label: "snake_case" },
                { id: "camel", label: "camelCase" },
                { id: "pascal", label: "PascalCase" },
                { id: "kebab", label: "kebab-case" },
                { id: "clean_title", label: "Clean Title Case" },
              ].map((m) => (
                <button
                  key={m.id}
                  onClick={() => setHeaderCase(m.id as any)}
                  className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${
                    headerCase === m.id
                      ? "bg-emerald-500 text-slate-950 font-bold"
                      : "text-slate-400 hover:text-white"
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>
          </div>

          <div className="mt-6 grid gap-6 sm:grid-cols-2">
            <div>
              <label className="block text-xs font-semibold text-slate-300">Raw Headers (Comma or Line Separated)</label>
              <textarea
                rows={5}
                placeholder="e.g. Total Revenue USD $, Customer First & Last Name..."
                value={headerInput}
                onChange={(e) => setHeaderInput(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-white placeholder-slate-500 focus:border-emerald-500 focus:outline-none"
              />
              <button
                onClick={handleProcessHeaders}
                className="mt-3 rounded-xl bg-emerald-500 px-5 py-2 text-xs font-bold text-slate-950 hover:bg-emerald-400 transition shadow-md shadow-emerald-500/20"
              >
                Sanitize Identifiers
              </button>
            </div>

            <div>
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-300">SQL &amp; Python Compliant Output</label>
                {headerOutput && (
                  <CopyDataButton
                    payload={{ type: "text", text: headerOutput }}
                    buttonTheme="cyan"
                    size="xs"
                  />
                )}
              </div>
              <textarea
                rows={5}
                readOnly
                placeholder="Sanitized headers will appear here..."
                value={headerOutput}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-950 p-4 font-mono text-xs text-emerald-300 placeholder-slate-600 focus:outline-none"
              />
            </div>
          </div>
        </section>

        {/* 4. Financial Number & Accounting Formatter */}
        <section className="rounded-3xl border border-amber-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl shadow-2xl">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
            <div>
              <div className="inline-flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-amber-400">
                <Hash className="h-4 w-4" />
                Tool 4 &bull; Financial Number &amp; Accounting Cleaner
              </div>
              <h2 className="mt-1 text-2xl font-bold text-white">
                Normalize Accounting Brackets, Currencies &amp; Decimals
              </h2>
            </div>
          </div>

          <div className="mt-6 grid gap-6 sm:grid-cols-2">
            <div>
              <label className="block text-xs font-semibold text-slate-300">Messy Financial Numbers / Brackets</label>
              <textarea
                rows={6}
                placeholder="e.g. (1,234.50), $4.500,20..."
                value={numberInput}
                onChange={(e) => setNumberInput(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-white placeholder-slate-500 focus:border-amber-500 focus:outline-none"
              />
              <button
                onClick={handleProcessNumbers}
                className="mt-3 rounded-xl bg-amber-500 px-5 py-2 text-xs font-bold text-slate-950 hover:bg-amber-400 transition shadow-md shadow-amber-500/20"
              >
                Standardize Floats
              </button>
            </div>

            <div>
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-300">Clean Decimal Floats (Machine-Readable)</label>
                {numberOutput && (
                  <CopyDataButton
                    payload={{ type: "text", text: numberOutput }}
                    buttonTheme="dark"
                    size="xs"
                  />
                )}
              </div>
              <textarea
                rows={6}
                readOnly
                placeholder="Normalized decimal numbers will appear here..."
                value={numberOutput}
                className="mt-2 w-full rounded-2xl border border-slate-800 bg-slate-950 p-4 font-mono text-xs text-amber-300 placeholder-slate-600 focus:outline-none"
              />
            </div>
          </div>
        </section>
      </div>
    </div>
  );
};

export default FreeDataToolsPage;
