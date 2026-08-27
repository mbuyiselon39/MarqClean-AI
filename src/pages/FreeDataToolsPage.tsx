import React, { useState } from "react";
import {
  Sparkles,
  Copy,
  Check,
  ArrowRight,
  Wand2,
} from "lucide-react";
import { LiveMarketHub } from "../components/LiveMarketHub";

interface FreeDataToolsPageProps {
  onRequestAudit: () => void;
}

export const FreeDataToolsPage: React.FC<FreeDataToolsPageProps> = ({ onRequestAudit }) => {
  // Quick text tool state
  const [quickInput, setQuickInput] = useState("");
  const [quickMode, setQuickMode] = useState<"proper" | "upper" | "lower" | "dedupe" | "email-clean">("proper");
  const [quickOutput, setQuickOutput] = useState("");
  const [copied, setCopied] = useState(false);

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

  const handleCopy = () => {
    navigator.clipboard.writeText(quickOutput);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="min-h-screen bg-[#030611] text-slate-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-7xl space-y-12">
        {/* Page Header */}
        <div className="flex flex-col gap-4 border-b border-slate-800 pb-8 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <Sparkles className="h-3.5 w-3.5" />
              100% Free Public Utilities &amp; Live API Feeds
            </div>
            <h1 className="mt-3 text-3xl font-extrabold text-white sm:text-4xl">
              Free Data Toolbox &amp; Live Global Connectors
            </h1>
            <p className="mt-2 text-sm text-slate-400 max-w-2xl">
              Access real-time financial market streams, download curated global research datasets, or normalize messy text columns instantly in your browser with zero registration.
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

        {/* 1. Integrated Live Global Market & Research Hub */}
        <LiveMarketHub />

        {/* 2. Instant Text & List Normalizer Tool */}
        <section className="rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-6 sm:p-10 backdrop-blur-xl">
          <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
            <div>
              <div className="inline-flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-cyan-400">
                <Wand2 className="h-4 w-4" />
                Instant Browser Text &amp; List Transformer
              </div>
              <h2 className="mt-1 text-2xl font-bold text-white">
                Capitalization, Deduplication &amp; Email Extractor
              </h2>
            </div>

            {/* Mode selector */}
            <div className="flex flex-wrap gap-1.5 rounded-xl border border-slate-800 bg-slate-900 p-1">
              {[
                { id: "proper", label: "Proper Case" },
                { id: "upper", label: "UPPERCASE" },
                { id: "lower", label: "lowercase" },
                { id: "dedupe", label: "Deduplicate Lines" },
                { id: "email-clean", label: "Extract Clean Emails" },
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
                  <button
                    onClick={handleCopy}
                    className="flex items-center gap-1 text-xs text-cyan-400 hover:text-cyan-300 transition"
                  >
                    {copied ? <Check className="h-3.5 w-3.5 text-emerald-400" /> : <Copy className="h-3.5 w-3.5" />}
                    <span>{copied ? "Copied!" : "Copy to Clipboard"}</span>
                  </button>
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
      </div>
    </div>
  );
};

export default FreeDataToolsPage;
