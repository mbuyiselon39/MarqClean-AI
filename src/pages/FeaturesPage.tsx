import React, { useState } from "react";
import { motion } from "framer-motion";
import {
  CheckCircle,
  Table,
  Cpu,
  Shield,
  Sparkles,
  ArrowRight,
  Lock,
  Layers,
  FileCheck,
  Server,
} from "lucide-react";
import ReconciliationHub from "../reconciliation/ReconciliationHub";
import DataToolbox from "../reconciliation/DataToolbox";

export interface FeaturesPageProps {
  initialSubTab?: string;
  onRequestAudit: () => void;
}

export const FeaturesPage: React.FC<FeaturesPageProps> = ({ initialSubTab = "recon-hub", onRequestAudit }) => {
  const [activeTab, setActiveTab] = useState<"recon-hub" | "data-toolbox" | "excel-automation" | "capabilities">(
    (initialSubTab as any) || "recon-hub"
  );

  return (
    <div className="min-h-screen bg-[#030611] text-slate-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-7xl">
        {/* Header */}
        <div className="flex flex-col gap-4 border-b border-slate-800 pb-8 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <Sparkles className="h-3.5 w-3.5" />
              Advanced Data Architecture &amp; Engine
            </div>
            <h1 className="mt-3 text-3xl font-extrabold text-white sm:text-4xl">
              Platform Features &amp; Reconciliation Suite
            </h1>
            <p className="mt-2 text-sm text-slate-400 max-w-2xl">
              High-throughput data comparison, multi-sheet workbook managers, custom DAX pattern generators, and zero-downtime client-side security.
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

        {/* Feature Navigation Tabs */}
        <div className="mt-8 flex flex-wrap gap-2 sm:gap-3">
          <button
            onClick={() => setActiveTab("recon-hub")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "recon-hub"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <CheckCircle className="h-4 w-4 text-cyan-400" />
            <span>Verification &amp; Reconciliation Hub</span>
          </button>

          <button
            onClick={() => setActiveTab("data-toolbox")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "data-toolbox"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Table className="h-4 w-4 text-blue-400" />
            <span>Free Data Toolbox</span>
          </button>

          <button
            onClick={() => setActiveTab("excel-automation")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "excel-automation"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Cpu className="h-4 w-4 text-emerald-400" />
            <span>Excel Automation Engine</span>
          </button>

          <button
            onClick={() => setActiveTab("capabilities")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "capabilities"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Shield className="h-4 w-4 text-purple-400" />
            <span>Platform Capabilities</span>
          </button>
        </div>

        {/* ------------------------------------------------------------- */}
        {/* FEATURE 1: Verification & Reconciliation Hub */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "recon-hub" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8">
            <div className="rounded-3xl border border-cyan-500/20 bg-slate-950/90 p-4 sm:p-6 backdrop-blur-xl">
              <ReconciliationHub onExit={() => setActiveTab("capabilities")} />
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* FEATURE 2: Free Data Toolbox */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "data-toolbox" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8">
            <div className="rounded-3xl border border-blue-500/20 bg-slate-950/90 p-4 sm:p-6 backdrop-blur-xl">
              <DataToolbox onExit={() => setActiveTab("capabilities")} />
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* FEATURE 3: Excel Automation Engine */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "excel-automation" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-6">
            <div className="rounded-3xl border border-emerald-500/20 bg-slate-950/80 p-6 sm:p-10 backdrop-blur-xl">
              <div className="flex items-center gap-3">
                <div className="rounded-2xl bg-emerald-500/10 p-3 text-emerald-400">
                  <Cpu className="h-8 w-8" />
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-white">Excel Automation &amp; DAX Model Engine</h2>
                  <p className="text-xs text-slate-400">
                    High-performance tabular formula synthesis, relationship integrity validation, and Power BI DAX patterns.
                  </p>
                </div>
              </div>

              <div className="mt-8 grid gap-6 sm:grid-cols-3">
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <h3 className="text-sm font-bold text-white">Dynamic Array Formulas</h3>
                  <p className="mt-2 text-xs text-slate-400 leading-relaxed">
                    Automated generation of spill-aware modern formulas: <code>FILTER</code>, <code>UNIQUE</code>, <code>SORTBY</code>, and <code>XLOOKUP</code> without manual cell coordinates.
                  </p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <h3 className="text-sm font-bold text-white">Power BI DAX Patterns</h3>
                  <p className="mt-2 text-xs text-slate-400 leading-relaxed">
                    Pre-compiled time intelligence measures, Year-to-Date (YTD), Compound Annual Growth Rate (CAGR), and cumulative balances.
                  </p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <h3 className="text-sm font-bold text-white">Macro-Free Workbooks</h3>
                  <p className="mt-2 text-xs text-slate-400 leading-relaxed">
                    Zero VBA or insecure macros required. Everything is exported as pure native OpenXML (.xlsx) files supported across Microsoft 365, Excel for Web, and LibreOffice.
                  </p>
                </div>
              </div>
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* FEATURE 4: Platform Capabilities */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "capabilities" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-8">
            <div className="rounded-3xl border border-purple-500/20 bg-slate-950/80 p-6 sm:p-10 backdrop-blur-xl">
              <h2 className="text-2xl font-bold text-white">Enterprise Platform Architecture &amp; Security</h2>
              <p className="mt-2 text-sm text-slate-400 max-w-3xl">
                MarqClean AI is built for maximum speed, strict enterprise data sovereignty, and zero-compromise confidentiality.
              </p>

              <div className="mt-8 grid gap-6 sm:grid-cols-2 lg:grid-cols-4">
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <Lock className="h-6 w-6 text-cyan-400" />
                  <h3 className="mt-3 text-sm font-bold text-white">100% Client-Side Privacy</h3>
                  <p className="mt-2 text-xs text-slate-400">
                    Files are processed in ephemeral browser memory. Financial statements and PII never transit to third-party servers.
                  </p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <Server className="h-6 w-6 text-emerald-400" />
                  <h3 className="mt-3 text-sm font-bold text-white">2GB Batch Streaming</h3>
                  <p className="mt-2 text-xs text-slate-400">
                    Chunked web-worker architecture handles large multi-sheet ledger workbooks without browser tab freezes.
                  </p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <FileCheck className="h-6 w-6 text-blue-400" />
                  <h3 className="mt-3 text-sm font-bold text-white">ISO 14644 Alignment</h3>
                  <p className="mt-2 text-xs text-slate-400">
                    Cleanroom data protocols mathematically verify zero stray characters, null references, or broken formulas.
                  </p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <Layers className="h-6 w-6 text-purple-400" />
                  <h3 className="mt-3 text-sm font-bold text-white">Multi-Format Interoperability</h3>
                  <p className="mt-2 text-xs text-slate-400">
                    Seamless conversion between CSV, TSV, XLSX, XLS, PDF bank tables, JSON, and QIF accounting formats.
                  </p>
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </div>
    </div>
  );
};

export default FeaturesPage;
