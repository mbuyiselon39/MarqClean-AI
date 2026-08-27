import React, { useState } from "react";
import { motion } from "framer-motion";
import {
  Sparkles,
  ArrowRight,
  ShieldCheck,
  Cpu,
  CheckCircle2,
  Lock,
  Activity,
  Zap,
} from "lucide-react";
import { LiveMarketHub } from "../components/LiveMarketHub";
import { NavPageKey } from "../components/Navbar";

interface HomePageProps {
  onNavigate: (page: NavPageKey, subTab?: string) => void;
  onRequestAudit: (email?: string) => void;
}

export const HomePage: React.FC<HomePageProps> = ({ onNavigate, onRequestAudit }) => {
  const [quickAuditEmail, setQuickAuditEmail] = useState("");
  const [activePipelineTab, setActivePipelineTab] = useState<"raw" | "ai-clean" | "reconciled">("ai-clean");

  const handleQuickAuditSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (quickAuditEmail.trim()) {
      onRequestAudit(quickAuditEmail.trim());
    } else {
      onRequestAudit();
    }
  };

  return (
    <div className="relative min-h-screen bg-[#030611] text-slate-100 selection:bg-cyan-500/30 selection:text-cyan-200">
      {/* ----------------------------------------------------------------------- */}
      {/* 1. HERO SECTION (Follows DESIGN UI layout precisely) */}
      {/* ----------------------------------------------------------------------- */}
      <section className="relative overflow-hidden pt-12 pb-20 lg:pt-20 lg:pb-28">
        {/* Background Ambient Glows & Grid Mesh */}
        <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(ellipse_80%_60%_at_50%_-10%,rgba(0,210,255,0.18),rgba(0,0,0,0))]" />
        <div className="pointer-events-none absolute -left-40 top-1/4 h-96 w-96 rounded-full bg-cyan-600/10 blur-[120px]" />
        <div className="pointer-events-none absolute -right-40 top-1/3 h-96 w-96 rounded-full bg-blue-600/10 blur-[120px]" />
        
        <div className="relative mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <div className="mx-auto max-w-4xl text-center">
            {/* Pill Eyebrow Badge */}
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.4 }}
              className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-4 py-1.5 text-xs font-bold uppercase tracking-widest text-cyan-300 shadow-[0_0_16px_rgba(0,210,255,0.2)]"
            >
              <Sparkles className="h-3.5 w-3.5 text-cyan-400 animate-pulse" />
              <span>PRECISION ENGINEERED DATA AI</span>
            </motion.div>

            {/* Main Headline H1 */}
            <motion.h1
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.1 }}
              className="mt-6 text-4xl font-extrabold tracking-tight text-white sm:text-5xl lg:text-6xl"
            >
              AI-Driven Precision Data Cleaning for{" "}
              <span className="bg-gradient-to-r from-cyan-300 via-cyan-400 to-blue-500 bg-clip-text text-transparent drop-shadow-[0_0_24px_rgba(0,210,255,0.35)]">
                Critical Environments.
              </span>
            </motion.h1>

            {/* Centered, carefully balanced descriptive paragraph */}
            <motion.p
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.2 }}
              className="mx-auto mt-6 max-w-3xl text-base sm:text-lg leading-relaxed text-slate-300 text-center"
            >
              Maintaining flawless operational conditions for enterprise data pipelines, financial reconciliation, and high-volume spreadsheets through intelligent diagnostics and zero-tolerance sanitation standards.
            </motion.p>

            {/* CTAs */}
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.3 }}
              className="mt-8 flex flex-wrap items-center justify-center gap-4"
            >
              <button
                onClick={() => onRequestAudit()}
                className="flex items-center gap-2 rounded-2xl bg-gradient-to-r from-cyan-400 via-cyan-500 to-blue-600 px-7 py-3.5 text-xs font-black uppercase tracking-wider text-slate-950 transition hover:brightness-110 shadow-xl shadow-cyan-500/30 active:scale-95"
              >
                <span>REQUEST A FREE AUDIT</span>
                <ArrowRight className="h-4 w-4" />
              </button>

              <button
                onClick={() => onNavigate("products")}
                className="flex items-center gap-2 rounded-2xl border border-cyan-500/30 bg-slate-900/80 px-6 py-3.5 text-xs font-bold uppercase tracking-wider text-cyan-200 backdrop-blur-md transition hover:border-cyan-400 hover:bg-cyan-950/40 hover:text-white"
              >
                <span>EXPLORE PRODUCTS</span>
              </button>
            </motion.div>
          </div>

          {/* Interactive AI Data Pipeline Showcase (Replaces old hero clutter) */}
          <motion.div
            initial={{ opacity: 0, y: 25 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.4 }}
            className="mt-14 overflow-hidden rounded-3xl border border-cyan-500/25 bg-[#060b18]/90 p-4 sm:p-6 shadow-2xl shadow-cyan-950/80 backdrop-blur-2xl"
          >
            {/* Visualizer header & control switches */}
            <div className="flex flex-col gap-3 border-b border-slate-800 pb-4 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex items-center gap-3">
                <div className="flex h-3 w-3 items-center justify-center rounded-full bg-cyan-400 animate-ping" />
                <span className="text-xs font-bold uppercase tracking-wider text-cyan-300">
                  Live Stream: AI Sanitation &amp; Variance Engine
                </span>
              </div>

              <div className="flex rounded-xl border border-slate-800 bg-slate-900/90 p-1 text-xs">
                <button
                  onClick={() => setActivePipelineTab("raw")}
                  className={`rounded-lg px-3 py-1.5 font-semibold transition ${
                    activePipelineTab === "raw" ? "bg-rose-500/20 text-rose-300 border border-rose-500/40" : "text-slate-400 hover:text-white"
                  }`}
                >
                  Raw Disordered Feed
                </button>
                <button
                  onClick={() => setActivePipelineTab("ai-clean")}
                  className={`rounded-lg px-3 py-1.5 font-semibold transition ${
                    activePipelineTab === "ai-clean" ? "bg-cyan-500 text-slate-950 font-bold shadow-md shadow-cyan-500/30" : "text-slate-400 hover:text-white"
                  }`}
                >
                  AI Sanitized &amp; Formatted
                </button>
                <button
                  onClick={() => setActivePipelineTab("reconciled")}
                  className={`rounded-lg px-3 py-1.5 font-semibold transition ${
                    activePipelineTab === "reconciled" ? "bg-emerald-500/20 text-emerald-300 border border-emerald-500/40" : "text-slate-400 hover:text-white"
                  }`}
                >
                  Zero-Variance Ledger
                </button>
              </div>
            </div>

            {/* Pipeline Preview Data Grid */}
            <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-800/80 bg-slate-950/70 p-3 font-mono text-xs">
              {activePipelineTab === "raw" && (
                <div className="space-y-2 text-slate-400">
                  <div className="rounded-lg bg-rose-950/20 p-2.5 border border-rose-900/30 text-rose-300 flex items-center justify-between">
                    <span>⚠ 4 Anomalies Detected: Inconsistent delimiters, mixed casing, casing corruptions, invalid emails.</span>
                    <span className="text-[10px] uppercase font-bold text-rose-400">Status: Unsanitized</span>
                  </div>
                  <table className="w-full text-left">
                    <thead className="border-b border-slate-800 text-slate-500 uppercase text-[10px]">
                      <tr>
                        <th className="p-2">Raw Entity</th>
                        <th className="p-2">Uncleaned Contact</th>
                        <th className="p-2">Location / Currency</th>
                        <th className="p-2 text-right">Ledger Amount</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/40 text-slate-300">
                      <tr>
                        <td className="p-2 text-rose-300 line-through">apex CAPITAl holdings,,</td>
                        <td className="p-2 text-rose-300">alex..smith@APEX.COM </td>
                        <td className="p-2 text-slate-400">johannesburg / zar</td>
                        <td className="p-2 text-right text-rose-400">R 450,200.00??</td>
                      </tr>
                      <tr>
                        <td className="p-2 text-rose-300 line-through">bio_gene LABS llc</td>
                        <td className="p-2 text-rose-300">DR.SARAH@BIOGEN.ORG</td>
                        <td className="p-2 text-slate-400">cape town; ZAF</td>
                        <td className="p-2 text-right text-rose-400">120450.50</td>
                      </tr>
                      <tr>
                        <td className="p-2 text-rose-300 line-through">nova logistics (pty) ltd</td>
                        <td className="p-2 text-rose-300">OPS@NOVALOG.CO.ZA</td>
                        <td className="p-2 text-slate-400">london / gbp</td>
                        <td className="p-2 text-right text-rose-400">£84,900.00</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}

              {activePipelineTab === "ai-clean" && (
                <div className="space-y-2">
                  <div className="rounded-lg bg-cyan-950/30 p-2.5 border border-cyan-500/30 text-cyan-200 flex items-center justify-between">
                    <span className="flex items-center gap-2">
                      <CheckCircle2 className="h-4 w-4 text-cyan-400" />
                      100% Sanitized: Proper Case applied, ISO country codes standardized, clean RFC emails.
                    </span>
                    <span className="text-[10px] uppercase font-bold text-cyan-400">Status: Cleanroom Ready</span>
                  </div>
                  <table className="w-full text-left">
                    <thead className="border-b border-slate-800 text-cyan-400 uppercase text-[10px]">
                      <tr>
                        <th className="p-2">Standardized Entity</th>
                        <th className="p-2">Verified Email</th>
                        <th className="p-2">Jurisdiction</th>
                        <th className="p-2 text-right">Sanitized Value</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/40 text-slate-200 font-medium">
                      <tr className="hover:bg-cyan-500/5">
                        <td className="p-2 text-white font-semibold">Apex Capital Holdings (Pty) Ltd</td>
                        <td className="p-2 text-cyan-300">alex.smith@apex.com</td>
                        <td className="p-2 text-slate-300">Johannesburg (ZAF)</td>
                        <td className="p-2 text-right text-emerald-400 font-bold">R 450,200.00</td>
                      </tr>
                      <tr className="hover:bg-cyan-500/5">
                        <td className="p-2 text-white font-semibold">BioGene Laboratories LLC</td>
                        <td className="p-2 text-cyan-300">sarah@biogen.org</td>
                        <td className="p-2 text-slate-300">Cape Town (ZAF)</td>
                        <td className="p-2 text-right text-emerald-400 font-bold">R 120,450.50</td>
                      </tr>
                      <tr className="hover:bg-cyan-500/5">
                        <td className="p-2 text-white font-semibold">Nova Logistics Ltd</td>
                        <td className="p-2 text-cyan-300">ops@novalog.co.za</td>
                        <td className="p-2 text-slate-300">London (GBR)</td>
                        <td className="p-2 text-right text-emerald-400 font-bold">£ 84,900.00</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}

              {activePipelineTab === "reconciled" && (
                <div className="space-y-2">
                  <div className="rounded-lg bg-emerald-950/30 p-2.5 border border-emerald-500/30 text-emerald-200 flex items-center justify-between">
                    <span className="flex items-center gap-2">
                      <CheckCircle2 className="h-4 w-4 text-emerald-400" />
                      Two-Way Auto-Reconciliation Complete: Bank statement matches internal ledger with 0.00 variance.
                    </span>
                    <span className="text-[10px] uppercase font-bold text-emerald-400">Variance: 0.00 (Match)</span>
                  </div>
                  <table className="w-full text-left">
                    <thead className="border-b border-slate-800 text-emerald-400 uppercase text-[10px]">
                      <tr>
                        <th className="p-2">Transaction ID</th>
                        <th className="p-2">Bank Statement Reference</th>
                        <th className="p-2">Internal Ledger</th>
                        <th className="p-2 text-right">Variance</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/40 text-slate-200 font-medium">
                      <tr>
                        <td className="p-2 font-mono text-cyan-300">TX-2024-901</td>
                        <td className="p-2">APEX CAP - SETTLEMENT #109</td>
                        <td className="p-2 text-slate-300">Apex Capital Holdings (Pty) Ltd</td>
                        <td className="p-2 text-right text-emerald-400 font-bold">0.00 (Exact)</td>
                      </tr>
                      <tr>
                        <td className="p-2 font-mono text-cyan-300">TX-2024-902</td>
                        <td className="p-2">BIOGENE LABS - TRUST ACC</td>
                        <td className="p-2 text-slate-300">BioGene Laboratories LLC</td>
                        <td className="p-2 text-right text-emerald-400 font-bold">0.00 (Exact)</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* Bottom action trigger inside preview */}
            <div className="mt-4 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-400">
              <span className="flex items-center gap-1.5 text-cyan-300">
                <Lock className="h-3.5 w-3.5 text-emerald-400" />
                Processed in-memory via high-speed client WebAssembly / WebWorkers.
              </span>
              <button
                onClick={() => onNavigate("products", "cleaner")}
                className="flex items-center gap-1 font-bold text-cyan-400 hover:text-cyan-200 transition"
              >
                <span>Launch Quick Data &amp; CSV Cleaner</span>
                <ArrowRight className="h-3 w-3" />
              </button>
            </div>
          </motion.div>
        </div>
      </section>

      {/* ----------------------------------------------------------------------- */}
      {/* 2. CORE METHODOLOGIES (Matches DESIGN UI layout) */}
      {/* ----------------------------------------------------------------------- */}
      <section className="relative border-t border-slate-800/80 bg-[#040816] py-20 lg:py-24">
        <div className="mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <div className="mx-auto max-w-3xl text-center">
            <h2 className="text-3xl font-extrabold tracking-tight text-white sm:text-4xl">
              Core Methodologies
            </h2>
            <p className="mt-3 text-base text-slate-400">
              Technical precision meets intelligent execution.
            </p>
          </div>

          <div className="mt-14 grid grid-cols-1 gap-8 md:grid-cols-3">
            {/* Card 1: Precision Data Cleaning */}
            <div className="group relative rounded-3xl border border-slate-800 bg-slate-900/60 p-8 transition-all duration-300 hover:border-cyan-500/40 hover:bg-slate-900/90 hover:shadow-2xl hover:shadow-cyan-950/40">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-cyan-500/10 text-cyan-400 ring-1 ring-cyan-500/30 group-hover:bg-cyan-500 group-hover:text-slate-950 transition-colors">
                <Zap className="h-7 w-7" />
              </div>
              <h3 className="mt-6 text-xl font-bold text-white">
                Precision Data Cleaning
              </h3>
              <p className="mt-3 text-sm leading-relaxed text-slate-400">
                Micro-level field repair, acronym standardization, email validation, and duplicate elimination ensuring optimal data health and CRM readiness.
              </p>
              <button
                onClick={() => onNavigate("products", "cleaner")}
                className="mt-6 inline-flex items-center gap-1.5 text-xs font-bold text-cyan-400 hover:text-cyan-300 transition"
              >
                <span>Launch Cleaner Engine</span>
                <ArrowRight className="h-3.5 w-3.5" />
              </button>
            </div>

            {/* Card 2: AI-Powered Diagnostics */}
            <div className="group relative rounded-3xl border border-slate-800 bg-slate-900/60 p-8 transition-all duration-300 hover:border-cyan-500/40 hover:bg-slate-900/90 hover:shadow-2xl hover:shadow-cyan-950/40">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-blue-500/10 text-blue-400 ring-1 ring-blue-500/30 group-hover:bg-blue-500 group-hover:text-slate-950 transition-colors">
                <Activity className="h-7 w-7" />
              </div>
              <h3 className="mt-6 text-xl font-bold text-white">
                AI-Powered Diagnostics
              </h3>
              <p className="mt-3 text-sm leading-relaxed text-slate-400">
                Predictive matching, fuzzy record reconciliation, and variance anomaly detection across multi-format ledgers before thresholds are breached.
              </p>
              <button
                onClick={() => onNavigate("features", "recon-hub")}
                className="mt-6 inline-flex items-center gap-1.5 text-xs font-bold text-cyan-400 hover:text-cyan-300 transition"
              >
                <span>Explore Reconciliation</span>
                <ArrowRight className="h-3.5 w-3.5" />
              </button>
            </div>

            {/* Card 3: Facility Compliance */}
            <div className="group relative rounded-3xl border border-slate-800 bg-slate-900/60 p-8 transition-all duration-300 hover:border-cyan-500/40 hover:bg-slate-900/90 hover:shadow-2xl hover:shadow-cyan-950/40">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-emerald-500/10 text-emerald-400 ring-1 ring-emerald-500/30 group-hover:bg-emerald-500 group-hover:text-slate-950 transition-colors">
                <ShieldCheck className="h-7 w-7" />
              </div>
              <h3 className="mt-6 text-xl font-bold text-white">
                Facility &amp; Format Compliance
              </h3>
              <p className="mt-3 text-sm leading-relaxed text-slate-400">
                Rigorous structural auditing, ISO &amp; audit-trail alignment, ensuring adherence to enterprise compliance standards across Excel, CSV, PDF, and JSON.
              </p>
              <button
                onClick={() => onNavigate("features", "capabilities")}
                className="mt-6 inline-flex items-center gap-1.5 text-xs font-bold text-cyan-400 hover:text-cyan-300 transition"
              >
                <span>View Security Specs</span>
                <ArrowRight className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* ----------------------------------------------------------------------- */}
      {/* 3. THE MARQCLEAN ADVANTAGE (Bento Section matching DESIGN UI) */}
      {/* ----------------------------------------------------------------------- */}
      <section className="relative py-20 lg:py-24">
        <div className="mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <div className="mx-auto max-w-3xl text-center">
            <h2 className="text-3xl font-extrabold tracking-tight text-white sm:text-4xl">
              The MarqClean Advantage
            </h2>
            <p className="mt-3 text-base text-slate-400">
              Engineered from the ground up for financial analysts, auditors, and high-velocity data teams.
            </p>
          </div>

          <div className="mt-14 grid grid-cols-1 gap-8 lg:grid-cols-12">
            {/* Bento Big Left Feature */}
            <div className="relative overflow-hidden rounded-3xl border border-cyan-500/30 bg-gradient-to-br from-slate-900/90 via-[#071226] to-[#040816] p-8 lg:col-span-7 lg:p-10 flex flex-col justify-between">
              <div className="pointer-events-none absolute -right-20 -top-20 h-64 w-64 rounded-full bg-cyan-500/10 blur-3xl" />
              <div>
                <span className="rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3 py-1 text-[11px] font-bold uppercase tracking-wider text-cyan-300">
                  Cleanroom Standards
                </span>
                <h3 className="mt-4 text-2xl font-extrabold text-white sm:text-3xl">
                  ISO-Grade Spreadsheet &amp; Ledger Protocols
                </h3>
                <p className="mt-3 text-sm leading-relaxed text-slate-300">
                  We bring cleanroom rigor to spreadsheet workflows. Eliminate manual copying, corrupt Excel headers, and rounding discrepancies with automated multi-pass verification.
                </p>
              </div>

              <div className="mt-8 grid grid-cols-2 gap-4 rounded-2xl border border-slate-800 bg-slate-950/60 p-4 font-mono text-xs">
                <div>
                  <span className="text-[10px] uppercase text-slate-500">Validation Protocol</span>
                  <p className="mt-0.5 font-bold text-white">ISO 14644 Strict</p>
                </div>
                <div>
                  <span className="text-[10px] uppercase text-slate-500">File Processing</span>
                  <p className="mt-0.5 font-bold text-cyan-400">Up to 2GB Local Batches</p>
                </div>
              </div>
            </div>

            {/* Bento Right Column (Stacked Cards) */}
            <div className="flex flex-col gap-6 lg:col-span-5">
              {/* Stack 1: Zero Downtime Local Processing */}
              <div className="rounded-3xl border border-slate-800 bg-slate-900/60 p-7 transition hover:border-cyan-500/30">
                <div className="flex items-center gap-3">
                  <div className="rounded-xl bg-cyan-500/10 p-2.5 text-cyan-400">
                    <Lock className="h-5 w-5" />
                  </div>
                  <h3 className="text-lg font-bold text-white">
                    Zero Downtime Local Processing
                  </h3>
                </div>
                <p className="mt-3 text-xs leading-relaxed text-slate-400">
                  Non-intrusive in-browser execution designed to operate parallel to your active workflows without transmitting sensitive client records to third-party clouds.
                </p>
              </div>

              {/* Stack 2: Certified Specialists & AI Models */}
              <div className="rounded-3xl border border-slate-800 bg-slate-900/60 p-7 transition hover:border-cyan-500/30">
                <div className="flex items-center gap-3">
                  <div className="rounded-xl bg-blue-500/10 p-2.5 text-blue-400">
                    <Cpu className="h-5 w-5" />
                  </div>
                  <h3 className="text-lg font-bold text-white">
                    Certified Algorithmic Accuracy
                  </h3>
                </div>
                <p className="mt-3 text-xs leading-relaxed text-slate-400">
                  Our data sanitation algorithms undergo rigorous validation against complex financial instruments, trust accounts, multi-sheet workbooks, and global tax formats.
                </p>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* ----------------------------------------------------------------------- */}
      {/* 4. PROOF / METRICS BAR (Dark High-Contrast Section from DESIGN UI) */}
      {/* ----------------------------------------------------------------------- */}
      <section className="border-y border-cyan-500/20 bg-[#050b1a] py-12">
        <div className="mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <div className="grid grid-cols-1 gap-8 text-center sm:grid-cols-3">
            <div className="p-4">
              <div className="text-4xl font-extrabold text-white sm:text-5xl font-mono">
                <span className="text-cyan-400">99.99</span>%
              </div>
              <p className="mt-2 text-xs font-bold uppercase tracking-wider text-slate-400">
                Cleanroom Data Precision
              </p>
            </div>

            <div className="p-4 border-t border-slate-800 sm:border-t-0 sm:border-l sm:border-r">
              <div className="text-4xl font-extrabold text-white sm:text-5xl font-mono">
                <span className="text-cyan-400">24/7</span>
              </div>
              <p className="mt-2 text-xs font-bold uppercase tracking-wider text-slate-400">
                Autonomous Verification &amp; Reconciliation
              </p>
            </div>

            <div className="p-4 border-t border-slate-800 sm:border-t-0">
              <div className="text-4xl font-extrabold text-white sm:text-5xl font-mono">
                <span className="text-cyan-400">500k+</span>
              </div>
              <p className="mt-2 text-xs font-bold uppercase tracking-wider text-slate-400">
                Daily Records Sanitized Globally
              </p>
            </div>
          </div>
        </div>
      </section>

      {/* ----------------------------------------------------------------------- */}
      {/* 5. LIVE GLOBAL MARKET & RESEARCH CONNECTORS SECTION */}
      {/* ----------------------------------------------------------------------- */}
      <section className="relative py-20 lg:py-24">
        <div className="mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <LiveMarketHub />
        </div>
      </section>

      {/* ----------------------------------------------------------------------- */}
      {/* 6. SECURE YOUR DATA PIPELINE TODAY (Matches DESIGN UI layout) */}
      {/* ----------------------------------------------------------------------- */}
      <section className="relative overflow-hidden border-t border-slate-800/80 bg-gradient-to-b from-[#060c1d] to-[#02040a] py-20 lg:py-24">
        <div className="pointer-events-none absolute inset-x-0 bottom-0 h-64 bg-cyan-500/5 blur-3xl" />
        
        <div className="relative mx-auto max-w-4xl px-4 text-center sm:px-6 lg:px-8">
          <h2 className="text-3xl font-extrabold tracking-tight text-white sm:text-4xl lg:text-5xl">
            Secure Your Data Pipeline Today.
          </h2>
          <p className="mx-auto mt-4 max-w-2xl text-sm sm:text-base leading-relaxed text-slate-300">
            Don't let microscopic data anomalies compromise your macroscopic operations. Schedule a comprehensive diagnostic audit or test your files instantly.
          </p>

          <form
            onSubmit={handleQuickAuditSubmit}
            className="mx-auto mt-8 flex max-w-md flex-col gap-3 sm:flex-row"
          >
            <input
              type="email"
              placeholder="Enter your enterprise email..."
              value={quickAuditEmail}
              onChange={(e) => setQuickAuditEmail(e.target.value)}
              className="w-full rounded-2xl border border-slate-700 bg-slate-900/90 px-4 py-3.5 text-sm text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none focus:ring-1 focus:ring-cyan-500"
            />
            <button
              type="submit"
              className="rounded-2xl bg-gradient-to-r from-cyan-400 to-blue-600 px-7 py-3.5 text-xs font-black uppercase tracking-wider text-slate-950 transition hover:brightness-110 whitespace-nowrap shadow-lg shadow-cyan-500/25"
            >
              INITIATE AUDIT
            </button>
          </form>

          <p className="mt-4 text-xs text-slate-500">
            Trusted by asset managers, compliance officers, and analytics engineers across South Africa, the UK, and global hubs.
          </p>
        </div>
      </section>
    </div>
  );
};

export default HomePage;
