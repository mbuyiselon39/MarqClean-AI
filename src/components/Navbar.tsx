import React, { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { MarqLogo } from "./MarqLogo";
import {
  ChevronDown,
  Wand2,
  FileSpreadsheet,
  Calculator,
  Landmark,
  Layers,
  CheckCircle,
  Table,
  Cpu,
  Menu,
  X,
  ArrowRight,
  Shield,
} from "lucide-react";

export type NavPageKey = "home" | "products" | "features" | "free-data-tools" | "excel-academy" | "contact" | "about";

interface NavbarProps {
  currentPage: NavPageKey;
  onNavigate: (page: NavPageKey, subTab?: string) => void;
  onRequestAudit: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({ currentPage, onNavigate, onRequestAudit }) => {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [productsOpen, setProductsOpen] = useState(false);
  const [featuresOpen, setFeaturesOpen] = useState(false);

  const handleLinkClick = (page: NavPageKey, subTab?: string) => {
    onNavigate(page, subTab);
    setMobileMenuOpen(false);
    setProductsOpen(false);
    setFeaturesOpen(false);
  };

  return (
    <header className="sticky top-0 z-40 w-full border-b border-cyan-500/15 bg-[#040711]/90 backdrop-blur-xl transition-all">
      <div className="mx-auto flex h-20 max-w-7xl items-center justify-between px-4 sm:px-6 lg:px-8">
        {/* Left: Brand Logo matching MarqC 2 LOGO 2 */}
        <div
          onClick={() => handleLinkClick("home")}
          className="cursor-pointer transition-transform hover:scale-[1.02]"
        >
          <MarqLogo size="md" showText={true} showTagline={true} />
        </div>

        {/* Center: Desktop Navigation items matching DESIGN UI */}
        <nav className="hidden lg:flex items-center gap-1 xl:gap-2">
          {/* Home */}
          <button
            onClick={() => handleLinkClick("home")}
            className={`rounded-xl px-3.5 py-2 text-xs font-bold uppercase tracking-wider transition ${
              currentPage === "home"
                ? "bg-cyan-500/15 text-cyan-300 ring-1 ring-cyan-500/30"
                : "text-slate-300 hover:text-white hover:bg-slate-900/60"
            }`}
          >
            Overview
          </button>

          {/* Products Dropdown */}
          <div
            className="relative"
            onMouseEnter={() => setProductsOpen(true)}
            onMouseLeave={() => setProductsOpen(false)}
          >
            <button
              onClick={() => handleLinkClick("products")}
              className={`flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold uppercase tracking-wider transition ${
                currentPage === "products"
                  ? "bg-cyan-500/15 text-cyan-300 ring-1 ring-cyan-500/30"
                  : "text-slate-300 hover:text-white hover:bg-slate-900/60"
              }`}
            >
              <span>Products</span>
              <ChevronDown className={`h-3.5 w-3.5 transition-transform ${productsOpen ? "rotate-180" : ""}`} />
            </button>

            <AnimatePresence>
              {productsOpen && (
                <motion.div
                  initial={{ opacity: 0, y: 8 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: 8 }}
                  transition={{ duration: 0.15 }}
                  className="absolute left-0 top-full mt-1 w-80 rounded-2xl border border-cyan-500/25 bg-[#070c1a] p-3 shadow-2xl shadow-cyan-950/80 backdrop-blur-2xl"
                >
                  <div className="space-y-1">
                    <button
                      onClick={() => handleLinkClick("products", "cleaner")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-cyan-500/20 p-2 text-cyan-400 group-hover:bg-cyan-500 group-hover:text-slate-950 transition">
                        <Wand2 className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Quick Data & CSV Cleaner</div>
                        <p className="text-[11px] text-slate-400">Standardize CRM lists, fix casing, remove duplicates.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("products", "converter")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-blue-500/20 p-2 text-blue-400 group-hover:bg-blue-500 group-hover:text-slate-950 transition">
                        <FileSpreadsheet className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">CSV to Excel Converter</div>
                        <p className="text-[11px] text-slate-400">Intelligent Text to Columns & delimiter separation.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("products", "formulas")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-emerald-500/20 p-2 text-emerald-400 group-hover:bg-emerald-500 group-hover:text-slate-950 transition">
                        <Calculator className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Excel Formulas & Formatting</div>
                        <p className="text-[11px] text-slate-400">Automate subtotals, currencies, dates & formula logic.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("products", "bank")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-amber-500/20 p-2 text-amber-400 group-hover:bg-amber-500 group-hover:text-slate-950 transition">
                        <Landmark className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Bank Ledger X</div>
                        <p className="text-[11px] text-slate-400">Turn PDF bank statements into structured Excel / QIF.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("products", "client-funds")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-purple-500/20 p-2 text-purple-400 group-hover:bg-purple-500 group-hover:text-slate-950 transition">
                        <Layers className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Client Funds Manager</div>
                        <p className="text-[11px] text-slate-400">Root SRN grouping, trust subtotaling & validation.</p>
                      </div>
                    </button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* Features Dropdown */}
          <div
            className="relative"
            onMouseEnter={() => setFeaturesOpen(true)}
            onMouseLeave={() => setFeaturesOpen(false)}
          >
            <button
              onClick={() => handleLinkClick("features")}
              className={`flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold uppercase tracking-wider transition ${
                currentPage === "features"
                  ? "bg-cyan-500/15 text-cyan-300 ring-1 ring-cyan-500/30"
                  : "text-slate-300 hover:text-white hover:bg-slate-900/60"
              }`}
            >
              <span>Features</span>
              <ChevronDown className={`h-3.5 w-3.5 transition-transform ${featuresOpen ? "rotate-180" : ""}`} />
            </button>

            <AnimatePresence>
              {featuresOpen && (
                <motion.div
                  initial={{ opacity: 0, y: 8 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: 8 }}
                  transition={{ duration: 0.15 }}
                  className="absolute left-0 top-full mt-1 w-80 rounded-2xl border border-cyan-500/25 bg-[#070c1a] p-3 shadow-2xl shadow-cyan-950/80 backdrop-blur-2xl"
                >
                  <div className="space-y-1">
                    <button
                      onClick={() => handleLinkClick("features", "recon-hub")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-cyan-500/20 p-2 text-cyan-400 group-hover:bg-cyan-500 group-hover:text-slate-950 transition">
                        <CheckCircle className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Verification & Recon Hub</div>
                        <p className="text-[11px] text-slate-400">Multi-format comparison & variance exception reporting.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("features", "data-toolbox")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-blue-500/20 p-2 text-blue-400 group-hover:bg-blue-500 group-hover:text-slate-950 transition">
                        <Table className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Free Data Toolbox</div>
                        <p className="text-[11px] text-slate-400">Deduplication, case normalization & multi-sheet engine.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("features", "excel-automation")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-emerald-500/20 p-2 text-emerald-400 group-hover:bg-emerald-500 group-hover:text-slate-950 transition">
                        <Cpu className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Excel Automation Engine</div>
                        <p className="text-[11px] text-slate-400">DAX patterns, relationships & dynamic array calculations.</p>
                      </div>
                    </button>

                    <button
                      onClick={() => handleLinkClick("features", "capabilities")}
                      className="flex w-full items-start gap-3 rounded-xl p-2.5 text-left transition hover:bg-slate-800/60 group"
                    >
                      <div className="mt-0.5 rounded-lg bg-purple-500/20 p-2 text-purple-400 group-hover:bg-purple-500 group-hover:text-slate-950 transition">
                        <Shield className="h-4 w-4" />
                      </div>
                      <div>
                        <div className="text-xs font-bold text-white group-hover:text-cyan-300 transition">Platform Capabilities</div>
                        <p className="text-[11px] text-slate-400">Zero-downtime, local security & 2GB batch streaming.</p>
                      </div>
                    </button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* Dedicated Free Data Tools Tab */}
          <button
            onClick={() => handleLinkClick("free-data-tools")}
            className={`rounded-xl px-3.5 py-2 text-xs font-bold uppercase tracking-wider transition ${
              currentPage === "free-data-tools"
                ? "bg-cyan-500/15 text-cyan-300 ring-1 ring-cyan-500/30"
                : "text-slate-300 hover:text-white hover:bg-slate-900/60"
            }`}
          >
            <span className="flex items-center gap-1.5">
              <span>Free Data Tools</span>
              <span className="rounded-full bg-cyan-500/20 px-1.5 py-0.2 text-[9px] font-extrabold text-cyan-300">
                LIVE
              </span>
            </span>
          </button>

          {/* Excel Academy Tab */}
          <button
            onClick={() => handleLinkClick("excel-academy")}
            className={`rounded-xl px-3.5 py-2 text-xs font-bold uppercase tracking-wider transition ${
              currentPage === "excel-academy"
                ? "bg-cyan-500/15 text-cyan-300 ring-1 ring-cyan-500/30"
                : "text-slate-300 hover:text-white hover:bg-slate-900/60"
            }`}
          >
            Excel Academy
          </button>
        </nav>

        {/* Right side CTAs matching DESIGN UI */}
        <div className="hidden lg:flex items-center gap-3">
          <div className="flex items-center gap-2 rounded-full border border-emerald-500/30 bg-emerald-950/40 px-3 py-1 text-[11px] font-semibold text-emerald-300">
            <span className="h-2 w-2 rounded-full bg-emerald-400 animate-pulse" />
            <span>AI Pipeline Active</span>
          </div>

          <button
            onClick={onRequestAudit}
            className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 px-5 py-2.5 text-xs font-extrabold uppercase tracking-wider text-slate-950 transition hover:brightness-110 shadow-lg shadow-cyan-500/25 active:scale-95"
          >
            <span>Request Audit</span>
            <ArrowRight className="h-3.5 w-3.5" />
          </button>
        </div>

        {/* Mobile menu toggle */}
        <div className="flex lg:hidden items-center gap-2">
          <button
            onClick={onRequestAudit}
            className="rounded-lg bg-cyan-500 px-3 py-1.5 text-xs font-bold text-slate-950"
          >
            Audit
          </button>
          <button
            onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
            className="rounded-xl border border-slate-800 p-2 text-slate-300 hover:text-white"
          >
            {mobileMenuOpen ? <X className="h-6 w-6" /> : <Menu className="h-6 w-6" />}
          </button>
        </div>
      </div>

      {/* Mobile Menu */}
      <AnimatePresence>
        {mobileMenuOpen && (
          <motion.div
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: "auto" }}
            exit={{ opacity: 0, height: 0 }}
            className="border-b border-slate-800 bg-[#060b18] px-4 py-6 lg:hidden"
          >
            <div className="space-y-3 font-semibold text-sm">
              <button
                onClick={() => handleLinkClick("home")}
                className="block w-full text-left py-2 px-3 rounded-lg hover:bg-slate-800 text-slate-200"
              >
                Overview
              </button>
              <button
                onClick={() => handleLinkClick("products")}
                className="block w-full text-left py-2 px-3 rounded-lg hover:bg-slate-800 text-slate-200"
              >
                Products (CSV Cleaner, Converter, Bank Ledger X)
              </button>
              <button
                onClick={() => handleLinkClick("features")}
                className="block w-full text-left py-2 px-3 rounded-lg hover:bg-slate-800 text-slate-200"
              >
                Features (Reconciliation Hub, Data Toolbox)
              </button>
              <button
                onClick={() => handleLinkClick("free-data-tools")}
                className="block w-full text-left py-2 px-3 rounded-lg hover:bg-slate-800 text-cyan-300 font-bold"
              >
                Free Data Tools
              </button>
              <button
                onClick={() => handleLinkClick("excel-academy")}
                className="block w-full text-left py-2 px-3 rounded-lg hover:bg-slate-800 text-slate-200"
              >
                Excel Academy
              </button>

              <div className="pt-3">
                <button
                  onClick={() => {
                    setMobileMenuOpen(false);
                    onRequestAudit();
                  }}
                  className="w-full rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 py-3 text-center text-xs font-bold uppercase tracking-wider text-slate-950"
                >
                  Request Diagnostic Audit
                </button>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </header>
  );
};

export default Navbar;
