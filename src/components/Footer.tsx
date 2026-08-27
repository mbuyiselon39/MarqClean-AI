import React from "react";
import { MarqLogo } from "./MarqLogo";
import { ShieldCheck, Lock, Globe, Server } from "lucide-react";
import { NavPageKey } from "./Navbar";

interface FooterProps {
  onNavigate: (page: NavPageKey, subTab?: string) => void;
  onRequestAudit: () => void;
}

export const Footer: React.FC<FooterProps> = ({ onNavigate, onRequestAudit }) => {
  return (
    <footer className="border-t border-cyan-500/20 bg-[#030611] text-slate-400">
      {/* Top Banner: Enterprise Security & Standard */}
      <div className="border-b border-slate-800/80 bg-slate-950/60 py-6">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-4 px-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-3">
            <ShieldCheck className="h-5 w-5 text-cyan-400" />
            <span className="text-xs font-semibold text-slate-300">
              Cleanroom Precision Grade: <span className="text-cyan-300">ISO 14644 &amp; SOC2 Type II Compatible Processing</span>
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-6 text-xs text-slate-400">
            <span className="flex items-center gap-1.5">
              <Lock className="h-3.5 w-3.5 text-emerald-400" />
              100% In-Browser Privacy
            </span>
            <span className="flex items-center gap-1.5">
              <Server className="h-3.5 w-3.5 text-blue-400" />
              Zero Server Data Retention
            </span>
            <span className="flex items-center gap-1.5">
              <Globe className="h-3.5 w-3.5 text-cyan-400" />
              Global Financial Formats
            </span>
          </div>
        </div>
      </div>

      {/* Main Footer Links */}
      <div className="mx-auto max-w-7xl px-4 py-12 sm:px-6 lg:px-8 lg:py-16">
        <div className="grid grid-cols-1 gap-8 md:grid-cols-2 lg:grid-cols-5">
          {/* Col 1: Brand Info */}
          <div className="lg:col-span-2 space-y-4">
            <MarqLogo size="lg" showText={true} showTagline={true} />
            <p className="max-w-sm text-xs leading-relaxed text-slate-400">
              MarqClean AI is an enterprise-grade AI data automation and reconciliation platform. Engineered to transform complex, disordered spreadsheets, bank statements, and multi-format ledgers into zero-defect Excel intelligence.
            </p>
            <div className="pt-2 text-xs text-slate-500">
              A flagship product of <strong className="text-slate-300">Vertex Stream Group</strong>. Hosted at{" "}
              <a
                href="https://marqcleanai.vertexsg.co.za/"
                target="_blank"
                rel="noreferrer"
                className="text-cyan-400 hover:underline"
              >
                marqcleanai.vertexsg.co.za
              </a>
            </div>
          </div>

          {/* Col 2: Products */}
          <div>
            <h3 className="text-xs font-extrabold uppercase tracking-wider text-white">Products</h3>
            <ul className="mt-4 space-y-2.5 text-xs">
              <li>
                <button
                  onClick={() => onNavigate("products", "cleaner")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Quick Data &amp; CSV Cleaner
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("products", "converter")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  CSV to Excel Converter
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("products", "formulas")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Excel Formulas &amp; Formatting
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("products", "bank")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Bank Ledger X (PDF Extractor)
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("products", "client-funds")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Client Funds Manager
                </button>
              </li>
            </ul>
          </div>

          {/* Col 3: Features & Hubs */}
          <div>
            <h3 className="text-xs font-extrabold uppercase tracking-wider text-white">Features</h3>
            <ul className="mt-4 space-y-2.5 text-xs">
              <li>
                <button
                  onClick={() => onNavigate("features", "recon-hub")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Verification &amp; Recon Hub
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("features", "data-toolbox")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Free Data Toolbox
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("features", "excel-automation")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Excel Automation Engine
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("features", "capabilities")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Platform Architecture
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("excel-academy")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Interactive Excel Academy
                </button>
              </li>
            </ul>
          </div>

          {/* Col 4: Free Data Tools & Audit */}
          <div>
            <h3 className="text-xs font-extrabold uppercase tracking-wider text-white">Free Data Tools</h3>
            <ul className="mt-4 space-y-2.5 text-xs">
              <li>
                <button
                  onClick={() => onNavigate("free-data-tools")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Text &amp; RFC Email Filter
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("free-data-tools")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  CSV &bull; JSON &bull; TSV Converter
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("free-data-tools")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Column Header Sanitizer
                </button>
              </li>
              <li>
                <button
                  onClick={() => onNavigate("free-data-tools")}
                  className="hover:text-cyan-300 transition text-left"
                >
                  Financial Accounting Cleaner
                </button>
              </li>
              <li className="pt-2">
                <button
                  onClick={onRequestAudit}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-cyan-500/20 px-3 py-1.5 font-bold text-cyan-300 hover:bg-cyan-500 hover:text-slate-950 transition"
                >
                  Request Diagnostic Audit
                </button>
              </li>
            </ul>
          </div>
        </div>

        {/* Bottom copyright */}
        <div className="mt-12 flex flex-col items-center justify-between border-t border-slate-800/80 pt-8 sm:flex-row text-xs text-slate-500">
          <p>© {new Date().getFullYear()} MarqClean AI. Owned and operated by Vertex Stream Group (Pty) Ltd. All rights reserved.</p>
          <div className="mt-4 flex gap-6 sm:mt-0">
            <span>Privacy Policy</span>
            <span>Terms of Service</span>
            <span>Security Statement</span>
          </div>
        </div>
      </div>
    </footer>
  );
};

export default Footer;
