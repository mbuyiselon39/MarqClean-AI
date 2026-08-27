import React, { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { X, ShieldCheck, CheckCircle2, ArrowRight, Sparkles, Building2, Mail, User } from "lucide-react";

interface AuditModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialEmail?: string;
}

export const AuditModal: React.FC<AuditModalProps> = ({ isOpen, onClose, initialEmail = "" }) => {
  const [email, setEmail] = useState(initialEmail);
  const [fullName, setFullName] = useState("");
  const [company, setCompany] = useState("");
  const [datasetType, setDatasetType] = useState<string>("financial-ledgers");
  const [estimatedVolume, setEstimatedVolume] = useState<string>("10k-100k");
  const [submitted, setSubmitted] = useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!email) return;
    setSubmitted(true);
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
        {/* Backdrop */}
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="fixed inset-0 bg-slate-950/80 backdrop-blur-md"
        />

        {/* Dialog Content */}
        <motion.div
          initial={{ opacity: 0, scale: 0.95, y: 20 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 20 }}
          className="relative w-full max-w-xl overflow-hidden rounded-3xl border border-cyan-500/30 bg-slate-950 p-6 shadow-2xl shadow-cyan-950/60 sm:p-8 z-10"
        >
          {/* Close button */}
          <button
            onClick={onClose}
            className="absolute right-5 top-5 rounded-full p-2 text-slate-400 hover:bg-slate-800 hover:text-white transition"
          >
            <X className="h-5 w-5" />
          </button>

          {!submitted ? (
            <div>
              <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-cyan-400">
                <ShieldCheck className="h-4 w-4" />
                Enterprise Facility & Data Pipeline Audit
              </div>

              <h3 className="mt-2 text-2xl font-extrabold text-white sm:text-3xl">
                Initiate Precision Diagnostic Audit
              </h3>
              <p className="mt-2 text-sm text-slate-400">
                Receive an automated audit of your spreadsheet sanitization protocols, ledger reconciliation accuracy, and compliance alignment.
              </p>

              <form onSubmit={handleSubmit} className="mt-6 space-y-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-300">Corporate Work Email *</label>
                  <div className="relative mt-1">
                    <Mail className="absolute left-3.5 top-3 h-4 w-4 text-slate-500" />
                    <input
                      type="email"
                      required
                      placeholder="alex.smith@enterprise.com"
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      className="w-full rounded-xl border border-slate-800 bg-slate-900/90 py-2.5 pl-10 pr-4 text-sm text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none focus:ring-1 focus:ring-cyan-500"
                    />
                  </div>
                </div>

                <div className="grid gap-4 sm:grid-cols-2">
                  <div>
                    <label className="block text-xs font-semibold text-slate-300">Your Full Name</label>
                    <div className="relative mt-1">
                      <User className="absolute left-3.5 top-3 h-4 w-4 text-slate-500" />
                      <input
                        type="text"
                        placeholder="Alex Smith"
                        value={fullName}
                        onChange={(e) => setFullName(e.target.value)}
                        className="w-full rounded-xl border border-slate-800 bg-slate-900/90 py-2.5 pl-10 pr-4 text-sm text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300">Company / Organization</label>
                    <div className="relative mt-1">
                      <Building2 className="absolute left-3.5 top-3 h-4 w-4 text-slate-500" />
                      <input
                        type="text"
                        placeholder="Apex Financial Corp"
                        value={company}
                        onChange={(e) => setCompany(e.target.value)}
                        className="w-full rounded-xl border border-slate-800 bg-slate-900/90 py-2.5 pl-10 pr-4 text-sm text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
                      />
                    </div>
                  </div>
                </div>

                <div className="grid gap-4 sm:grid-cols-2">
                  <div>
                    <label className="block text-xs font-semibold text-slate-300">Primary Data Category</label>
                    <select
                      value={datasetType}
                      onChange={(e) => setDatasetType(e.target.value)}
                      className="mt-1 w-full rounded-xl border border-slate-800 bg-slate-900 py-2.5 px-3 text-sm text-white focus:border-cyan-500 focus:outline-none"
                    >
                      <option value="financial-ledgers">Financial & Bank Ledgers</option>
                      <option value="client-funds">Client Funds & Trust Accounts</option>
                      <option value="marketing-leads">CRM & Marketing Lead Lists</option>
                      <option value="multi-format">PDF / Excel / Word Comparison</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300">Monthly Record Volume</label>
                    <select
                      value={estimatedVolume}
                      onChange={(e) => setEstimatedVolume(e.target.value)}
                      className="mt-1 w-full rounded-xl border border-slate-800 bg-slate-900 py-2.5 px-3 text-sm text-white focus:border-cyan-500 focus:outline-none"
                    >
                      <option value="under-10k">&lt; 10,000 records</option>
                      <option value="10k-100k">10,000 - 100,000 records</option>
                      <option value="100k-1m">100,000 - 1,000,000 records</option>
                      <option value="1m-plus">1,000,000+ Enterprise records</option>
                    </select>
                  </div>
                </div>

                <div className="mt-6">
                  <button
                    type="submit"
                    className="flex w-full items-center justify-center gap-2 rounded-2xl bg-gradient-to-r from-cyan-500 to-blue-600 py-3.5 text-sm font-bold text-slate-950 transition hover:brightness-110 shadow-lg shadow-cyan-500/25"
                  >
                    <Sparkles className="h-4 w-4" />
                    Request Diagnostic Audit Report
                    <ArrowRight className="h-4 w-4" />
                  </button>
                </div>

                <p className="text-center text-[11px] text-slate-500">
                  Zero data exposure. Processing is verified locally in browser with ISO cleanroom privacy standards.
                </p>
              </form>
            </div>
          ) : (
            <div className="py-4 text-center">
              <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-cyan-500/20 text-cyan-400 ring-1 ring-cyan-500/40">
                <CheckCircle2 className="h-8 w-8" />
              </div>

              <h3 className="mt-4 text-2xl font-extrabold text-white">
                Diagnostic Request Initialized!
              </h3>
              <p className="mt-2 text-sm text-slate-300">
                Your enterprise data sanitation profile has been registered for <span className="font-bold text-cyan-300">{email}</span>.
              </p>

              <div className="mt-6 rounded-2xl border border-slate-800 bg-slate-900/60 p-4 text-left font-mono text-xs">
                <div className="flex justify-between border-b border-slate-800 pb-2">
                  <span className="text-slate-400">Baseline Clarity Score:</span>
                  <span className="font-bold text-emerald-400">98.4% (Optimal)</span>
                </div>
                <div className="flex justify-between border-b border-slate-800 py-2">
                  <span className="text-slate-400">Reconciliation Precision:</span>
                  <span className="font-bold text-cyan-400">Zero Variance Tolerance</span>
                </div>
                <div className="flex justify-between pt-2">
                  <span className="text-slate-400">Engine Protocol:</span>
                  <span className="font-bold text-white">MarqClean AI v2.4 Enterprise</span>
                </div>
              </div>

              <button
                onClick={onClose}
                className="mt-6 w-full rounded-2xl bg-slate-800 py-3 text-sm font-bold text-white hover:bg-slate-700 transition"
              >
                Close & Return to Suite
              </button>
            </div>
          )}
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

export default AuditModal;
