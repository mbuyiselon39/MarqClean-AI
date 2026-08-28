import React, { useState } from "react";
import { motion } from "framer-motion";
import {
  Wand2,
  FileSpreadsheet,
  Calculator,
  Landmark,
  Layers,
  Download,
  CheckCircle,
  Sparkles,
  ArrowRight,
} from "lucide-react";
import * as XLSX from "xlsx";
import ClientFunds from "../reconciliation/ClientFunds";
import { CopyDataButton } from "../components/CopyDataButton";
import { BulkDataCleaner } from "../components/BulkDataCleaner";

export interface ProductsPageProps {
  initialSubTab?: string;
  onRequestAudit: () => void;
}

export const ProductsPage: React.FC<ProductsPageProps> = ({ initialSubTab = "cleaner", onRequestAudit }) => {
  const [activeTab, setActiveTab] = useState<"cleaner" | "converter" | "formulas" | "bank" | "client-funds">(
    (initialSubTab as any) || "cleaner"
  );

  // CSV to Excel Converter State
  const [rawCsvText, setRawCsvText] = useState("");
  const [parsedMatrix, setParsedMatrix] = useState<string[][] | null>(null);
  const [delimiter, setDelimiter] = useState<"," | ";" | "\t" | "|">(",");

  // Formulas State
  const [selectedCalc, setSelectedCalc] = useState<"sum" | "average" | "subtotal">("subtotal");
  const [generatedFormula, setGeneratedFormula] = useState("=SUBTOTAL(109, [Net_Amount])");

  // Bank Ledger State
  const bankParsedRows = [
    { date: "2026-08-01", desc: "CLIENT SETTLEMENT REF #9901", amount: 15420.0, balance: 115420.0 },
    { date: "2026-08-03", desc: "MONTHLY AUDIT RETAINER", amount: -2850.0, balance: 112570.0 },
    { date: "2026-08-05", desc: "WIRE TRANSFER INWARD ZAR", amount: 48900.0, balance: 161470.0 },
  ];

  const handleConvertCsv = () => {
    if (!rawCsvText.trim()) return;
    const lines = rawCsvText.trim().split("\n");
    const matrix = lines.map((line) => line.split(delimiter).map((c) => c.trim().replace(/^["']|["']$/g, "")));
    setParsedMatrix(matrix);
  };

  const downloadConvertedWorkbook = () => {
    if (!parsedMatrix) return;
    const ws = XLSX.utils.aoa_to_sheet(parsedMatrix);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Converted Sheet");
    XLSX.writeFile(wb, "MarqClean_Converted_Spreadsheet.xlsx");
  };

  return (
    <div className="min-h-screen bg-[#030611] text-slate-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-7xl">
        {/* Top Header */}
        <div className="flex flex-col gap-4 border-b border-slate-800 pb-8 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <Sparkles className="h-3.5 w-3.5" />
              MarqClean Enterprise Product Suite
            </div>
            <h1 className="mt-3 text-3xl font-extrabold text-white sm:text-4xl">
              Precision Spreadsheet &amp; Data Sanitation Products
            </h1>
            <p className="mt-2 text-sm text-slate-400 max-w-2xl">
              Select any of our dedicated client-side tools to clean CSVs, convert delimiters, generate Excel formulas, or extract bank ledger transactions.
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

        {/* Tab Navigation */}
        <div className="mt-8 flex flex-wrap gap-2 sm:gap-3">
          <button
            onClick={() => setActiveTab("cleaner")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "cleaner"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Wand2 className="h-4 w-4 text-cyan-400" />
            <span>Bulk Data &amp; CSV Cleaner</span>
          </button>

          <button
            onClick={() => setActiveTab("converter")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "converter"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <FileSpreadsheet className="h-4 w-4 text-blue-400" />
            <span>CSV to Excel Converter</span>
          </button>

          <button
            onClick={() => setActiveTab("formulas")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "formulas"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Calculator className="h-4 w-4 text-emerald-400" />
            <span>Excel Formulas &amp; Formatting</span>
          </button>

          <button
            onClick={() => setActiveTab("bank")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "bank"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Landmark className="h-4 w-4 text-amber-400" />
            <span>Bank Ledger X</span>
          </button>

          <button
            onClick={() => setActiveTab("client-funds")}
            className={`flex items-center gap-2 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-bold transition ${
              activeTab === "client-funds"
                ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
                : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:text-white"
            }`}
          >
            <Layers className="h-4 w-4 text-purple-400" />
            <span>Client Funds Manager</span>
          </button>
        </div>

        {/* ------------------------------------------------------------- */}
        {/* PRODUCT 1: Bulk Data & CSV Cleaner */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "cleaner" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8">
            <BulkDataCleaner onRequestAudit={onRequestAudit} />
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* PRODUCT 2: CSV to Excel Converter */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "converter" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-6">
            <div className="rounded-3xl border border-blue-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl">
              <h2 className="text-xl font-bold text-white">CSV to Excel Converter &amp; Delimiter Separator</h2>
              <p className="text-xs text-slate-400 mt-1">
                Split custom delimited raw text, unformatted exports, or TSV streams into structured, typed Excel worksheets with clean borders.
              </p>

              <div className="mt-6 space-y-4">
                <div className="flex items-center gap-3">
                  <span className="text-xs font-semibold text-slate-300">Choose Delimiter:</span>
                  {(
                    [
                      { label: "Comma (,)", val: "," },
                      { label: "Semicolon (;)", val: ";" },
                      { label: "Tab (\\t)", val: "\t" },
                      { label: "Pipe (|)", val: "|" },
                    ] as const
                  ).map((d) => (
                    <button
                      key={d.val}
                      onClick={() => setDelimiter(d.val)}
                      className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${
                        delimiter === d.val
                          ? "bg-blue-500 text-slate-950 font-bold"
                          : "border border-slate-800 bg-slate-900 text-slate-400 hover:text-white"
                      }`}
                    >
                      {d.label}
                    </button>
                  ))}
                </div>

                <textarea
                  rows={5}
                  placeholder={`Paste raw delimited text here...\nExample:\nName${delimiter}Email${delimiter}Role\nJohn Smith${delimiter}john@apex.com${delimiter}Lead Auditor`}
                  value={rawCsvText}
                  onChange={(e) => setRawCsvText(e.target.value)}
                  className="w-full rounded-2xl border border-slate-800 bg-slate-900/90 p-4 font-mono text-xs text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
                />

                <div className="flex items-center gap-3">
                  <button
                    onClick={handleConvertCsv}
                    className="rounded-xl bg-blue-500 px-5 py-2.5 text-xs font-bold text-slate-950 hover:bg-blue-400 transition shadow-md shadow-blue-500/20"
                  >
                    Convert &amp; Parse Columns
                  </button>
                </div>

                {parsedMatrix && (
                  <div className="mt-6 space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <span className="text-xs font-semibold text-emerald-400 flex items-center gap-1.5">
                        <CheckCircle className="h-4 w-4" />
                        Converted {parsedMatrix.length} rows successfully
                      </span>
                      <div className="flex items-center flex-wrap gap-2">
                        <CopyDataButton
                          payload={{ type: "matrix", data: parsedMatrix }}
                          showCsv={true}
                          showTsv={true}
                          showJson={true}
                          size="xs"
                          buttonTheme="dark"
                        />
                        <button
                          onClick={downloadConvertedWorkbook}
                          className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-blue-500 to-cyan-500 px-3.5 py-1.5 text-xs font-bold text-slate-950 hover:brightness-110 transition"
                        >
                          <Download className="h-3.5 w-3.5" />
                          <span>Export Excel (.xlsx)</span>
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto rounded-2xl border border-slate-800 bg-slate-900/60 font-mono text-xs">
                      <table className="min-w-full text-left">
                        <tbody className="divide-y divide-slate-800/60 text-slate-200">
                          {parsedMatrix.map((row, rI) => (
                            <tr key={rI} className={rI === 0 ? "bg-slate-900 font-bold text-blue-300" : "hover:bg-slate-800/40"}>
                              {row.map((cell, cI) => (
                                <td key={cI} className="px-4 py-2.5">{cell}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* PRODUCT 3: Excel Formulas & Formatting */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "formulas" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-6">
            <div className="rounded-3xl border border-emerald-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl">
              <h2 className="text-xl font-bold text-white">Excel Formulas &amp; Automated Financial Formatting</h2>
              <p className="text-xs text-slate-400 mt-1">
                Automate advanced formulas, dynamic subtotals, structured table references, and currency/date number formatting.
              </p>

              <div className="mt-6 grid gap-6 sm:grid-cols-2">
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5 space-y-4">
                  <h3 className="text-sm font-bold text-emerald-400">Formula Selector</h3>
                  <div className="space-y-2">
                    {[
                      { key: "subtotal", name: "Dynamic Filtered Subtotal", formula: "=SUBTOTAL(109, [Net_Amount])" },
                      { key: "sum", name: "Aggregate Sum with Condition", formula: '=SUMIFS([Net_Amount], [Status], "Closed")' },
                      { key: "xlookup", name: "Modern Safe Lookup", formula: '=XLOOKUP(A2, Master[ID], Master[Balance], 0)' },
                    ].map((item) => (
                      <div
                        key={item.key}
                        onClick={() => {
                          setSelectedCalc(item.key as any);
                          setGeneratedFormula(item.formula);
                        }}
                        className={`cursor-pointer rounded-xl border p-3 text-xs transition ${
                          selectedCalc === item.key
                            ? "border-emerald-500 bg-emerald-950/40 text-white"
                            : "border-slate-800 bg-slate-900/80 text-slate-400 hover:text-white"
                        }`}
                      >
                        <div className="font-bold text-slate-200">{item.name}</div>
                        <code className="mt-1 block text-[11px] text-emerald-400 font-mono">{item.formula}</code>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5 space-y-4">
                  <div className="flex items-center justify-between">
                    <h3 className="text-sm font-bold text-white">Generated Formula &amp; Syntax</h3>
                    <CopyDataButton
                      payload={{ type: "text", text: generatedFormula }}
                      buttonTheme="cyan"
                      size="xs"
                      titlePrefix="Formula:"
                    />
                  </div>
                  <div className="rounded-xl border border-slate-800 bg-slate-950 p-4 font-mono text-xs text-emerald-300">
                    {generatedFormula}
                  </div>
                  <p className="text-xs text-slate-400">
                    This formula automatically avoids double-counting subtotals when filters are active in Microsoft 365 or Excel 2019+.
                  </p>
                </div>
              </div>
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* PRODUCT 4: Bank Ledger X */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "bank" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-6">
            <div className="rounded-3xl border border-amber-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl">
              <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h2 className="text-xl font-bold text-white">Bank Ledger X - Statement Parser &amp; QIF Generator</h2>
                  <p className="text-xs text-slate-400 mt-1">
                    Extract transactional lines from bank statements, standardize credit/debit figures, and export directly to QIF or Excel.
                  </p>
                </div>
                <div className="flex items-center flex-wrap gap-2">
                  <CopyDataButton
                    payload={{ type: "records", data: bankParsedRows }}
                    showCsv={true}
                    showTsv={true}
                    showJson={true}
                    size="xs"
                    buttonTheme="dark"
                  />
                  <button
                    onClick={() => {
                      const ws = XLSX.utils.json_to_sheet(bankParsedRows);
                      const wb = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, "Bank Ledger");
                      XLSX.writeFile(wb, "MarqClean_Bank_Ledger.xlsx");
                    }}
                    className="flex items-center gap-2 rounded-xl bg-amber-500 px-3.5 py-1.5 text-xs font-bold text-slate-950 hover:bg-amber-400 transition"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>Export Excel (.xlsx)</span>
                  </button>
                </div>
              </div>

              <div className="mt-6 overflow-x-auto rounded-2xl border border-slate-800 bg-slate-900/60 font-mono text-xs">
                <table className="min-w-full text-left">
                  <thead className="border-b border-slate-800 bg-slate-900 text-amber-300">
                    <tr>
                      <th className="px-4 py-3">Posting Date</th>
                      <th className="px-4 py-3">Transaction Description</th>
                      <th className="px-4 py-3 text-right">Amount (ZAR/USD)</th>
                      <th className="px-4 py-3 text-right">Running Balance</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 text-slate-200">
                    {bankParsedRows.map((r, i) => (
                      <tr key={i} className="hover:bg-slate-800/40">
                        <td className="px-4 py-2.5">{r.date}</td>
                        <td className="px-4 py-2.5 font-sans font-medium text-white">{r.desc}</td>
                        <td className={`px-4 py-2.5 text-right font-bold ${r.amount >= 0 ? "text-emerald-400" : "text-rose-400"}`}>
                          {r.amount >= 0 ? "+" : ""}{r.amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                        </td>
                        <td className="px-4 py-2.5 text-right text-slate-300">
                          {r.balance.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </motion.div>
        )}

        {/* ------------------------------------------------------------- */}
        {/* PRODUCT 5: Client Funds Manager */}
        {/* ------------------------------------------------------------- */}
        {activeTab === "client-funds" && (
          <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-8 space-y-6">
            <div className="rounded-3xl border border-purple-500/20 bg-slate-950/80 p-4 sm:p-6 backdrop-blur-xl">
              <ClientFunds onExit={() => setActiveTab("cleaner")} />
            </div>
          </motion.div>
        )}
      </div>
    </div>
  );
};

export default ProductsPage;
