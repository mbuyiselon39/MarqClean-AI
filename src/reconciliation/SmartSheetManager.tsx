import React, { useState } from "react";
import { FileSpreadsheet, Download, Trash2 } from "lucide-react";

export const SmartSheetManager: React.FC = () => {
  const [sheets, setSheets] = useState<Array<{ name: string; rowCount: number; selected: boolean }>>([
    { name: "General_Ledger_2026", rowCount: 1450, selected: true },
    { name: "Disbursements_Q3", rowCount: 890, selected: true },
    { name: "Tax_Withholding_Exemptions", rowCount: 320, selected: false },
    { name: "Notes_Scratchpad", rowCount: 45, selected: false },
  ]);

  const toggleSelect = (name: string) => {
    setSheets((prev) => prev.map((s) => (s.name === name ? { ...s, selected: !s.selected } : s)));
  };

  const removeSheet = (name: string) => {
    setSheets((prev) => prev.filter((s) => s.name !== name));
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-slate-800 pb-4">
        <div>
          <h3 className="text-base font-bold text-white">Smart Sheet Manager &amp; Multi-Tab Optimizer</h3>
          <p className="text-xs text-slate-400 mt-0.5">
            Audit, re-order, purge scratchpad tabs, and consolidate workbook sheets before dispatch.
          </p>
        </div>

        <button className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 px-4 py-2 text-xs font-bold text-slate-950 hover:brightness-110 transition">
          <Download className="h-4 w-4" />
          <span>Export Cleaned Multi-Sheet Workbook</span>
        </button>
      </div>

      <div className="divide-y divide-slate-800 rounded-2xl border border-slate-800 bg-slate-900/60 overflow-hidden">
        {sheets.map((s) => (
          <div key={s.name} className="flex items-center justify-between p-4 hover:bg-slate-800/40">
            <div className="flex items-center gap-3">
              <input
                type="checkbox"
                checked={s.selected}
                onChange={() => toggleSelect(s.name)}
                className="rounded border-slate-700 bg-slate-800 text-cyan-500 focus:ring-cyan-500"
              />
              <FileSpreadsheet className="h-5 w-5 text-emerald-400" />
              <div>
                <span className="text-sm font-semibold text-white">{s.name}</span>
                <span className="text-xs text-slate-400 font-mono ml-3">{s.rowCount} rows</span>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => removeSheet(s.name)}
                className="p-1.5 text-slate-500 hover:text-rose-400 transition"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default SmartSheetManager;
