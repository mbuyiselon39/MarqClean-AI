import React from "react";
import { Sparkles, ArrowLeft } from "lucide-react";

export interface GlobalHeaderProps {
  title?: string;
  subtitle?: string;
  onExit?: () => void;
}

export const GlobalHeader: React.FC<GlobalHeaderProps> = ({
  title = "Reconciliation Hub",
  subtitle = "High-throughput financial and dataset reconciliation engine",
  onExit,
}) => {
  return (
    <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-slate-800 pb-4">
      <div>
        <div className="flex items-center gap-2">
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-cyan-500/20 text-cyan-400">
            <Sparkles className="h-3.5 w-3.5" />
          </span>
          <h2 className="text-lg font-bold text-white">{title}</h2>
        </div>
        <p className="text-xs text-slate-400 mt-0.5">{subtitle}</p>
      </div>

      {onExit && (
        <button
          onClick={onExit}
          className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-1.5 text-xs font-semibold text-slate-300 hover:text-white transition"
        >
          <ArrowLeft className="h-3.5 w-3.5" />
          <span>Exit Module</span>
        </button>
      )}
    </div>
  );
};

export interface ToolLaunchpadProps {
  tools?: Array<{ key: any; label: string; blurb: string; icon: string }>;
  active?: any;
  onSelect?: (nav: any) => void;
  items?: Array<{ key: any; label: string; blurb: string; icon: string }>;
  activeNav?: any;
  onSelectNav?: (nav: any) => void;
}

export const ToolLaunchpad: React.FC<ToolLaunchpadProps> = ({
  tools,
  active,
  onSelect,
  items,
  activeNav,
  onSelectNav,
}) => {
  const navItems = tools || items || [];
  const currentActive = active !== undefined ? active : activeNav;
  const handleSelect = onSelect || onSelectNav || (() => {});

  return (
    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2">
      {navItems.map((item) => {
        const isActive = currentActive === item.key;
        return (
          <button
            key={item.key}
            onClick={() => handleSelect(item.key)}
            className={`flex flex-col items-start p-3 rounded-2xl border text-left transition ${
              isActive
                ? "border-cyan-400 bg-cyan-950/40 shadow-[0_0_15px_rgba(0,210,255,0.2)]"
                : "border-slate-800 bg-slate-900/60 hover:border-slate-700 hover:bg-slate-900"
            }`}
          >
            <span className="text-lg">{item.icon}</span>
            <span className={`mt-1.5 text-xs font-bold ${isActive ? "text-cyan-300" : "text-slate-200"}`}>
              {item.label}
            </span>
            <span className="mt-0.5 text-[10px] text-slate-500 line-clamp-2">{item.blurb}</span>
          </button>
        );
      })}
    </div>
  );
};
