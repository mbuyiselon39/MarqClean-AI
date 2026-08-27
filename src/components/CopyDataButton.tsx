import React, { useState } from "react";
import { Copy, Check, FileJson, FileSpreadsheet, ClipboardCopy } from "lucide-react";

export type DataPayload =
  | { type: "records"; data: Record<string, any>[] }
  | { type: "matrix"; data: string[][] }
  | { type: "dataTable"; headers: string[]; rows: string[][] }
  | { type: "text"; text: string };

interface CopyDataButtonProps {
  payload: DataPayload;
  className?: string;
  buttonTheme?: "dark" | "cyan" | "light" | "outline";
  showJson?: boolean;
  showTsv?: boolean; // Tab-separated for instant Excel pasting
  showCsv?: boolean;
  size?: "xs" | "sm" | "md";
  titlePrefix?: string;
}

export function formatDataAsCsv(payload: DataPayload): string {
  if (payload.type === "text") return payload.text;

  if (payload.type === "records") {
    if (!payload.data.length) return "";
    const headers = Object.keys(payload.data[0]);
    const escape = (val: any) => {
      const s = String(val ?? "");
      return s.includes(",") || s.includes('"') || s.includes("\n")
        ? `"${s.replace(/"/g, '""')}"`
        : s;
    };
    const headerLine = headers.map(escape).join(",");
    const rowLines = payload.data.map((row) =>
      headers.map((h) => escape(row[h])).join(",")
    );
    return [headerLine, ...rowLines].join("\n");
  }

  if (payload.type === "matrix") {
    const escape = (val: any) => {
      const s = String(val ?? "");
      return s.includes(",") || s.includes('"') || s.includes("\n")
        ? `"${s.replace(/"/g, '""')}"`
        : s;
    };
    return payload.data.map((row) => row.map(escape).join(",")).join("\n");
  }

  if (payload.type === "dataTable") {
    const escape = (val: any) => {
      const s = String(val ?? "");
      return s.includes(",") || s.includes('"') || s.includes("\n")
        ? `"${s.replace(/"/g, '""')}"`
        : s;
    };
    const headerLine = payload.headers.map(escape).join(",");
    const rowLines = payload.rows.map((row) => row.map(escape).join(","));
    return [headerLine, ...rowLines].join("\n");
  }

  return "";
}

export function formatDataAsTsv(payload: DataPayload): string {
  if (payload.type === "text") return payload.text;

  if (payload.type === "records") {
    if (!payload.data.length) return "";
    const headers = Object.keys(payload.data[0]);
    const headerLine = headers.join("\t");
    const rowLines = payload.data.map((row) =>
      headers.map((h) => String(row[h] ?? "").replace(/\t/g, " ")).join("\t")
    );
    return [headerLine, ...rowLines].join("\n");
  }

  if (payload.type === "matrix") {
    return payload.data
      .map((row) => row.map((c) => String(c ?? "").replace(/\t/g, " ")).join("\t"))
      .join("\n");
  }

  if (payload.type === "dataTable") {
    const headerLine = payload.headers.join("\t");
    const rowLines = payload.rows.map((row) =>
      row.map((c) => String(c ?? "").replace(/\t/g, " ")).join("\t")
    );
    return [headerLine, ...rowLines].join("\n");
  }

  return "";
}

export function formatDataAsJson(payload: DataPayload): string {
  if (payload.type === "text") return payload.text;

  if (payload.type === "records") {
    return JSON.stringify(payload.data, null, 2);
  }

  if (payload.type === "matrix") {
    if (!payload.data.length) return "[]";
    const headers = payload.data[0];
    const rows = payload.data.slice(1);
    const objects = rows.map((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((h, i) => {
        obj[h || `Col_${i + 1}`] = row[i] ?? "";
      });
      return obj;
    });
    return JSON.stringify(objects, null, 2);
  }

  if (payload.type === "dataTable") {
    const objects = payload.rows.map((row) => {
      const obj: Record<string, string> = {};
      payload.headers.forEach((h, i) => {
        obj[h || `Col_${i + 1}`] = row[i] ?? "";
      });
      return obj;
    });
    return JSON.stringify(objects, null, 2);
  }

  return "";
}

export const CopyDataButton: React.FC<CopyDataButtonProps> = ({
  payload,
  className = "",
  buttonTheme = "dark",
  showJson = true,
  showTsv = true,
  showCsv = true,
  size = "xs",
  titlePrefix = "",
}) => {
  const [copiedFormat, setCopiedFormat] = useState<string | null>(null);

  const copyToClipboard = (format: "csv" | "tsv" | "json" | "text") => {
    let content = "";
    if (format === "csv") content = formatDataAsCsv(payload);
    else if (format === "tsv") content = formatDataAsTsv(payload);
    else if (format === "json") content = formatDataAsJson(payload);
    else content = payload.type === "text" ? payload.text : formatDataAsCsv(payload);

    navigator.clipboard.writeText(content);
    setCopiedFormat(format);
    setTimeout(() => {
      setCopiedFormat(null);
    }, 2000);
  };

  // Size styling
  const sizeClasses =
    size === "xs"
      ? "px-2.5 py-1 text-[11px] gap-1.5"
      : size === "sm"
      ? "px-3 py-1.5 text-xs gap-1.5"
      : "px-4 py-2 text-sm gap-2";

  // Theme styling
  const getThemeClasses = (isActive: boolean) => {
    if (isActive) {
      return "border-emerald-500/50 bg-emerald-950/80 text-emerald-300 font-semibold";
    }

    if (buttonTheme === "cyan") {
      return "border-cyan-500/30 bg-cyan-950/40 text-cyan-200 hover:bg-cyan-900/60 hover:text-white";
    }
    if (buttonTheme === "light") {
      return "border-slate-300 bg-white text-slate-700 hover:bg-slate-50 hover:text-slate-950 hover:border-slate-400";
    }
    if (buttonTheme === "outline") {
      return "border-slate-700 bg-transparent text-slate-300 hover:border-slate-500 hover:text-white";
    }
    // Default "dark"
    return "border-slate-800 bg-slate-900/80 text-slate-300 hover:border-slate-700 hover:bg-slate-800 hover:text-white";
  };

  if (payload.type === "text") {
    const isCopied = copiedFormat === "text";
    return (
      <button
        type="button"
        onClick={() => copyToClipboard("text")}
        className={`inline-flex items-center rounded-xl border font-medium transition shadow-sm ${sizeClasses} ${getThemeClasses(
          isCopied
        )} ${className}`}
        title={`${titlePrefix}Copy text to clipboard`}
      >
        {isCopied ? (
          <>
            <Check className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
            <span className="font-bold">Copied!</span>
          </>
        ) : (
          <>
            <Copy className="h-3.5 w-3.5 text-cyan-400 shrink-0" />
            <span>{titlePrefix ? `${titlePrefix} ` : ""}Copy to Clipboard</span>
          </>
        )}
      </button>
    );
  }

  return (
    <div className={`inline-flex items-center flex-wrap gap-1.5 ${className}`}>
      {showCsv && (
        <button
          type="button"
          onClick={() => copyToClipboard("csv")}
          className={`inline-flex items-center rounded-xl border font-medium transition shadow-sm ${sizeClasses} ${getThemeClasses(
            copiedFormat === "csv"
          )}`}
          title="Copy output formatted as Comma-Separated Values (CSV)"
        >
          {copiedFormat === "csv" ? (
            <>
              <Check className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
              <span className="font-bold">Copied CSV!</span>
            </>
          ) : (
            <>
              <ClipboardCopy className="h-3.5 w-3.5 text-cyan-400 shrink-0" />
              <span>Copy CSV</span>
            </>
          )}
        </button>
      )}

      {showTsv && (
        <button
          type="button"
          onClick={() => copyToClipboard("tsv")}
          className={`inline-flex items-center rounded-xl border font-medium transition shadow-sm ${sizeClasses} ${getThemeClasses(
            copiedFormat === "tsv"
          )}`}
          title="Copy as Tab-Separated Values (ready to paste directly into Excel or Google Sheets)"
        >
          {copiedFormat === "tsv" ? (
            <>
              <Check className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
              <span className="font-bold">Copied for Excel!</span>
            </>
          ) : (
            <>
              <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
              <span>Copy for Excel</span>
            </>
          )}
        </button>
      )}

      {showJson && (
        <button
          type="button"
          onClick={() => copyToClipboard("json")}
          className={`inline-flex items-center rounded-xl border font-medium transition shadow-sm ${sizeClasses} ${getThemeClasses(
            copiedFormat === "json"
          )}`}
          title="Copy output structured as JSON Array"
        >
          {copiedFormat === "json" ? (
            <>
              <Check className="h-3.5 w-3.5 text-emerald-400 shrink-0" />
              <span className="font-bold">Copied JSON!</span>
            </>
          ) : (
            <>
              <FileJson className="h-3.5 w-3.5 text-amber-400 shrink-0" />
              <span>Copy JSON</span>
            </>
          )}
        </button>
      )}
    </div>
  );
};
