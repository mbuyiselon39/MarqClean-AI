import React, { useState, useRef } from "react";
import {
  Upload,
  Sparkles,
  CheckCircle2,
  AlertCircle,
  Clock,
  Download,
  FileSpreadsheet,
  FileText,
  Trash2,
  Eye,
  RefreshCw,
  Play,
  Layers,
  ArrowRight,
  ShieldCheck,
  Search,
  Database,
  Plus,
  Archive,
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import * as fflate from "fflate";
import { CopyDataButton } from "./CopyDataButton";

export interface BatchFileItem {
  id: string;
  file?: File;
  name: string;
  sizeBytes: number;
  format: "csv" | "xlsx" | "tsv" | "txt";
  status: "queued" | "cleaning" | "completed" | "error";
  errorMessage?: string;
  rawRowCount: number;
  rawColCount: number;
  cleanedRows: Record<string, string>[];
  headers: string[];
  stats: {
    rawRows: number;
    cleanedRows: number;
    duplicatesPurged: number;
    fieldsSanitized: number;
    headersFixed: number;
    processingTimeMs: number;
  };
}

interface BulkDataCleanerProps {
  onRequestAudit?: () => void;
}

// Built-in realistic enterprise sample datasets
const SAMPLE_BATCH: Array<{ name: string; rawCsv: string }> = [
  {
    name: "Q3_Enterprise_Inbound_Leads.csv",
    rawCsv: `Lead ID, full_name , Email Address , Company Name , Deal size USD , Region , Status
LD-901,  alexander hamilton  , ALEX.H@TREASURY.GOV , Treasury FinTech , $ 125,000.50 , North America , NEW
LD-902, SARAH CONNOR , sarah.c@cyberdyne.io , Cyberdyne Systems , 85400.00 , West Coast , CONTACTED
LD-903,  elena rostova , elena@novatech.eu , NovaTech Solutions , € 45.200,00 , Europe , QUALIFIED
LD-901,  alexander hamilton  , ALEX.H@TREASURY.GOV , Treasury FinTech , $ 125,000.50 , North America , NEW
LD-904, michael c. chang , MCHANG@APEX-LOGISTICS.COM , Apex Logistics , (34,500.00) , APAC , PENDING
LD-905,  MARIA GARCIA-LOPEZ  , maria.garcia@solarpower.es , Solaria Energía , $ 210,000.00 , EMEA , CLOSED_WON
LD-906, david miller , dmiller@acme.org , Acme Corp , 14500 , LATAM , DISQUALIFIED
LD-902, SARAH CONNOR , sarah.c@cyberdyne.io , Cyberdyne Systems , 85400.00 , West Coast , CONTACTED`,
  },
  {
    name: "August_Vendor_Disbursements.csv",
    rawCsv: `Transaction Ref #, Vendor Full Legal Name , Invoice Total , Currency Code , Payment Status , Routing / IBAN
TXN-8801, AMAZON WEB SERVICES INC. , $ 14,892.45 , USD , PAID , US64SVBK9902
TXN-8802, snowflake computing pte ltd , $ 8,230.00 , usd , PENDING_APPROVAL , US33CHAS1102
TXN-8803, DATADOG IRELAND LIMITED , € 3.450,80 , EUR , PROCESSED , IE44BOFI9021
TXN-8804,  GOOGLE CLOUD EMEA LTD  , $ 19,450.10 , USD , PAID , IE99AIBK4401
TXN-8801, AMAZON WEB SERVICES INC. , $ 14,892.45 , USD , PAID , US64SVBK9902
TXN-8805, Microsoft Operations Ireland , € 11.200,00 , eur , SCHEDULED , IE22IPBS8810
TXN-8806,  salesforce.com emea  , (4,250.00) , USD , CREDIT_NOTE , US11WELL7732`,
  },
  {
    name: "Global_HR_Staff_Directory.csv",
    rawCsv: `Employee ID , FIRST & LAST NAME , corporate_email , Department Unit , Employment Grade , Base Salary
EMP-101,  ALAN TURING  , alan.turing@bletchley.ac.uk , Cryptanalysis , Principal Fellow , 145000
EMP-102, ADA LOVELACE , ada@analytical-engine.org , Computer Science , Chief Architect , 185000
EMP-103,  grace hopper , ghopper@navy-comp.mil , Systems Architecture , Rear Admiral , 165000
EMP-104, claude shannon , SHANNON@BELL-LABS.COM , Information Theory , Senior Director , 172000
EMP-102, ADA LOVELACE , ada@analytical-engine.org , Computer Science , Chief Architect , 185000
EMP-105,  katherine johnson  , k.johnson@nasa-space.gov , Orbital Dynamics , Lead Mathematician , 158000
EMP-106, JOHN VON NEUMANN , jvn@ias-princeton.edu , Quantum Computation , Professor Chair , 195000`,
  },
];

// Helper to format bytes
function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

// Core cleaning algorithm for raw row records
function cleanRowData(rawRows: Record<string, any>[]): {
  cleanedRows: Record<string, string>[];
  headers: string[];
  stats: {
    rawRows: number;
    cleanedRows: number;
    duplicatesPurged: number;
    fieldsSanitized: number;
    headersFixed: number;
  };
} {
  if (!rawRows || rawRows.length === 0) {
    return {
      cleanedRows: [],
      headers: [],
      stats: { rawRows: 0, cleanedRows: 0, duplicatesPurged: 0, fieldsSanitized: 0, headersFixed: 0 },
    };
  }

  // 1. Sanitize column headers
  const rawHeaders = Object.keys(rawRows[0] || {});
  const headerMap: Record<string, string> = {};
  let headersFixed = 0;

  rawHeaders.forEach((origHeader) => {
    // Strip leading/trailing spaces and internal excessive underscores/special characters
    let cleanHeader = origHeader
      .trim()
      .replace(/[#&$/()\\.,]/g, "")
      .replace(/[_-\s]+/g, " ")
      .trim();

    // Standardize to Proper Title Case
    cleanHeader = cleanHeader.replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase());

    if (!cleanHeader) cleanHeader = "Column";
    if (cleanHeader !== origHeader) {
      headersFixed++;
    }
    headerMap[origHeader] = cleanHeader;
  });

  const finalHeaders = Array.from(new Set(Object.values(headerMap)));

  // 2. Normalize and Deduplicate rows
  const seenRowSignatures = new Set<string>();
  const cleanedRows: Record<string, string>[] = [];
  let duplicatesPurged = 0;
  let fieldsSanitized = 0;

  rawRows.forEach((row) => {
    const cleanedRow: Record<string, string> = {};
    let hasContent = false;

    Object.entries(row).forEach(([origKey, val]) => {
      const targetKey = headerMap[origKey] || origKey;
      let cellStr = String(val ?? "").trim();

      if (cellStr.length > 0) {
        hasContent = true;
      }

      // Value Transformations
      if (cellStr) {
        const origCell = cellStr;

        // Check if Email -> lower case and clean
        if (cellStr.includes("@") && !cellStr.includes(" ")) {
          cellStr = cellStr.toLowerCase();
        }
        // Check if Currency / Accounting Bracket -> format standard number
        else if (cellStr.startsWith("(") && cellStr.endsWith(")")) {
          const inner = cellStr.slice(1, -1).replace(/[^0-9.,]/g, "");
          const num = parseFloat(inner.replace(/,/g, ""));
          if (!isNaN(num)) {
            cellStr = `-${num.toFixed(2)}`;
          }
        } else if (/^[$€£¥]\s*[\d,.]+/.test(cellStr) || /[\d,.]+\s*[$€£¥]/.test(cellStr)) {
          let numStr = cellStr.replace(/[^0-9.,-]/g, "");
          // Handle European 1.234,56
          if (numStr.includes(",") && numStr.includes(".")) {
            if (numStr.indexOf(".") < numStr.indexOf(",")) {
              numStr = numStr.replace(/\./g, "").replace(",", ".");
            } else {
              numStr = numStr.replace(/,/g, "");
            }
          }
          const num = parseFloat(numStr);
          if (!isNaN(num)) {
            cellStr = num.toFixed(2);
          }
        }
        // General text casing (if string name, title case)
        else if (isNaN(Number(cellStr)) && !cellStr.startsWith("http") && !cellStr.startsWith("TXN-") && !cellStr.startsWith("EMP-") && !cellStr.startsWith("LD-") && !cellStr.includes("_")) {
          // If all caps or all lowercase words, title-case them
          if (cellStr === cellStr.toUpperCase() || cellStr === cellStr.toLowerCase()) {
            cellStr = cellStr.replace(/\b\w+/g, (w) => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase());
          }
        }

        if (cellStr !== origCell) {
          fieldsSanitized++;
        }
      }

      cleanedRow[targetKey] = cellStr;
    });

    if (!hasContent) return; // Skip totally blank rows

    // Deduplication check
    const signature = JSON.stringify(cleanedRow);
    if (seenRowSignatures.has(signature)) {
      duplicatesPurged++;
      return;
    }
    seenRowSignatures.add(signature);
    cleanedRows.push(cleanedRow);
  });

  return {
    cleanedRows,
    headers: finalHeaders,
    stats: {
      rawRows: rawRows.length,
      cleanedRows: cleanedRows.length,
      duplicatesPurged,
      fieldsSanitized,
      headersFixed,
    },
  };
}

export const BulkDataCleaner: React.FC<BulkDataCleanerProps> = ({ onRequestAudit }) => {
  const [fileQueue, setFileQueue] = useState<BatchFileItem[]>([]);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isBatchProcessing, setIsBatchProcessing] = useState(false);
  const [inspectedFileId, setInspectedFileId] = useState<string | null>(null);
  const [searchFilter, setSearchFilter] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Ingest raw File or text data into queue item
  const processRawFileToQueue = async (file: File): Promise<BatchFileItem> => {
    const ext = file.name.split(".").pop()?.toLowerCase() || "csv";
    const format = ext === "xlsx" || ext === "xls" ? "xlsx" : ext === "tsv" ? "tsv" : "csv";

    return new Promise((resolve) => {
      if (format === "xlsx") {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rawJson = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet);

            const headers = rawJson.length > 0 ? Object.keys(rawJson[0]) : [];
            resolve({
              id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
              file,
              name: file.name,
              sizeBytes: file.size,
              format: "xlsx",
              status: "queued",
              rawRowCount: rawJson.length,
              rawColCount: headers.length,
              cleanedRows: [],
              headers,
              stats: {
                rawRows: rawJson.length,
                cleanedRows: 0,
                duplicatesPurged: 0,
                fieldsSanitized: 0,
                headersFixed: 0,
                processingTimeMs: 0,
              },
            });
          } catch (err: any) {
            resolve({
              id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
              file,
              name: file.name,
              sizeBytes: file.size,
              format: "xlsx",
              status: "error",
              errorMessage: err.message || "Failed to read Excel workbook.",
              rawRowCount: 0,
              rawColCount: 0,
              cleanedRows: [],
              headers: [],
              stats: { rawRows: 0, cleanedRows: 0, duplicatesPurged: 0, fieldsSanitized: 0, headersFixed: 0, processingTimeMs: 0 },
            });
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        // CSV / TSV / TXT
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            const rawData = results.data as Record<string, any>[];
            const headers = rawData.length > 0 ? Object.keys(rawData[0]) : [];
            resolve({
              id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
              file,
              name: file.name,
              sizeBytes: file.size,
              format: format as any,
              status: "queued",
              rawRowCount: rawData.length,
              rawColCount: headers.length,
              cleanedRows: [],
              headers,
              stats: {
                rawRows: rawData.length,
                cleanedRows: 0,
                duplicatesPurged: 0,
                fieldsSanitized: 0,
                headersFixed: 0,
                processingTimeMs: 0,
              },
            });
          },
          error: (err) => {
            resolve({
              id: `file_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
              file,
              name: file.name,
              sizeBytes: file.size,
              format: format as any,
              status: "error",
              errorMessage: err.message || "Failed to parse CSV file.",
              rawRowCount: 0,
              rawColCount: 0,
              cleanedRows: [],
              headers: [],
              stats: { rawRows: 0, cleanedRows: 0, duplicatesPurged: 0, fieldsSanitized: 0, headersFixed: 0, processingTimeMs: 0 },
            });
          },
        });
      }
    });
  };

  // Add files to queue
  const handleAddFiles = async (files: FileList | File[]) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const newItems: BatchFileItem[] = [];
    for (const f of fileArray) {
      const item = await processRawFileToQueue(f);
      newItems.push(item);
    }

    setFileQueue((prev) => [...prev, ...newItems]);
  };

  // Load Built-in Demo Sample Batch
  const handleLoadSampleBatch = () => {
    const sampleItems: BatchFileItem[] = SAMPLE_BATCH.map((sample, idx) => {
      const parsed = Papa.parse(sample.rawCsv.trim(), { header: true, skipEmptyLines: true });
      const rawData = parsed.data as Record<string, any>[];
      const headers = rawData.length > 0 ? Object.keys(rawData[0]) : [];

      return {
        id: `sample_${idx}_${Date.now()}`,
        name: sample.name,
        sizeBytes: sample.rawCsv.length,
        format: "csv",
        status: "queued",
        rawRowCount: rawData.length,
        rawColCount: headers.length,
        cleanedRows: [],
        headers,
        stats: {
          rawRows: rawData.length,
          cleanedRows: 0,
          duplicatesPurged: 0,
          fieldsSanitized: 0,
          headersFixed: 0,
          processingTimeMs: 0,
        },
      };
    });

    setFileQueue(sampleItems);
  };

  // Drag & Drop Handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleAddFiles(e.dataTransfer.files);
    }
  };

  // Clean a single item from the queue
  const cleanSingleItem = async (item: BatchFileItem): Promise<BatchFileItem> => {
    const startTime = performance.now();
    return new Promise((resolve) => {
      // If sample file without real File object, parse from sample text
      if (!item.file) {
        const matchingSample = SAMPLE_BATCH.find((s) => s.name === item.name);
        const csvText = matchingSample ? matchingSample.rawCsv : "";
        const parsed = Papa.parse(csvText.trim(), { header: true, skipEmptyLines: true });
        const res = cleanRowData(parsed.data as Record<string, any>[]);
        const elapsed = Math.round(performance.now() - startTime);

        resolve({
          ...item,
          status: "completed",
          cleanedRows: res.cleanedRows,
          headers: res.headers,
          rawColCount: res.headers.length,
          stats: {
            ...res.stats,
            processingTimeMs: elapsed,
          },
        });
        return;
      }

      // Real File parsing
      if (item.format === "xlsx") {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawJson = XLSX.utils.sheet_to_json<Record<string, any>>(sheet);
            const res = cleanRowData(rawJson);
            const elapsed = Math.round(performance.now() - startTime);

            resolve({
              ...item,
              status: "completed",
              cleanedRows: res.cleanedRows,
              headers: res.headers,
              rawColCount: res.headers.length,
              stats: {
                ...res.stats,
                processingTimeMs: elapsed,
              },
            });
          } catch (err: any) {
            resolve({
              ...item,
              status: "error",
              errorMessage: err.message || "Failed to process sheet.",
            });
          }
        };
        reader.readAsArrayBuffer(item.file);
      } else {
        Papa.parse(item.file, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            const res = cleanRowData(results.data as Record<string, any>[]);
            const elapsed = Math.round(performance.now() - startTime);
            resolve({
              ...item,
              status: "completed",
              cleanedRows: res.cleanedRows,
              headers: res.headers,
              rawColCount: res.headers.length,
              stats: {
                ...res.stats,
                processingTimeMs: elapsed,
              },
            });
          },
          error: (err) => {
            resolve({
              ...item,
              status: "error",
              errorMessage: err.message || "CSV parse error.",
            });
          },
        });
      }
    });
  };

  // Trigger Cleanup Sequence for ALL files simultaneously
  const handleExecuteBatchCleanup = async () => {
    if (fileQueue.length === 0 || isBatchProcessing) return;

    setIsBatchProcessing(true);

    // Set all queued items to cleaning state
    setFileQueue((prev) =>
      prev.map((item) => (item.status !== "completed" ? { ...item, status: "cleaning" } : item))
    );

    // Small delay to ensure smooth UI animation transition
    await new Promise((r) => setTimeout(r, 250));

    // Process all files in parallel
    const updated = await Promise.all(
      fileQueue.map(async (item) => {
        if (item.status === "completed" && item.cleanedRows.length > 0) {
          return item;
        }
        return await cleanSingleItem(item);
      })
    );

    setFileQueue(updated);
    setIsBatchProcessing(false);

    // If no file inspected yet, open inspector for first completed file
    const firstCompleted = updated.find((u) => u.status === "completed");
    if (firstCompleted && !inspectedFileId) {
      setInspectedFileId(firstCompleted.id);
    }
  };

  // Remove individual file from queue
  const handleRemoveFile = (id: string, e?: React.MouseEvent) => {
    e?.stopPropagation();
    setFileQueue((prev) => prev.filter((item) => item.id !== id));
    if (inspectedFileId === id) {
      setInspectedFileId(null);
    }
  };

  // Clear entire queue
  const handleClearAll = () => {
    setFileQueue([]);
    setInspectedFileId(null);
  };

  // Master Combined Excel Export (All files as separate sheets)
  const handleExportCombinedExcel = () => {
    const completedFiles = fileQueue.filter((f) => f.status === "completed" && f.cleanedRows.length > 0);
    if (completedFiles.length === 0) return;

    const wb = XLSX.utils.book_new();

    completedFiles.forEach((f, idx) => {
      const ws = XLSX.utils.json_to_sheet(f.cleanedRows);
      // Clean sheet name (max 31 chars, no invalid chars)
      let sheetName = f.name.replace(/\.(csv|xlsx|tsv|txt)$/i, "").replace(/[\\/*?[\]:]/g, "");
      sheetName = sheetName.substring(0, 28) || `Sheet_${idx + 1}`;
      if (wb.SheetNames.includes(sheetName)) {
        sheetName = `${sheetName.substring(0, 25)}_${idx + 1}`;
      }
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, `MarqClean_Sanitized_Batch_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Master ZIP Export of all cleaned CSVs or Excel sheets
  const handleExportZipArchive = () => {
    const completedFiles = fileQueue.filter((f) => f.status === "completed" && f.cleanedRows.length > 0);
    if (completedFiles.length === 0) return;

    const zipEntries: Record<string, Uint8Array> = {};

    completedFiles.forEach((f) => {
      const csvString = Papa.unparse(f.cleanedRows);
      const cleanName = f.name.replace(/\.(csv|xlsx|tsv|txt)$/i, "_Cleaned.csv");
      zipEntries[cleanName] = fflate.strToU8(csvString);
    });

    // Add a summary report manifest text inside zip
    const summaryText = `MARQCLEAN AI BATCH SANITATION REPORT
Generated: ${new Date().toISOString()}
Total Files Processed: ${completedFiles.length}
Total Clean Rows: ${completedFiles.reduce((acc, f) => acc + f.stats.cleanedRows, 0)}
Total Duplicates Purged: ${completedFiles.reduce((acc, f) => acc + f.stats.duplicatesPurged, 0)}
Total Fields Normalized: ${completedFiles.reduce((acc, f) => acc + f.stats.fieldsSanitized, 0)}
==================================================
Files List:
${completedFiles.map((f, i) => `${i + 1}. ${f.name} -> ${f.stats.cleanedRows} clean records (${f.stats.duplicatesPurged} duplicates purged)`).join("\n")}
`;
    zipEntries["BATCH_SUMMARY_MANIFEST.txt"] = fflate.strToU8(summaryText);

    const zipped = fflate.zipSync(zipEntries);
    const blob = new Blob([zipped], { type: "application/zip" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `MarqClean_Sanitized_Batch_${Date.now()}.zip`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Download single cleaned file
  const handleDownloadSingleFile = (item: BatchFileItem, format: "xlsx" | "csv", e?: React.MouseEvent) => {
    e?.stopPropagation();
    if (item.cleanedRows.length === 0) return;

    if (format === "xlsx") {
      const ws = XLSX.utils.json_to_sheet(item.cleanedRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sanitized Data");
      XLSX.writeFile(wb, `${item.name.replace(/\.[^/.]+$/, "")}_Cleaned.xlsx`);
    } else {
      const csv = Papa.unparse(item.cleanedRows);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${item.name.replace(/\.[^/.]+$/, "")}_Cleaned.csv`;
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  // Aggregated Batch Metrics
  const totalQueued = fileQueue.length;
  const totalCompleted = fileQueue.filter((f) => f.status === "completed").length;
  const totalInputRows = fileQueue.reduce((sum, f) => sum + (f.rawRowCount || 0), 0);
  const totalCleanRows = fileQueue.reduce((sum, f) => sum + (f.stats?.cleanedRows || 0), 0);
  const totalDupesPurged = fileQueue.reduce((sum, f) => sum + (f.stats?.duplicatesPurged || 0), 0);
  const totalNormalizations = fileQueue.reduce((sum, f) => sum + (f.stats?.fieldsSanitized || 0) + (f.stats?.headersFixed || 0), 0);

  // Inspected File
  const currentInspected = fileQueue.find((f) => f.id === inspectedFileId) || null;

  // Filtered rows for inspected table preview
  const inspectedDisplayRows = currentInspected
    ? currentInspected.cleanedRows.filter((r) => {
        if (!searchFilter.trim()) return true;
        const q = searchFilter.toLowerCase();
        return Object.values(r).some((v) => String(v).toLowerCase().includes(q));
      })
    : [];

  return (
    <div className="space-y-8">
      {/* 1. Header Banner & Security Guarantee */}
      <div className="rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-6 sm:p-8 backdrop-blur-xl shadow-2xl">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-slate-800 pb-6">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-bold uppercase tracking-wider text-cyan-300">
              <Layers className="h-3.5 w-3.5" />
              Simultaneous Multi-File Processing
            </div>
            <h2 className="mt-2 text-2xl font-extrabold text-white sm:text-3xl">
              Bulk Data &amp; Multi-CSV Batch Cleanroom
            </h2>
            <p className="mt-1.5 text-xs sm:text-sm text-slate-400 max-w-2xl">
              Drop dozens of messy CSV or Excel spreadsheets at once. Standardize headers, eliminate duplicates, format accounting currencies, and export consolidated multi-sheet workbooks in one click.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <button
              onClick={handleLoadSampleBatch}
              className="flex items-center gap-2 rounded-xl border border-cyan-500/40 bg-cyan-950/40 px-4 py-2.5 text-xs font-bold text-cyan-300 hover:bg-cyan-900/50 transition shadow-sm shadow-cyan-500/10"
            >
              <Sparkles className="h-4 w-4" />
              <span>Load 3 Sample Datasets</span>
            </button>

            {onRequestAudit && (
              <button
                onClick={onRequestAudit}
                className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 px-4 py-2.5 text-xs font-bold text-slate-950 hover:brightness-110 transition shadow-lg shadow-cyan-500/20"
              >
                <span>Enterprise Audit</span>
                <ArrowRight className="h-3.5 w-3.5" />
              </button>
            )}
          </div>
        </div>

        {/* 2. Drag & Drop Upload Zone (Multi-file enabled) */}
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`mt-6 relative cursor-pointer rounded-2xl border-2 border-dashed p-8 sm:p-10 text-center transition-all ${
            isDragOver
              ? "border-cyan-400 bg-cyan-950/40 scale-[1.01] shadow-[0_0_30px_rgba(0,210,255,0.3)]"
              : "border-cyan-500/30 bg-slate-900/40 hover:border-cyan-400 hover:bg-slate-900/70"
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            multiple
            accept=".csv,.xlsx,.xls,.tsv,.txt"
            className="hidden"
            onChange={(e) => {
              if (e.target.files) handleAddFiles(e.target.files);
              // Reset input so re-selecting same files triggers change
              e.target.value = "";
            }}
          />

          <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-cyan-500/20 to-blue-600/20 border border-cyan-500/40 text-cyan-300 shadow-inner">
            <Upload className="h-7 w-7 animate-pulse" />
          </div>

          <h3 className="mt-4 text-base font-bold text-white">
            {isDragOver ? "Release to Queue All Files Instantly" : "Drop Multiple CSV or Excel Files Here"}
          </h3>
          <p className="mt-1.5 text-xs text-slate-400 max-w-md mx-auto">
            Drag &amp; drop multiple files at once or click to browse. Supports <span className="text-cyan-300 font-mono font-semibold">.CSV</span>, <span className="text-cyan-300 font-mono font-semibold">.XLSX</span>, and <span className="text-cyan-300 font-mono font-semibold">.TSV</span> tables.
          </p>

          <div className="mt-4 flex flex-wrap items-center justify-center gap-4 text-[11px] font-mono text-emerald-400">
            <div className="flex items-center gap-1.5 bg-emerald-950/40 border border-emerald-500/30 px-3 py-1 rounded-full">
              <ShieldCheck className="h-3.5 w-3.5" />
              <span>100% In-Memory Processing &bull; Zero Server Upload</span>
            </div>
            <div className="flex items-center gap-1.5 text-slate-400">
              <Database className="h-3.5 w-3.5 text-cyan-400" />
              <span>Auto Key Deduplication &amp; Header Sanitizer</span>
            </div>
          </div>
        </div>

        {/* 3. Aggregated Batch Stats Overview */}
        {fileQueue.length > 0 && (
          <div className="mt-8 space-y-6">
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4 font-mono text-xs">
              <div className="rounded-2xl border border-cyan-500/30 bg-slate-900/80 p-4 shadow-lg">
                <span className="text-[11px] font-sans font-bold uppercase tracking-wider text-slate-400">
                  Batch Queue Status
                </span>
                <div className="mt-1 flex items-baseline justify-between">
                  <span className="text-2xl font-black text-cyan-400">
                    {totalCompleted} / {totalQueued}
                  </span>
                  <span className="text-xs text-slate-400">Files Cleaned</span>
                </div>
                <div className="mt-2 h-1.5 w-full overflow-hidden rounded-full bg-slate-800">
                  <div
                    className="h-full bg-cyan-400 transition-all duration-500"
                    style={{ width: `${totalQueued > 0 ? (totalCompleted / totalQueued) * 100 : 0}%` }}
                  />
                </div>
              </div>

              <div className="rounded-2xl border border-slate-800 bg-slate-900/80 p-4 shadow-lg">
                <span className="text-[11px] font-sans font-bold uppercase tracking-wider text-slate-400">
                  Clean Records Output
                </span>
                <div className="mt-1 flex items-baseline justify-between">
                  <span className="text-2xl font-black text-emerald-400">
                    {totalCleanRows.toLocaleString()}
                  </span>
                  <span className="text-[11px] text-slate-400">from {totalInputRows.toLocaleString()} raw</span>
                </div>
                <p className="mt-1 text-[10px] text-slate-500">100% normalized &amp; typed</p>
              </div>

              <div className="rounded-2xl border border-slate-800 bg-slate-900/80 p-4 shadow-lg">
                <span className="text-[11px] font-sans font-bold uppercase tracking-wider text-slate-400">
                  Duplicates Purged
                </span>
                <div className="mt-1 flex items-baseline justify-between">
                  <span className="text-2xl font-black text-rose-400">
                    {totalDupesPurged.toLocaleString()}
                  </span>
                  <span className="text-[11px] text-rose-300 font-sans font-semibold">Exact Rows</span>
                </div>
                <p className="mt-1 text-[10px] text-slate-500">Deep fingerprint matching</p>
              </div>

              <div className="rounded-2xl border border-slate-800 bg-slate-900/80 p-4 shadow-lg">
                <span className="text-[11px] font-sans font-bold uppercase tracking-wider text-slate-400">
                  Fields &amp; Headers Sanitized
                </span>
                <div className="mt-1 flex items-baseline justify-between">
                  <span className="text-2xl font-black text-blue-400">
                    {totalNormalizations.toLocaleString()}
                  </span>
                  <span className="text-[11px] text-blue-300 font-sans font-semibold">Corrections</span>
                </div>
                <p className="mt-1 text-[10px] text-slate-500">Names, emails &amp; currencies</p>
              </div>
            </div>

            {/* 4. Global Action & Batch Trigger Toolbar */}
            <div className="flex flex-col gap-4 rounded-2xl border border-slate-800 bg-slate-900/90 p-4 sm:flex-row sm:items-center sm:justify-between shadow-xl">
              <div className="flex flex-wrap items-center gap-3">
                {/* Execute All Trigger */}
                <button
                  onClick={handleExecuteBatchCleanup}
                  disabled={isBatchProcessing || fileQueue.every((f) => f.status === "completed")}
                  className={`flex items-center gap-2 rounded-xl px-5 py-3 text-xs font-extrabold uppercase tracking-wider transition shadow-lg ${
                    fileQueue.every((f) => f.status === "completed")
                      ? "bg-slate-800 text-slate-500 cursor-not-allowed"
                      : "bg-gradient-to-r from-cyan-400 via-cyan-500 to-blue-600 text-slate-950 hover:brightness-110 shadow-cyan-500/25 active:scale-95"
                  }`}
                >
                  {isBatchProcessing ? (
                    <>
                      <RefreshCw className="h-4 w-4 animate-spin text-slate-950" />
                      <span>Cleaning All Files in Parallel...</span>
                    </>
                  ) : fileQueue.every((f) => f.status === "completed") ? (
                    <>
                      <CheckCircle2 className="h-4 w-4 text-emerald-400" />
                      <span>All Files Sanitized</span>
                    </>
                  ) : (
                    <>
                      <Play className="h-4 w-4 fill-slate-950" />
                      <span>Trigger Simultaneous Batch Cleanup ({fileQueue.filter((f) => f.status !== "completed").length})</span>
                    </>
                  )}
                </button>

                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-800/80 px-3.5 py-2.5 text-xs font-semibold text-slate-300 hover:text-white hover:bg-slate-700 transition"
                >
                  <Plus className="h-3.5 w-3.5" />
                  <span>Add More Files</span>
                </button>
              </div>

              {/* Master Batch Export Options */}
              {totalCompleted > 0 && (
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={handleExportCombinedExcel}
                    className="flex items-center gap-2 rounded-xl border border-emerald-500/50 bg-emerald-950/40 px-4 py-2.5 text-xs font-bold text-emerald-300 hover:bg-emerald-900/50 transition shadow-md shadow-emerald-500/10"
                    title="Export all cleaned datasets as tabs in one master Excel workbook"
                  >
                    <FileSpreadsheet className="h-4 w-4 text-emerald-400" />
                    <span>Download Multi-Sheet Excel (.xlsx)</span>
                  </button>

                  <button
                    onClick={handleExportZipArchive}
                    className="flex items-center gap-2 rounded-xl border border-cyan-500/50 bg-cyan-950/40 px-4 py-2.5 text-xs font-bold text-cyan-300 hover:bg-cyan-900/50 transition shadow-md shadow-cyan-500/10"
                    title="Download individual cleaned CSV files and manifest in a ZIP archive"
                  >
                    <Archive className="h-4 w-4 text-cyan-400" />
                    <span>Download Batch ZIP</span>
                  </button>

                  <button
                    onClick={handleClearAll}
                    className="p-2.5 text-slate-400 hover:text-rose-400 transition rounded-xl hover:bg-slate-800"
                    title="Clear file queue"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              )}
            </div>

            {/* 5. Unified File Queue Grid / Table */}
            <div className="overflow-hidden rounded-2xl border border-slate-800 bg-slate-900/60 shadow-2xl">
              <div className="border-b border-slate-800 bg-slate-950/80 px-6 py-3.5 flex items-center justify-between">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-300">
                  Active File Queue ({fileQueue.length} items)
                </span>
                <span className="text-xs text-slate-400 font-mono">
                  {totalCompleted} Cleaned &bull; {totalQueued - totalCompleted} Pending
                </span>
              </div>

              <div className="divide-y divide-slate-800/80">
                {fileQueue.map((item) => (
                  <div
                    key={item.id}
                    onClick={() => {
                      if (item.status === "completed") {
                        setInspectedFileId(item.id);
                      }
                    }}
                    className={`flex flex-col gap-4 p-4 sm:flex-row sm:items-center sm:justify-between transition cursor-pointer ${
                      inspectedFileId === item.id
                        ? "bg-cyan-950/30 border-l-4 border-l-cyan-400"
                        : "hover:bg-slate-850 hover:bg-slate-800/40"
                    }`}
                  >
                    {/* Left: File details */}
                    <div className="flex items-center gap-3.5">
                      <div
                        className={`flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-xl font-bold ${
                          item.format === "xlsx"
                            ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/30"
                            : "bg-cyan-500/20 text-cyan-400 border border-cyan-500/30"
                        }`}
                      >
                        {item.format === "xlsx" ? (
                          <FileSpreadsheet className="h-5 w-5" />
                        ) : (
                          <FileText className="h-5 w-5" />
                        )}
                      </div>

                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-sm font-bold text-white hover:text-cyan-300 transition">
                            {item.name}
                          </span>
                          <span className="rounded-md bg-slate-800 px-2 py-0.5 text-[10px] font-mono text-slate-400 uppercase">
                            {formatFileSize(item.sizeBytes)}
                          </span>
                        </div>
                        <div className="mt-1 flex items-center gap-3 text-xs text-slate-400 font-mono">
                          <span>{item.rawRowCount} raw rows</span>
                          <span>&bull;</span>
                          <span>{item.rawColCount} columns</span>
                          {item.stats?.processingTimeMs > 0 && (
                            <>
                              <span>&bull;</span>
                              <span className="text-cyan-400 font-semibold">{item.stats.processingTimeMs}ms</span>
                            </>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Middle: Status & Performance Metrics */}
                    <div className="flex items-center gap-4">
                      {item.status === "queued" && (
                        <div className="flex items-center gap-1.5 rounded-full border border-amber-500/30 bg-amber-950/40 px-3 py-1 text-xs font-semibold text-amber-300">
                          <Clock className="h-3.5 w-3.5 animate-pulse" />
                          <span>Queued for Cleaning</span>
                        </div>
                      )}

                      {item.status === "cleaning" && (
                        <div className="flex items-center gap-1.5 rounded-full border border-cyan-500/40 bg-cyan-950/60 px-3 py-1 text-xs font-bold text-cyan-300">
                          <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                          <span>Sanitizing Cells...</span>
                        </div>
                      )}

                      {item.status === "completed" && (
                        <div className="flex items-center gap-4 font-mono text-xs">
                          <div className="text-right hidden sm:block">
                            <span className="text-emerald-400 font-bold">{item.stats.cleanedRows} clean</span>
                            <span className="text-slate-500 block text-[10px]">
                              -{item.stats.duplicatesPurged} dupes &bull; {item.stats.fieldsSanitized} fixed
                            </span>
                          </div>
                          <div className="flex items-center gap-1.5 rounded-full border border-emerald-500/40 bg-emerald-950/50 px-3 py-1 text-xs font-bold text-emerald-300">
                            <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400" />
                            <span>Cleaned</span>
                          </div>
                        </div>
                      )}

                      {item.status === "error" && (
                        <div className="flex items-center gap-1.5 rounded-full border border-rose-500/40 bg-rose-950/50 px-3 py-1 text-xs font-bold text-rose-300">
                          <AlertCircle className="h-3.5 w-3.5" />
                          <span>Error: {item.errorMessage}</span>
                        </div>
                      )}

                      {/* Right: Actions */}
                      <div className="flex items-center gap-1.5" onClick={(e) => e.stopPropagation()}>
                        {item.status === "queued" && (
                          <button
                            onClick={async () => {
                              const cleaned = await cleanSingleItem(item);
                              setFileQueue((prev) => prev.map((f) => (f.id === item.id ? cleaned : f)));
                              setInspectedFileId(item.id);
                            }}
                            className="rounded-lg bg-cyan-500 px-3 py-1.5 text-xs font-bold text-slate-950 hover:bg-cyan-400 transition"
                          >
                            Clean
                          </button>
                        )}

                        {item.status === "completed" && (
                          <>
                            <button
                              onClick={() => setInspectedFileId(item.id)}
                              className={`flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-xs font-semibold transition ${
                                inspectedFileId === item.id
                                  ? "bg-cyan-500 text-slate-950 font-bold"
                                  : "border border-slate-700 bg-slate-800 text-slate-300 hover:text-white"
                              }`}
                              title="Inspect cleaned table"
                            >
                              <Eye className="h-3.5 w-3.5" />
                              <span className="hidden md:inline">Inspect</span>
                            </button>

                            <CopyDataButton
                              payload={{ type: "records", data: item.cleanedRows }}
                              showCsv={true}
                              showTsv={true}
                              showJson={true}
                              size="xs"
                              buttonTheme="dark"
                            />

                            <button
                              onClick={(e) => handleDownloadSingleFile(item, "xlsx", e)}
                              className="flex items-center gap-1 rounded-lg border border-slate-700 bg-slate-800 px-2.5 py-1.5 text-xs font-semibold text-slate-300 hover:text-white transition"
                              title="Download as Excel (.xlsx)"
                            >
                              <Download className="h-3.5 w-3.5 text-emerald-400" />
                              <span className="hidden md:inline">Excel</span>
                            </button>
                          </>
                        )}

                        <button
                          onClick={(e) => handleRemoveFile(item.id, e)}
                          className="rounded-lg p-1.5 text-slate-500 hover:text-rose-400 hover:bg-slate-800 transition"
                          title="Remove from queue"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* 6. Interactive Inspector & Clean Data Preview Panel */}
            {currentInspected && currentInspected.status === "completed" && (
              <div className="rounded-3xl border border-cyan-500/30 bg-slate-950/90 p-6 sm:p-8 backdrop-blur-xl shadow-2xl space-y-6">
                {/* Inspector Header */}
                <div className="flex flex-col gap-4 border-b border-slate-800 pb-6 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="rounded-md bg-cyan-950/80 border border-cyan-500/40 px-2 py-0.5 text-[10px] font-mono font-bold text-cyan-300 uppercase">
                        Active File Inspector
                      </span>
                      <span className="text-xs text-slate-400 font-mono">
                        {currentInspected.name}
                      </span>
                    </div>
                    <h3 className="mt-1.5 text-xl font-bold text-white">
                      Sanitized Data Preview &amp; Transformation Breakdown
                    </h3>
                  </div>

                  <div className="flex items-center flex-wrap gap-2.5">
                    <CopyDataButton
                      payload={{ type: "records", data: currentInspected.cleanedRows }}
                      showCsv={true}
                      showTsv={true}
                      showJson={true}
                      size="sm"
                      buttonTheme="cyan"
                    />

                    <button
                      onClick={() => handleDownloadSingleFile(currentInspected, "xlsx")}
                      className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-cyan-400 to-blue-600 px-4 py-2 text-xs font-bold text-slate-950 hover:brightness-110 transition shadow-md shadow-cyan-500/25"
                    >
                      <Download className="h-4 w-4" />
                      <span>Download Excel Workbook (.xlsx)</span>
                    </button>
                  </div>
                </div>

                {/* Performance stats chips for current inspected file */}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 font-mono text-xs">
                  <div className="rounded-xl border border-slate-800 bg-slate-900/60 p-3">
                    <span className="text-slate-400 text-[10px] uppercase">Final Clean Records</span>
                    <p className="text-lg font-bold text-cyan-300 mt-0.5">{currentInspected.stats.cleanedRows}</p>
                  </div>
                  <div className="rounded-xl border border-slate-800 bg-slate-900/60 p-3">
                    <span className="text-slate-400 text-[10px] uppercase">Duplicate Rows Purged</span>
                    <p className="text-lg font-bold text-rose-400 mt-0.5">-{currentInspected.stats.duplicatesPurged}</p>
                  </div>
                  <div className="rounded-xl border border-slate-800 bg-slate-900/60 p-3">
                    <span className="text-slate-400 text-[10px] uppercase">Normalizations Applied</span>
                    <p className="text-lg font-bold text-emerald-400 mt-0.5">{currentInspected.stats.fieldsSanitized}</p>
                  </div>
                  <div className="rounded-xl border border-slate-800 bg-slate-900/60 p-3">
                    <span className="text-slate-400 text-[10px] uppercase">Headers Standardized</span>
                    <p className="text-lg font-bold text-blue-400 mt-0.5">{currentInspected.stats.headersFixed}</p>
                  </div>
                </div>

                {/* Table search & row count */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="relative w-full sm:w-72">
                    <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Filter records..."
                      value={searchFilter}
                      onChange={(e) => setSearchFilter(e.target.value)}
                      className="w-full rounded-xl border border-slate-800 bg-slate-900/80 py-2 pl-9 pr-4 text-xs text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
                    />
                  </div>

                  <span className="text-xs text-slate-400 font-mono">
                    Showing {inspectedDisplayRows.length} of {currentInspected.cleanedRows.length} records
                  </span>
                </div>

                {/* Table preview */}
                <div className="overflow-x-auto rounded-2xl border border-slate-800 bg-slate-900/80 max-h-[380px]">
                  <table className="min-w-full text-left text-xs font-mono">
                    <thead className="sticky top-0 z-10 border-b border-slate-800 bg-slate-950 text-cyan-300 font-bold">
                      <tr>
                        <th className="px-3.5 py-3 text-slate-500 w-12 text-center">#</th>
                        {currentInspected.headers.map((h) => (
                          <th key={h} className="px-4 py-3 whitespace-nowrap">
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-200">
                      {inspectedDisplayRows.slice(0, 50).map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-800/40">
                          <td className="px-3.5 py-2.5 text-slate-500 text-center text-[10px]">
                            {idx + 1}
                          </td>
                          {currentInspected.headers.map((h, cIdx) => (
                            <td key={cIdx} className="px-4 py-2.5 whitespace-nowrap">
                              {row[h] || <span className="text-slate-600">&mdash;</span>}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default BulkDataCleaner;
