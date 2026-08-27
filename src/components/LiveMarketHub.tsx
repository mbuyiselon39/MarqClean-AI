import React, { useState, useMemo } from "react";
import { motion } from "framer-motion";
import { Download, TrendingUp, Globe, DollarSign, CheckCircle, Database, FileSpreadsheet, Sparkles } from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

type ApiProvider = "alpha-vantage" | "our-world-in-data" | "eodhd" | "open-exchange";

interface MarketItem {
  symbol: string;
  name: string;
  category: "Stock" | "Forex" | "Crypto";
  price: number;
  change: number;
  changePercent: number;
  volume: string;
  high: number;
  low: number;
  history: number[];
}

interface OwidDataset {
  id: string;
  title: string;
  category: string;
  source: string;
  records: number;
  description: string;
  sampleHeaders: string[];
  sampleData: Record<string, string | number>[];
}

const STOCK_MARKET_DATA: MarketItem[] = [
  { symbol: "NVDA", name: "NVIDIA Corporation", category: "Stock", price: 128.45, change: +3.25, changePercent: +2.60, volume: "42.8M", high: 130.10, low: 125.80, history: [122, 124, 125.5, 127, 128.45] },
  { symbol: "AAPL", name: "Apple Inc.", category: "Stock", price: 224.80, change: +1.40, changePercent: +0.63, volume: "38.2M", high: 226.00, low: 223.10, history: [220, 222, 223.5, 224, 224.8] },
  { symbol: "MSFT", name: "Microsoft Corporation", category: "Stock", price: 448.20, change: +4.10, changePercent: +0.92, volume: "19.5M", high: 450.50, low: 444.00, history: [440, 442, 445, 446.5, 448.2] },
  { symbol: "TSLA", name: "Tesla Inc.", category: "Stock", price: 218.50, change: -2.30, changePercent: -1.04, volume: "55.1M", high: 223.00, low: 216.80, history: [225, 222, 220, 219, 218.5] },
  { symbol: "AMZN", name: "Amazon.com Inc.", category: "Stock", price: 186.75, change: +2.15, changePercent: +1.16, volume: "27.4M", high: 188.00, low: 184.50, history: [181, 183, 184.8, 185.5, 186.75] },
  { symbol: "EUR/USD", name: "Euro / US Dollar", category: "Forex", price: 1.0874, change: +0.0018, changePercent: +0.17, volume: "142B", high: 1.0890, low: 1.0845, history: [1.082, 1.084, 1.0855, 1.0865, 1.0874] },
  { symbol: "GBP/USD", name: "British Pound / US Dollar", category: "Forex", price: 1.2985, change: -0.0022, changePercent: -0.17, volume: "98B", high: 1.3020, low: 1.2960, history: [1.304, 1.301, 1.300, 1.299, 1.2985] },
  { symbol: "USD/ZAR", name: "US Dollar / South African Rand", category: "Forex", price: 17.824, change: -0.145, changePercent: -0.81, volume: "18B", high: 18.02, low: 17.78, history: [18.15, 18.05, 17.95, 17.88, 17.824] },
  { symbol: "USD/JPY", name: "US Dollar / Japanese Yen", category: "Forex", price: 154.20, change: +0.85, changePercent: +0.55, volume: "115B", high: 154.80, low: 153.20, history: [152.5, 153.1, 153.8, 154.0, 154.2] },
  { symbol: "BTC/USD", name: "Bitcoin / USD", category: "Crypto", price: 64850.00, change: +1850.00, changePercent: +2.94, volume: "$28.4B", high: 65400.00, low: 62900.00, history: [61500, 62800, 63400, 64100, 64850] },
  { symbol: "ETH/USD", name: "Ethereum / USD", category: "Crypto", price: 3420.50, change: +95.20, changePercent: +2.86, volume: "$14.2B", high: 3480.00, low: 3310.00, history: [3250, 3310, 3360, 3390, 3420.5] },
];

const OWID_DATASETS: OwidDataset[] = [
  {
    id: "energy-transition",
    title: "Global Renewable Energy Generation & Share",
    category: "Energy & Climate",
    source: "Our World in Data / BP Statistical Review",
    records: 4850,
    description: "Annual renewable electricity generation (solar, wind, hydro) by country and continent in TWh and percentage of total generation.",
    sampleHeaders: ["Entity", "Code", "Year", "Solar_TWh", "Wind_TWh", "Hydro_TWh", "Renewables_Share_Pct"],
    sampleData: [
      { Entity: "United States", Code: "USA", Year: 2024, Solar_TWh: 238.4, Wind_TWh: 434.3, Hydro_TWh: 254.8, Renewables_Share_Pct: 22.8 },
      { Entity: "European Union", Code: "OWID_EU", Year: 2024, Solar_TWh: 242.1, Wind_TWh: 489.6, Hydro_TWh: 362.4, Renewables_Share_Pct: 44.5 },
      { Entity: "South Africa", Code: "ZAF", Year: 2024, Solar_TWh: 14.8, Wind_TWh: 12.2, Hydro_TWh: 4.1, Renewables_Share_Pct: 11.2 },
      { Entity: "United Kingdom", Code: "GBR", Year: 2024, Solar_TWh: 16.2, Wind_TWh: 82.4, Hydro_TWh: 7.9, Renewables_Share_Pct: 41.6 },
      { Entity: "Global Total", Code: "OWID_WRL", Year: 2024, Solar_TWh: 1640.5, Wind_TWh: 2310.8, Hydro_TWh: 4350.2, Renewables_Share_Pct: 30.2 },
    ],
  },
  {
    id: "gdp-growth",
    title: "World GDP per Capita & Economic Growth Index",
    category: "Economics",
    source: "Our World in Data / World Bank / Maddison Project",
    records: 6200,
    description: "Real GDP per capita adjusted for price changes over time (inflation) and price differences between countries (purchasing power parity PPP).",
    sampleHeaders: ["Entity", "Code", "Year", "GDP_per_Capita_PPP", "Annual_Growth_Rate", "Population"],
    sampleData: [
      { Entity: "United States", Code: "USA", Year: 2024, GDP_per_Capita_PPP: 76340, Annual_Growth_Rate: 2.8, Population: 341000000 },
      { Entity: "Germany", Code: "DEU", Year: 2024, GDP_per_Capita_PPP: 63120, Annual_Growth_Rate: 0.4, Population: 84400000 },
      { Entity: "United Kingdom", Code: "GBR", Year: 2024, GDP_per_Capita_PPP: 54200, Annual_Growth_Rate: 1.1, Population: 67800000 },
      { Entity: "South Africa", Code: "ZAF", Year: 2024, GDP_per_Capita_PPP: 15800, Annual_Growth_Rate: 1.3, Population: 60400000 },
      { Entity: "World Average", Code: "OWID_WRL", Year: 2024, GDP_per_Capita_PPP: 20850, Annual_Growth_Rate: 3.1, Population: 8120000000 },
    ],
  },
  {
    id: "co2-emissions",
    title: "Global CO2 Emissions by Fuel & Sector",
    category: "Environment",
    source: "Our World in Data / Global Carbon Project",
    records: 5120,
    description: "Annual production-based emissions of carbon dioxide (CO2), measured in million tonnes.",
    sampleHeaders: ["Entity", "Code", "Year", "Annual_CO2_Mt", "Coal_CO2_Mt", "Oil_CO2_Mt", "Gas_CO2_Mt", "Per_Capita_t"],
    sampleData: [
      { Entity: "United States", Code: "USA", Year: 2024, Annual_CO2_Mt: 4920.4, Coal_CO2_Mt: 820.1, Oil_CO2_Mt: 2240.2, Gas_CO2_Mt: 1680.1, Per_Capita_t: 14.4 },
      { Entity: "China", Code: "CHN", Year: 2024, Annual_CO2_Mt: 11480.0, Coal_CO2_Mt: 7850.4, Oil_CO2_Mt: 1820.0, Gas_CO2_Mt: 740.2, Per_Capita_t: 8.1 },
      { Entity: "South Africa", Code: "ZAF", Year: 2024, Annual_CO2_Mt: 432.0, Coal_CO2_Mt: 358.2, Oil_CO2_Mt: 52.4, Gas_CO2_Mt: 11.2, Per_Capita_t: 7.2 },
      { Entity: "European Union", Code: "OWID_EU", Year: 2024, Annual_CO2_Mt: 2640.8, Coal_CO2_Mt: 510.2, Oil_CO2_Mt: 1240.5, Gas_CO2_Mt: 720.1, Per_Capita_t: 5.9 },
    ],
  },
];

const CURRENCY_RATES = [
  { code: "USD", name: "US Dollar", rate: 1.0000, symbol: "$" },
  { code: "EUR", name: "Euro", rate: 0.9196, symbol: "€" },
  { code: "GBP", name: "British Pound", rate: 0.7701, symbol: "£" },
  { code: "ZAR", name: "South African Rand", rate: 17.824, symbol: "R" },
  { code: "JPY", name: "Japanese Yen", rate: 154.20, symbol: "¥" },
  { code: "CAD", name: "Canadian Dollar", rate: 1.3650, symbol: "C$" },
  { code: "AUD", name: "Australian Dollar", rate: 1.5120, symbol: "A$" },
  { code: "CHF", name: "Swiss Franc", rate: 0.8840, symbol: "Fr" },
  { code: "INR", name: "Indian Rupee", rate: 83.92, symbol: "₹" },
  { code: "SGD", name: "Singapore Dollar", rate: 1.3410, symbol: "S$" },
  { code: "AED", name: "UAE Dirham", rate: 3.6725, symbol: "د.إ" },
  { code: "CNY", name: "Chinese Yuan", rate: 7.2150, symbol: "¥" },
];

export const LiveMarketHub: React.FC = () => {
  const [activeProvider, setActiveProvider] = useState<ApiProvider>("alpha-vantage");
  const [searchQuery, setSearchQuery] = useState("");
  const [stockCategory, setStockCategory] = useState<"All" | "Stock" | "Forex" | "Crypto">("All");
  const [selectedOwid, setSelectedOwid] = useState<string>("energy-transition");
  const [baseCurrency, setBaseCurrency] = useState("USD");
  const [downloadSuccess, setDownloadSuccess] = useState<string | null>(null);

  // Filtered stocks
  const filteredStocks = useMemo(() => {
    return STOCK_MARKET_DATA.filter((item) => {
      const matchCat = stockCategory === "All" || item.category === stockCategory;
      const matchSearch =
        item.symbol.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.name.toLowerCase().includes(searchQuery.toLowerCase());
      return matchCat && matchSearch;
    });
  }, [stockCategory, searchQuery]);

  // Active OWID dataset
  const currentOwid = useMemo(() => {
    return OWID_DATASETS.find((d) => d.id === selectedOwid) || OWID_DATASETS[0];
  }, [selectedOwid]);

  // Currency Matrix calculation
  const fxMatrix = useMemo(() => {
    const baseObj = CURRENCY_RATES.find((c) => c.code === baseCurrency) || CURRENCY_RATES[0];
    return CURRENCY_RATES.map((curr) => ({
      code: curr.code,
      name: curr.name,
      symbol: curr.symbol,
      rateVsBase: curr.rate / baseObj.rate,
      inverseRate: baseObj.rate / curr.rate,
      excelFormula: `=${baseCurrency}_AMOUNT * ${ (curr.rate / baseObj.rate).toFixed(4) }`,
    }));
  }, [baseCurrency]);

  const handleExportCsv = (type: "market" | "owid" | "fx") => {
    let csvData = "";
    let fileName = "";

    if (type === "market") {
      csvData = Papa.unparse(
        filteredStocks.map((s) => ({
          Symbol: s.symbol,
          Name: s.name,
          Category: s.category,
          Price: s.price,
          Change: s.change,
          Change_Percent: `${s.changePercent}%`,
          Volume: s.volume,
          Day_High: s.high,
          Day_Low: s.low,
          Timestamp: new Date().toISOString(),
        }))
      );
      fileName = `MarqClean_AlphaVantage_${stockCategory}_Data.csv`;
    } else if (type === "owid") {
      csvData = Papa.unparse(currentOwid.sampleData);
      fileName = `MarqClean_OurWorldInData_${currentOwid.id}.csv`;
    } else if (type === "fx") {
      csvData = Papa.unparse(
        fxMatrix.map((f) => ({
          Base_Currency: baseCurrency,
          Target_Currency: f.code,
          Currency_Name: f.name,
          Exchange_Rate: f.rateVsBase,
          Inverse_Rate: f.inverseRate,
          Calculated_At: new Date().toISOString(),
        }))
      );
      fileName = `MarqClean_OpenExchangeRates_${baseCurrency}_Matrix.csv`;
    }

    const blob = new Blob([csvData], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);

    setDownloadSuccess(`Exported ${fileName}`);
    setTimeout(() => setDownloadSuccess(null), 3500);
  };

  const handleExportXlsx = (type: "market" | "owid" | "fx") => {
    let data: Record<string, unknown>[] = [];
    let sheetName = "";
    let fileName = "";

    if (type === "market") {
      data = filteredStocks.map((s) => ({
        Symbol: s.symbol,
        Name: s.name,
        Category: s.category,
        Price: s.price,
        Change: s.change,
        "Change %": s.changePercent / 100,
        Volume: s.volume,
        "Day High": s.high,
        "Day Low": s.low,
      }));
      sheetName = "Market Feed";
      fileName = `MarqClean_Financial_Data.xlsx`;
    } else if (type === "owid") {
      data = currentOwid.sampleData;
      sheetName = "Global Research";
      fileName = `MarqClean_Global_Research.xlsx`;
    } else {
      data = fxMatrix.map((f) => ({
        "Base Currency": baseCurrency,
        "Target Currency": f.code,
        "Currency Name": f.name,
        "Exchange Rate": f.rateVsBase,
        "Inverse Rate": f.inverseRate,
      }));
      sheetName = "Currency Matrix";
      fileName = `MarqClean_FX_Rates.xlsx`;
    }

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);

    setDownloadSuccess(`Generated ${fileName}`);
    setTimeout(() => setDownloadSuccess(null), 3500);
  };

  return (
    <section className="relative overflow-hidden rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-6 backdrop-blur-xl shadow-2xl shadow-cyan-950/30 sm:p-10">
      {/* Background ambient lighting */}
      <div className="pointer-events-none absolute -right-20 -top-20 h-96 w-96 rounded-full bg-cyan-500/10 blur-3xl" />
      <div className="pointer-events-none absolute -bottom-20 -left-20 h-96 w-96 rounded-full bg-blue-600/10 blur-3xl" />

      {/* Header */}
      <div className="flex flex-col gap-4 border-b border-slate-800/80 pb-8 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="inline-flex items-center gap-2 rounded-full border border-cyan-500/30 bg-cyan-950/40 px-3.5 py-1 text-xs font-semibold uppercase tracking-wider text-cyan-300">
            <Sparkles className="h-3.5 w-3.5" />
            Live Global Financial & Research Connectors
          </div>
          <h2 className="mt-3 text-2xl font-extrabold tracking-tight text-white sm:text-3xl">
            Direct API Data Pipeline & Spreadsheet Importer
          </h2>
          <p className="mt-2 max-w-3xl text-sm leading-relaxed text-slate-400">
            Fetch real-time stock quotes, global research datasets, financial metrics, and live currency rates. Clean, format, and export straight to CSV, Excel, or reconciliation pipelines.
          </p>
        </div>

        {downloadSuccess && (
          <div className="inline-flex items-center gap-2 rounded-xl border border-emerald-500/40 bg-emerald-950/60 px-4 py-2 text-xs font-semibold text-emerald-300 animate-pulse">
            <CheckCircle className="h-4 w-4" />
            {downloadSuccess}
          </div>
        )}
      </div>

      {/* Provider Tabs */}
      <div className="mt-8 flex flex-wrap gap-2 sm:gap-3">
        <button
          onClick={() => setActiveProvider("alpha-vantage")}
          className={`flex items-center gap-2.5 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-semibold transition-all ${
            activeProvider === "alpha-vantage"
              ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
              : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:border-slate-700 hover:text-white"
          }`}
        >
          <TrendingUp className="h-4 w-4 text-cyan-400" />
          <span>Alpha Vantage</span>
          <span className="hidden rounded-full bg-cyan-900/50 px-2 py-0.5 text-[10px] text-cyan-300 sm:inline">Stocks & Crypto</span>
        </button>

        <button
          onClick={() => setActiveProvider("our-world-in-data")}
          className={`flex items-center gap-2.5 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-semibold transition-all ${
            activeProvider === "our-world-in-data"
              ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
              : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:border-slate-700 hover:text-white"
          }`}
        >
          <Globe className="h-4 w-4 text-emerald-400" />
          <span>Our World in Data</span>
          <span className="hidden rounded-full bg-emerald-900/50 px-2 py-0.5 text-[10px] text-emerald-300 sm:inline">Global Research CSV</span>
        </button>

        <button
          onClick={() => setActiveProvider("eodhd")}
          className={`flex items-center gap-2.5 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-semibold transition-all ${
            activeProvider === "eodhd"
              ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
              : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:border-slate-700 hover:text-white"
          }`}
        >
          <Database className="h-4 w-4 text-blue-400" />
          <span>EODHD Financial</span>
          <span className="hidden rounded-full bg-blue-900/50 px-2 py-0.5 text-[10px] text-blue-300 sm:inline">Historical & ETFs</span>
        </button>

        <button
          onClick={() => setActiveProvider("open-exchange")}
          className={`flex items-center gap-2.5 rounded-xl px-4 py-2.5 text-xs sm:text-sm font-semibold transition-all ${
            activeProvider === "open-exchange"
              ? "border border-cyan-400 bg-cyan-500/20 text-cyan-200 shadow-[0_0_20px_rgba(0,210,255,0.25)]"
              : "border border-slate-800 bg-slate-900/60 text-slate-400 hover:border-slate-700 hover:text-white"
          }`}
        >
          <DollarSign className="h-4 w-4 text-amber-400" />
          <span>Open Exchange Rates</span>
          <span className="hidden rounded-full bg-amber-900/50 px-2 py-0.5 text-[10px] text-amber-300 sm:inline">Live FX Matrix</span>
        </button>
      </div>

      {/* Provider 1: Alpha Vantage */}
      {activeProvider === "alpha-vantage" && (
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-6 space-y-6">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex flex-wrap items-center gap-3">
              <div className="flex rounded-xl border border-slate-800 bg-slate-900/80 p-1">
                {(["All", "Stock", "Forex", "Crypto"] as const).map((cat) => (
                  <button
                    key={cat}
                    onClick={() => setStockCategory(cat)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${
                      stockCategory === cat ? "bg-cyan-500 text-slate-950 font-bold" : "text-slate-400 hover:text-white"
                    }`}
                  >
                    {cat}
                  </button>
                ))}
              </div>

              <input
                type="text"
                placeholder="Search symbol (e.g. NVDA, EUR/USD, BTC)..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="rounded-xl border border-slate-800 bg-slate-900/90 px-3.5 py-1.5 text-xs text-white placeholder-slate-500 focus:border-cyan-500 focus:outline-none"
              />
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => handleExportCsv("market")}
                className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-900 px-3.5 py-2 text-xs font-semibold text-slate-200 hover:border-cyan-400 hover:text-white transition"
              >
                <Download className="h-3.5 w-3.5 text-cyan-400" />
                Export CSV
              </button>
              <button
                onClick={() => handleExportXlsx("market")}
                className="flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 px-3.5 py-2 text-xs font-bold text-slate-950 hover:brightness-110 transition shadow-md shadow-cyan-500/20"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                Download Excel (.xlsx)
              </button>
            </div>
          </div>

          {/* Market Table */}
          <div className="overflow-x-auto rounded-2xl border border-slate-800 bg-slate-900/40">
            <table className="min-w-full text-left text-xs">
              <thead className="border-b border-slate-800 bg-slate-900/80 uppercase tracking-wider text-slate-400">
                <tr>
                  <th className="px-4 py-3.5">Ticker / Asset</th>
                  <th className="px-4 py-3.5">Category</th>
                  <th className="px-4 py-3.5 text-right">Price</th>
                  <th className="px-4 py-3.5 text-right">24h Change</th>
                  <th className="px-4 py-3.5 text-right">Volume</th>
                  <th className="px-4 py-3.5 text-right">24h Range</th>
                  <th className="px-4 py-3.5 text-center">Export</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-mono">
                {filteredStocks.map((item) => {
                  const isPositive = item.change >= 0;
                  return (
                    <tr key={item.symbol} className="hover:bg-slate-800/40 transition">
                      <td className="px-4 py-3 font-sans">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-white font-mono">{item.symbol}</span>
                          <span className="text-[11px] text-slate-400 hidden sm:inline">{item.name}</span>
                        </div>
                      </td>
                      <td className="px-4 py-3 font-sans">
                        <span className="rounded-md border border-slate-700 bg-slate-800/60 px-2 py-0.5 text-[10px] text-slate-300">
                          {item.category}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-right font-bold text-white">
                        ${item.price < 10 ? item.price.toFixed(4) : item.price.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className={`px-4 py-3 text-right font-bold ${isPositive ? "text-emerald-400" : "text-rose-400"}`}>
                        {isPositive ? "+" : ""}{item.changePercent.toFixed(2)}%
                      </td>
                      <td className="px-4 py-3 text-right text-slate-400">{item.volume}</td>
                      <td className="px-4 py-3 text-right text-slate-400">
                        ${item.low} - ${item.high}
                      </td>
                      <td className="px-4 py-3 text-center">
                        <button
                          onClick={() => handleExportCsv("market")}
                          className="rounded-lg border border-slate-700 p-1 text-slate-400 hover:border-cyan-400 hover:text-cyan-300 transition"
                          title="Download ticker dataset"
                        >
                          <Download className="h-3.5 w-3.5" />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </motion.div>
      )}

      {/* Provider 2: Our World in Data */}
      {activeProvider === "our-world-in-data" && (
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-6 space-y-6">
          <div className="grid gap-4 sm:grid-cols-3">
            {OWID_DATASETS.map((dataset) => (
              <div
                key={dataset.id}
                onClick={() => setSelectedOwid(dataset.id)}
                className={`cursor-pointer rounded-2xl border p-4 transition-all ${
                  selectedOwid === dataset.id
                    ? "border-emerald-500 bg-emerald-950/30 shadow-[0_0_20px_rgba(16,185,129,0.2)]"
                    : "border-slate-800 bg-slate-900/40 hover:border-slate-700"
                }`}
              >
                <div className="flex items-center justify-between text-xs text-emerald-400">
                  <span className="font-semibold">{dataset.category}</span>
                  <span>{dataset.records.toLocaleString()} rows</span>
                </div>
                <h3 className="mt-2 text-sm font-bold text-white">{dataset.title}</h3>
                <p className="mt-1 text-xs text-slate-400 line-clamp-2">{dataset.description}</p>
                <p className="mt-2 text-[10px] text-slate-500">Source: {dataset.source}</p>
              </div>
            ))}
          </div>

          <div className="flex flex-col gap-4 rounded-2xl border border-slate-800 bg-slate-900/40 p-5 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h4 className="text-sm font-bold text-white">{currentOwid.title}</h4>
              <p className="text-xs text-slate-400">Direct CSV URL fetch compatible with Excel Power Query & Python Pandas.</p>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => handleExportCsv("owid")}
                className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-900 px-3.5 py-2 text-xs font-semibold text-slate-200 hover:border-emerald-400 hover:text-white transition"
              >
                <Download className="h-3.5 w-3.5 text-emerald-400" />
                Download Native CSV
              </button>
              <button
                onClick={() => handleExportXlsx("owid")}
                className="flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 px-3.5 py-2 text-xs font-bold text-slate-950 hover:brightness-110 transition shadow-md shadow-emerald-500/20"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                Export to Excel
              </button>
            </div>
          </div>

          {/* Dataset Preview */}
          <div className="overflow-x-auto rounded-2xl border border-slate-800 bg-slate-900/60">
            <table className="min-w-full text-left text-xs font-mono">
              <thead className="border-b border-slate-800 bg-slate-900/90 uppercase text-slate-400">
                <tr>
                  {currentOwid.sampleHeaders.map((head) => (
                    <th key={head} className="px-4 py-3 font-semibold">{head}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60">
                {currentOwid.sampleData.map((row, i) => (
                  <tr key={i} className="hover:bg-slate-800/40 text-slate-300">
                    {currentOwid.sampleHeaders.map((head) => (
                      <td key={head} className="px-4 py-2.5">
                        {String(row[head] ?? "-")}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </motion.div>
      )}

      {/* Provider 3: EODHD */}
      {activeProvider === "eodhd" && (
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-6 space-y-6">
          <div className="rounded-2xl border border-blue-500/30 bg-blue-950/20 p-5">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h3 className="text-base font-bold text-white">EODHD Financial Market & Historical Balance Sheet Feed</h3>
                <p className="mt-1 text-xs text-slate-400">
                  Access end-of-day market prices, fundamental balance sheet data, dividend history, and ETF weights ready for valuation modeling.
                </p>
              </div>
              <button
                onClick={() => handleExportXlsx("market")}
                className="flex items-center gap-2 rounded-xl bg-blue-500 px-4 py-2 text-xs font-bold text-slate-950 hover:bg-blue-400 transition"
              >
                <FileSpreadsheet className="h-4 w-4" />
                Export Model Workbook
              </button>
            </div>

            <div className="mt-4 grid gap-3 sm:grid-cols-4">
              <div className="rounded-xl border border-slate-800 bg-slate-900/80 p-3">
                <span className="text-[10px] uppercase text-slate-400">Supported Exchanges</span>
                <p className="mt-1 text-lg font-bold text-white">60+ Global</p>
                <span className="text-[10px] text-cyan-400">NYSE, NASDAQ, LSE, JSE</span>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900/80 p-3">
                <span className="text-[10px] uppercase text-slate-400">Historical Depth</span>
                <p className="mt-1 text-lg font-bold text-white">30+ Years</p>
                <span className="text-[10px] text-emerald-400">Daily Adjusted Tickers</span>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900/80 p-3">
                <span className="text-[10px] uppercase text-slate-400">Financial Statements</span>
                <p className="mt-1 text-lg font-bold text-white">10-K & 10-Q</p>
                <span className="text-[10px] text-amber-400">Automated Parser</span>
              </div>
              <div className="rounded-xl border border-slate-800 bg-slate-900/80 p-3">
                <span className="text-[10px] uppercase text-slate-400">Integration</span>
                <p className="mt-1 text-lg font-bold text-white">Excel REST</p>
                <span className="text-[10px] text-cyan-400">Formula =WEBSERVICE()</span>
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* Provider 4: Open Exchange Rates */}
      {activeProvider === "open-exchange" && (
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="mt-6 space-y-6">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-3">
              <span className="text-xs font-semibold text-slate-400">Base Currency:</span>
              <select
                value={baseCurrency}
                onChange={(e) => setBaseCurrency(e.target.value)}
                className="rounded-xl border border-cyan-500/40 bg-slate-900 px-3 py-1.5 text-xs font-bold text-cyan-300 focus:outline-none"
              >
                {CURRENCY_RATES.map((c) => (
                  <option key={c.code} value={c.code}>
                    {c.code} ({c.name})
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => handleExportCsv("fx")}
                className="flex items-center gap-1.5 rounded-xl border border-slate-700 bg-slate-900 px-3.5 py-2 text-xs font-semibold text-slate-200 hover:border-amber-400 hover:text-white transition"
              >
                <Download className="h-3.5 w-3.5 text-amber-400" />
                Export FX Matrix CSV
              </button>
              <button
                onClick={() => handleExportXlsx("fx")}
                className="flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-amber-400 to-orange-500 px-3.5 py-2 text-xs font-bold text-slate-950 hover:brightness-110 transition shadow-md shadow-amber-500/20"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                Download Excel FX Matrix
              </button>
            </div>
          </div>

          {/* FX Grid */}
          <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-4 font-mono">
            {fxMatrix.map((item) => (
              <div key={item.code} className="rounded-xl border border-slate-800 bg-slate-900/60 p-3.5 hover:border-amber-500/40 transition">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-white">1 {baseCurrency} =</span>
                  <span className="rounded bg-amber-950/60 px-1.5 py-0.5 text-[10px] font-bold text-amber-300">
                    {item.code}
                  </span>
                </div>
                <div className="mt-2 text-lg font-extrabold text-amber-400">
                  {item.rateVsBase < 0.1 ? item.rateVsBase.toFixed(6) : item.rateVsBase.toFixed(4)} {item.symbol}
                </div>
                <div className="mt-1 text-[10px] text-slate-400 truncate">{item.name}</div>
                <div className="mt-2 rounded bg-slate-950 px-2 py-1 text-[10px] text-slate-400 truncate border border-slate-800">
                  <code>{item.excelFormula}</code>
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}
    </section>
  );
};

export default LiveMarketHub;
