import { useEffect, useRef, useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onSampleLeads: () => void;
  onOpenReconciliation: () => void;
};

type Concept = {
  key: string;
  tabLabel: string;
  tag: string;
  headline: React.ReactNode;
  subhead: string;
  visual: React.ReactNode;
};

// ---------------------------------------------------------------------------
// 3D visual placeholders (pure CSS / SVG, ready to swap for real 3D assets)
// ---------------------------------------------------------------------------

function VisualDiff() {
  return (
    <div className="relative flex h-full w-full items-center justify-center">
      <div className="grid w-full max-w-md grid-cols-2 gap-3 font-mono text-[11px]">
        <div className="rounded-lg border border-white/8 bg-white/[0.02] p-3">
          <div className="mb-2 text-[9px] uppercase tracking-widest text-white/40">Raw input</div>
          <div className="space-y-2 text-white/45">
            <div className="truncate">john SMITH , n/a , chicago</div>
            <div className="truncate">MARY-jane O'brien,, NYC</div>
            <div className="truncate">bob@ACME .com, 0800555</div>
          </div>
        </div>
        <div className="rounded-lg border border-[#0D9488]/30 bg-[#0D9488]/10 p-3">
          <div className="mb-2 text-[9px] uppercase tracking-widest text-[#5EEAD4]">Cleaned</div>
          <div className="space-y-2 text-white/90">
            <div className="mc-anim-float truncate rounded bg-[#FB6F58]/15 px-1.5 py-0.5 ring-1 ring-[#FB6F58]/30">John Smith · Chicago</div>
            <div className="truncate rounded bg-[#FB6F58]/15 px-1.5 py-0.5 ring-1 ring-[#FB6F58]/30">Mary-Jane O'Brien · NYC</div>
            <div className="truncate rounded bg-[#FB6F58]/15 px-1.5 py-0.5 ring-1 ring-[#FB6F58]/30">bob@acme.com</div>
          </div>
        </div>
      </div>
      <div
        className="pointer-events-none absolute inset-y-6 left-1/2 w-px -translate-x-1/2 bg-gradient-to-b from-transparent via-[#0D9488] to-transparent"
        style={{ animation: "mc-pulse-line 2.4s ease-in-out infinite" }}
      />
    </div>
  );
}

function VisualGrid() {
  const rows = [
    ["Date", "Account", "Budget", "Actual"],
    ["1/31", "Revenue", "253,044", "287,922"],
    ["2/28", "COGS", "88,565", "100,773"],
    ["3/31", "Gross Profit", "164,479", "187,149"],
  ];
  return (
    <div className="flex h-full w-full flex-col items-center justify-center gap-3">
      <div className="w-full max-w-sm overflow-hidden rounded-lg border border-white/10">
        <div className="grid grid-cols-[22px_repeat(4,1fr)] bg-white/[0.05] text-[9px] text-white/35">
          <div />
          {["A", "B", "C", "D"].map((c) => (
            <div key={c} className={`truncate border-l border-white/8 px-2 py-1 ${c === "C" ? "bg-[#0D9488]/20 font-semibold text-[#5EEAD4]" : ""}`}>{c}</div>
          ))}
        </div>
        {rows.map((row, ri) => (
          <div key={ri} className={`grid grid-cols-[22px_repeat(4,1fr)] text-[10px] ${ri === 0 ? "bg-white/[0.07] font-semibold text-white" : ri % 2 ? "bg-white/[0.015] text-white/70" : "text-white/70"}`}>
            <div className="border-t border-white/8 px-1 py-1 text-[9px] text-white/25">{ri === 0 ? "" : ri}</div>
            {row.map((cell, ci) => (
              <div key={ci} className="truncate border-l border-t border-white/8 px-2 py-1">{cell}</div>
            ))}
          </div>
        ))}
      </div>
      <div className="mc-anim-float inline-flex items-center gap-1.5 rounded-md border border-[#0D9488]/30 bg-[#12142b] px-2.5 py-1.5 font-mono text-[10px] text-[#5EEAD4] shadow-[0_8px_24px_rgba(13,148,136,0.3)]">
        <span className="h-1.5 w-1.5 rounded-full bg-[#FB6F58]" /> =SUBTOTAL(109,[Actual])
      </div>
    </div>
  );
}

function VisualMatch() {
  const rows: Array<{ bank: string; ledger: string | null }> = [
    { bank: "R1042 · $4,230.00", ledger: "B-9981 · $4,230.00" },
    { bank: "R1043 · $1,880.50", ledger: null },
    { bank: "R1044 · $920.00", ledger: "B-9984 · $920.00" },
  ];
  return (
    <div className="flex h-full w-full flex-col items-center justify-center gap-2.5">
      {rows.map((r) => (
        <div key={r.bank} className="grid w-full max-w-sm grid-cols-[1fr_auto_1fr] items-center gap-2 text-[10px]">
          <div className="truncate rounded-md border border-[#0D9488]/30 bg-[#0D9488]/10 px-2.5 py-1.5 text-white/90">{r.bank}</div>
          <div className="flex items-center justify-center">
            {r.ledger ? (
              <svg className="h-3.5 w-3.5 text-[#FB6F58]" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M20 6 9 17l-5-5" />
              </svg>
            ) : (
              <span className="h-1 w-6 rounded-full bg-white/15" />
            )}
          </div>
          <div className={`truncate rounded-md border px-2.5 py-1.5 ${r.ledger ? "border-[#FB6F58]/30 bg-[#FB6F58]/10 text-white/90" : "border-white/8 bg-white/[0.02] text-white/35"}`}>
            {r.ledger ?? "— no match —"}
          </div>
        </div>
      ))}
    </div>
  );
}

function VisualPalette() {
  const rows = [
    { name: "XLOOKUP()", tag: "Lookup", active: true },
    { name: "SUMIFS()", tag: "Math", active: false },
    { name: "TEXTJOIN()", tag: "Text", active: false },
    { name: "IFERROR()", tag: "Logical", active: false },
  ];
  return (
    <div className="flex h-full w-full items-center justify-center">
      <div className="w-full max-w-sm rounded-lg border border-white/10 bg-white/[0.02] p-3">
        <div className="flex items-center gap-2 rounded-md border border-white/10 bg-black/25 px-3 py-2">
          <span className="text-white/40">⌕</span>
          <span className="font-mono text-xs text-white/70">xlookup</span>
          <span className="ml-auto rounded border border-white/10 px-1.5 py-0.5 text-[9px] text-white/35">⌘K</span>
        </div>
        <div className="mt-2 space-y-1">
          {rows.map((r) => (
            <div key={r.name} className={`flex items-center justify-between rounded-md px-2 py-1.5 ${r.active ? "bg-[#0D9488]/15" : ""}`}>
              <span className={`font-mono text-[11px] ${r.active ? "text-white" : "text-white/55"}`}>{r.name}</span>
              <span className={`text-[9px] uppercase tracking-widest ${r.active ? "text-[#5EEAD4]" : "text-white/30"}`}>{r.tag}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

export default function HeroCarousel({ onCleanFile, onSampleLeads, onOpenReconciliation }: HeroActions) {
  const [active, setActive] = useState(0);
  const rootRef = useRef<HTMLDivElement>(null);
  const [tilt, setTilt] = useState({ x: 0, y: 0 });

  // Auto-advance the carousel
  useEffect(() => {
    const timer = window.setInterval(() => setActive((a) => (a + 1) % 4), 8000);
    return () => window.clearInterval(timer);
  }, []);

  function onMouseMove(event: React.MouseEvent<HTMLDivElement>) {
    const rect = rootRef.current?.getBoundingClientRect();
    if (!rect) return;
    const px = (event.clientX - rect.left) / rect.width - 0.5;
    const py = (event.clientY - rect.top) / rect.height - 0.5;
    setTilt({ x: px, y: py });
  }

  const concepts: Concept[] = [
    {
      key: "cleaner",
      tabLabel: "Quick Data & CSV Cleaner",
      tag: "AUTOMATED CLEANING",
      headline: (<>Messy Data, <span className="mc-text-cyan">Instantly Refined.</span></>),
      subhead: "Ditch manual data cleaning. Let AI parse, sort, deduplicate, and fix formatting issues across large datasets automatically.",
      visual: <VisualDiff />,
    },
    {
      key: "excel-automation",
      tabLabel: "Excel Automation",
      tag: "STRUCTURE ENGINE",
      headline: (<><span className="mc-text-cyan">Structure</span> at Scale.</>),
      subhead: "Ditch manual pivot tables. Let AI parse, sort, and inject intelligent subtotals across millions of rows automatically.",
      visual: <VisualGrid />,
    },
    {
      key: "reconciliation-hub",
      tabLabel: "Reconciliation Hub",
      tag: "VALIDATION ENGINE",
      headline: (<>Match Every Record, <span className="mc-text-violet">Automatically.</span></>),
      subhead: "Extract, structure, and verify PDF statements and Excel master files, then produce a colour-coded reconciliation report without manual matching.",
      visual: <VisualMatch />,
    },
    {
      key: "data-toolbox",
      tabLabel: "Data Toolbox",
      tag: "ALL-IN-ONE PIPELINE",
      headline: (<>The <span className="mc-text-mixed">Swiss Army Knife</span> for Global Data Teams.</>),
      subhead: "A single, privacy-first pipeline built to handle every messy spreadsheet scenario across marketing, operations, and finance.",
      visual: <VisualPalette />,
    },
  ];

  const current = concepts[active];

  return (
    <section
      ref={rootRef}
      onMouseMove={onMouseMove}
      onMouseLeave={() => setTilt({ x: 0, y: 0 })}
      className="mc-hero-root relative flex items-center overflow-hidden text-white lg:min-h-[88vh]"
    >
      <div className="mc-grid-overlay pointer-events-none absolute inset-0" aria-hidden="true" />
      <div className="pointer-events-none absolute inset-0 overflow-hidden" aria-hidden="true">
        <div className="mc-anim-aurora-1 absolute -left-24 top-10 h-[28rem] w-[28rem] rounded-full bg-[#0D9488]/20 blur-[100px]" />
        <div className="mc-anim-aurora-2 absolute -right-16 top-1/3 h-[26rem] w-[26rem] rounded-full bg-[#14B8A6]/15 blur-[100px]" />
        <div className="mc-anim-aurora-1 absolute bottom-[-6rem] left-1/3 h-96 w-96 rounded-full bg-[#FB6F58]/10 blur-[100px]" style={{ animationDelay: "-9s" }} />
      </div>

      <div className="relative mx-auto w-full max-w-[1400px] px-5 pb-16 pt-24 sm:pb-20 sm:pt-28 lg:px-8 lg:py-24">
        {/* Tabs */}
        <div className="mb-6 flex flex-wrap gap-2" role="tablist" aria-label="Service concepts">
          {concepts.map((c, i) => (
            <button
              key={c.key}
              role="tab"
              aria-selected={active === i}
              onClick={() => setActive(i)}
              className={`mc-mono rounded-full border px-4 py-2 text-[11px] uppercase transition ${active === i ? "mc-tab-active" : "border-white/10 text-white/55 hover:border-white/30 hover:text-white/80"}`}
            >
              {c.tabLabel}
            </button>
          ))}
        </div>

        {/* Split layout: asymmetric 5/7, per redesign plan */}
        <div className="grid items-center gap-10 lg:grid-cols-12 lg:gap-8">
          {/* Left column */}
          <div key={`text-${current.key}`} className="mc-fade lg:col-span-5">
            <span className="mc-mono inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/5 px-3 py-1 text-[11px] uppercase text-teal-200">
              <span className="relative flex h-1.5 w-1.5">
                <span className="mc-anim-glow-pulse absolute inline-flex h-full w-full rounded-full bg-[#0D9488]" />
                <span className="relative h-1.5 w-1.5 rounded-full bg-[#0D9488] shadow-[0_0_10px_#0D9488]" />
              </span>
              {current.tag}
            </span>
            <h1 className="mc-display mt-4 text-4xl font-extrabold leading-[1.03] tracking-[-0.03em] sm:text-5xl lg:text-6xl">
              {current.headline}
            </h1>
            <p className="mt-4 max-w-xl text-base leading-7 text-white/65 lg:text-lg">{current.subhead}</p>

            <div className="mt-6 flex flex-col gap-3 sm:flex-row sm:items-center">
              <button className="mc-cta-cyan rounded-full px-6 py-3 text-sm font-semibold" onClick={onCleanFile}>
                Clean a File Free
              </button>
              <button className="rounded-full border border-white/20 px-6 py-3 text-sm font-semibold text-white transition hover:border-white/60 hover:bg-white/5" onClick={onSampleLeads}>
                Try a Sample File
              </button>
              <button className="group inline-flex items-center gap-1.5 px-2 py-2.5 text-sm font-semibold text-teal-200 transition hover:text-white" onClick={onOpenReconciliation}>
                Open Reconciliation Hub
                <svg className="h-4 w-4 transition group-hover:translate-x-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12h14" /><path d="m12 5 7 7-7 7" /></svg>
              </button>
            </div>

            <div className="mt-6 flex flex-wrap items-center gap-x-5 gap-y-2 text-xs">
              <span className="inline-flex items-center gap-1.5 font-semibold text-emerald-300">
                <svg className="h-4 w-4 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M20 6 9 17l-5-5" /></svg>
                No file ever leaves your browser
              </span>
              <span className="inline-flex items-center gap-1.5 text-white/45">
                <span className="h-1.5 w-1.5 rounded-full bg-[#FB6F58]" /> 40+ Excel functions built in
              </span>
            </div>

            <p className="mc-mono mt-6 text-[11px] uppercase text-[#9CA3AF]">A product of Vertex Stream Technologies, a division of Vertex Stream Group</p>
          </div>

          {/* Right column visual with parallax, spans wider 7/12 per redesign plan */}
          <div
            key={`visual-${current.key}`}
            className="mc-fade relative hidden lg:col-span-7 lg:block"
            style={{ transform: `perspective(1000px) rotateY(${tilt.x * 8}deg) rotateX(${-tilt.y * 8}deg)` }}
          >
            <div className="absolute -inset-6 -z-10 rounded-[2.25rem] bg-gradient-to-br from-[#0D9488]/20 via-[#14B8A6]/10 to-transparent blur-2xl" aria-hidden="true" />
            <div className="mc-glass relative min-h-[20rem] rounded-[2rem] p-1.5 lg:min-h-[24rem]">
              {/* window chrome, reinforces "this is a live product view" */}
              <div className="flex items-center gap-1.5 px-4 py-3">
                <span className="h-2.5 w-2.5 rounded-full bg-white/15" />
                <span className="h-2.5 w-2.5 rounded-full bg-white/15" />
                <span className="h-2.5 w-2.5 rounded-full bg-white/15" />
              </div>
              <div className="relative h-[calc(100%-2.75rem)] px-5 pb-5">
                {current.visual}
              </div>
              <div className="pointer-events-none absolute inset-0 rounded-[2rem] ring-1 ring-inset ring-white/5" />
            </div>
          </div>
        </div>
      </div>

      {/* seamless fade into the dark workspace section below */}
      <div className="pointer-events-none absolute inset-x-0 bottom-0 h-32 bg-gradient-to-t from-[#0a0b16] to-transparent" aria-hidden="true" />
    </section>
  );
}
