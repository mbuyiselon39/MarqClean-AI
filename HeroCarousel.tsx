import { useRef, useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onOpenReconciliation: () => void;
  onOpenExcelAutomation?: () => void;
  onFileDrop?: (file: File) => void;
};

function LineIcon({ type }: { type: "home" | "upload" | "clean" | "format" | "reconcile" | "formula" | "arrow" }) {
  const paths = {
    home: <><path d="m4 11 8-7 8 7" /><path d="M6.5 10.5V20h11v-9.5" /></>,
    upload: <><path d="M12 16V4" /><path d="m7 9 5-5 5 5" /><path d="M4 16.5V19a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-2.5" /></>,
    clean: <><path d="M4 7h16M4 12h16M4 17h10" /><path d="m16 16 2 2 3-3" /></>,
    format: <><rect x="4" y="4" width="16" height="16" rx="2" /><path d="M8 9h8M8 13h5M8 17h8" /></>,
    reconcile: <><path d="M7 7h10" /><path d="m14 4 3 3-3 3" /><path d="M17 17H7" /><path d="m10 14-3 3 3 3" /></>,
    formula: <><path d="M5 7h14M5 12h14M5 17h14" /><path d="m9 5 6 14" /></>,
    arrow: <><path d="M5 12h14" /><path d="m13 6 6 6-6 6" /></>,
  }[type];
  return <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths}</svg>;
}

function WorkspacePreview({ onCleanFile, onOpenReconciliation, onOpenExcelAutomation, onFileDrop }: Pick<HeroActions, "onCleanFile" | "onOpenReconciliation" | "onOpenExcelAutomation" | "onFileDrop">) {
  const [active, setActive] = useState("Smart Drop");
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const panels: Record<string, { title: string; label: string }> = {
    Overview: { label: "OVERVIEW", title: "Your data workspace." },
    "Smart Drop": { label: "SMART DROP", title: "Ready for your data." },
    Cleaners: { label: "CLEANERS", title: "Standardise and clean." },
    Formatters: { label: "FORMATTERS", title: "Make workbooks consistent." },
    Reconciliation: { label: "RECONCILIATION", title: "Compare records with confidence." },
    "Excel Functions": { label: "EXCEL FUNCTIONS", title: "Build formulas faster." },
    "Data Tools": { label: "DATA TOOLS", title: "Explore and transform." },
  };
  const sidebar = [["Overview","home"],["Smart Drop","upload"],["Cleaners","clean"],["Formatters","format"],["Reconciliation","reconcile"],["Excel Functions","formula"],["Data Tools","format"]] as const;
  const select = (item: string) => setActive(item);
  const action = (label: string) => {
    if (label === "Clean & format") onCleanFile();
    else if (label === "Reconcile files") onOpenReconciliation();
    else onOpenExcelAutomation?.();
  };
  return <div className="mc-linear-preview" aria-label="Interactive MarqClean AI workspace preview">
    <div className="mc-linear-preview__chrome">
      <div className="mc-linear-preview__brand"><span className="mc-linear-dot" /> MarqClean AI</div>
      <button type="button" className="mc-linear-preview__search" onClick={() => window.dispatchEvent(new CustomEvent("marqclean:open-command-palette"))} aria-label="Open global search"><LineIcon type="upload" /><span>Search tools, files, formulas</span></button>
      <div className="mc-linear-preview__avatar">M</div>
    </div>
    <div className="mc-linear-preview__body">
      <aside role="tablist" aria-label="Workspace preview navigation">
        <small>WORKSPACE</small>
        {sidebar.map(([item, icon]) => <button type="button" role="tab" aria-selected={active === item} key={item} className={active === item ? "is-active" : ""} onClick={() => select(item)} onKeyDown={(e) => { if (e.key === "ArrowDown" || e.key === "ArrowRight") { e.preventDefault(); const next = sidebar[(sidebar.findIndex(([name]) => name === item)+1)%sidebar.length][0]; select(next); } if (e.key === "ArrowUp" || e.key === "ArrowLeft") { e.preventDefault(); const next = sidebar[(sidebar.findIndex(([name]) => name === item)-1+sidebar.length)%sidebar.length][0]; select(next); } }}><span><LineIcon type={icon} /></span>{item}</button>)}
      </aside>
      <main>
        <div className="mc-linear-preview__heading"><div><small>{panels[active].label}</small><strong>{panels[active].title}</strong></div><span className="mc-local-pill">Processed locally</span></div>
        {active === "Smart Drop" ? <><input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" className="sr-only" onChange={(e) => { const file = e.target.files?.[0]; if (file) onFileDrop?.(file); }} /><button type="button" className={`mc-linear-preview__drop ${dragging ? "is-dragging" : ""}`} onClick={() => inputRef.current?.click()} onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(e) => { e.preventDefault(); setDragging(false); const file = e.dataTransfer.files?.[0]; if (file) onFileDrop?.(file); }}><span className="mc-drop-icon"><LineIcon type="upload" /></span><strong>{dragging ? "Release to process" : "Drop CSV or Excel"}</strong><small>or choose a file from your computer</small></button></> : <div className="mc-preview-panel-message"><strong>{panels[active].title}</strong><small>Select an action below or open the workspace for the full workflow.</small></div>}
        <div className="mc-linear-preview__cards">{[["Clean & format","clean"],["Reconcile files","reconcile"],["Run formulas","formula"]].map(([label, icon]) => <button type="button" key={label} onClick={() => action(label)}><span><LineIcon type={icon as "clean"|"reconcile"|"formula"} /></span><b>{label}</b><em><LineIcon type="arrow" /></em></button>)}</div>
      </main>
    </div>
  </div>;
}

export default function HeroCarousel({ onCleanFile, onOpenReconciliation, onOpenExcelAutomation, onOpenDataToolbox, onFileDrop }: HeroActions) {
  const [active, setActive] = useState(0);
  const tabs = [
    { label: "Data Cleaning", href: "#cleaner" },
    { label: "Excel Automation", href: "/excel-automation" },
    { label: "Reconciliation", href: "/reconciliation-hub" },
    { label: "Data Toolbox", href: "/data-toolbox" },
  ];

  return (
    <section className="mc-minimal-hero" aria-label="MarqClean AI data preparation workspace">
      <div className="mc-minimal-hero__inner">
        <div className="mc-minimal-hero__copy">
          <div className="mc-minimal-hero__eyebrow"><span /> Browser-local data workspace</div>
          <h1>Clean data.<br /><em>Clear decisions.</em></h1>
          <p className="mc-hero-copy">Clean, transform, reconcile and analyse CSV, Excel and PDF data without sending your files to a server.</p>
          <div className="mc-minimal-hero__actions">
            <button onClick={onCleanFile}>Start free <LineIcon type="arrow" /></button>
            <a className="secondary" href="#free-tools">See the tools</a>
          </div>
          <div className="mc-minimal-hero__trust"><span className="mc-status-icon" aria-hidden="true"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="m3 8 3 3 7-7" /></svg></span> Runs locally in your browser <i /> No sign-up <i /> No credit card</div>
        </div>
        <div className="mc-minimal-hero__product" aria-label="MarqClean AI workspace preview"><WorkspacePreview onCleanFile={onCleanFile} onOpenReconciliation={onOpenReconciliation} onOpenExcelAutomation={onOpenExcelAutomation} onFileDrop={onFileDrop} /></div>
      </div>
      <nav className="mc-minimal-hero__tabs" aria-label="Product workspaces">
        {tabs.map((tab, i) => (
          <a
            key={tab.label}
            className={active === i ? "is-active" : ""}
            href={tab.href}
            aria-current={active === i ? "page" : undefined}
            onClick={() => setActive(i)}
          >
            {tab.label}
          </a>
        ))}
      </nav>
    </section>
  );
}
