import { useRef, useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onOpenReconciliation: () => void;
  onOpenExcelAutomation?: () => void;
  onFileDrop?: (file: File) => void;
};

type IconName = "overview" | "upload" | "clean" | "reconcile" | "formula" | "arrow" | "search" | "shield" | "spark";

function Icon({ name, size = 18 }: { name: IconName; size?: number }) {
  const paths = {
    overview: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    upload: <><path d="M12 16V4m-5 5 5-5 5 5" /><path d="M4 16v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3" /></>,
    clean: <><path d="M4 7h16M4 12h11M4 17h8" /><path d="m15 17 2 2 4-4" /></>,
    reconcile: <><path d="M4 8h16m-4-4 4 4-4 4M20 16H4m4-4-4 4 4 4" /></>,
    formula: <><rect x="3" y="3" width="18" height="18" rx="2" /><path d="m7 9 3 3-3 3m6 0h4" /></>,
    arrow: <><path d="M4 12h16m-6-6 6 6-6 6" /></>,
    search: <><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></>,
    shield: <><path d="m12 2 8 4v6c0 5-3.4 8.3-8 10-4.6-1.7-8-5-8-10V6l8-4Z" /><path d="m9 12 2 2 4-4" /></>,
    spark: <><path d="m12 2 1.8 7.2L21 11l-7.2 1.8L12 20l-1.8-7.2L3 11l7.2-1.8L12 2ZM19 18l.6 1.4L21 20l-1.4.6L19 22l-.6-1.4L17 20l1.4-.6L19 18Z" /></>,
  }[name];
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths}</svg>;
}

const previewTabs = [
  { label: "Overview", icon: "overview" },
  { label: "Smart Drop", icon: "upload" },
  { label: "Data Cleaning", icon: "clean" },
  { label: "Reconciliation", icon: "reconcile" },
  { label: "Excel Automation", icon: "formula" },
] as const;

type PreviewTab = typeof previewTabs[number]["label"];

function WorkspacePreview({ onCleanFile, onOpenReconciliation, onOpenExcelAutomation, onFileDrop }: HeroActions) {
  const [active, setActive] = useState<PreviewTab>("Overview");
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const openTool = () => {
    if (active === "Reconciliation") onOpenReconciliation();
    else if (active === "Excel Automation") onOpenExcelAutomation?.();
    else onCleanFile();
  };
  const tabCopy: Record<PreviewTab, { heading: string; subheading: string; action: string }> = {
    Overview: { heading: "Data quality overview", subheading: "An illustrative look at your workspace", action: "Clean your data" },
    "Smart Drop": { heading: "Start with a file", subheading: "Drop a file to start a browser-local workflow", action: "Open cleaner" },
    "Data Cleaning": { heading: "Data cleaning", subheading: "Standardise records and review every change", action: "Open cleaner" },
    Reconciliation: { heading: "Reconciliation", subheading: "Compare sources and surface exceptions", action: "Open reconciliation" },
    "Excel Automation": { heading: "Excel automation", subheading: "Prepare polished, formula-ready workbooks", action: "Open Excel automation" },
  };
  const handleFile = (file?: File) => { if (file) onFileDrop?.(file); };

  return <div className="mc-preview" aria-label="Interactive MarqClean AI workspace preview">
    <div className="mc-preview__topbar">
      <div className="mc-preview__topbrand"><span className="mc-preview__mark">M</span><span>MarqClean <b>AI</b></span><span className="mc-preview__topdivider" /><span className="mc-preview__crumb">Workspace</span></div>
      <button type="button" className="mc-preview__search" onClick={() => window.dispatchEvent(new CustomEvent("marqclean:open-command-palette"))}><Icon name="search" size={15} /> <span>Search anything...</span><kbd>⌘ K</kbd></button>
      <span className="mc-preview__avatar" aria-hidden="true">MC</span>
    </div>
    <div className="mc-preview__layout">
      <aside className="mc-preview__sidebar" aria-label="Preview navigation">
        <span className="mc-preview__side-label">WORKSPACE</span>
        <div role="tablist" aria-label="Preview workspaces">
          {previewTabs.map((tab) => <button type="button" role="tab" aria-selected={active === tab.label} key={tab.label} className={active === tab.label ? "is-active" : ""} onClick={() => setActive(tab.label)}><Icon name={tab.icon} size={16} />{tab.label}</button>)}
        </div>
        <div className="mc-preview__sidebar-bottom"><span className="mc-preview__online" /> Browser-local processing</div>
      </aside>
      <div className="mc-preview__content">
        <div className="mc-preview__heading"><div><span className="mc-preview__breadcrumb">WORKSPACE / {active.toUpperCase()}</span><h2>{tabCopy[active].heading}</h2><p>{tabCopy[active].subheading}</p></div><span className="mc-preview__sample">SAMPLE VIEW</span></div>
        {active === "Smart Drop" ? <>
          <input ref={inputRef} className="sr-only" type="file" accept=".csv,.xlsx,.xls" aria-label="Choose a CSV or Excel file" onChange={(event) => { handleFile(event.target.files?.[0]); event.target.value = ""; }} />
          <button type="button" className={`mc-preview__drop ${dragging ? "is-dragging" : ""}`} onClick={() => inputRef.current?.click()} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); handleFile(event.dataTransfer.files?.[0]); }}><span className="mc-preview__drop-icon"><Icon name="upload" size={24} /></span><strong>{dragging ? "Release to process file" : "Drop a CSV or Excel file"}</strong><small>or click to browse · XLSX, XLS, CSV</small></button>
        </> : <>
          <div className="mc-preview__metrics">
            <div><span className="mc-preview__metric-icon mint"><Icon name="overview" size={16} /></span><small>Records processed</small><strong>12,480</strong><em>Illustrative data</em></div>
            <div><span className="mc-preview__metric-icon blue"><Icon name="clean" size={16} /></span><small>Data quality</small><strong>98.2<span>%</span></strong><em>After cleaning</em></div>
            <div><span className="mc-preview__metric-icon amber"><Icon name="spark" size={16} /></span><small>Issues flagged</small><strong>24</strong><em>Ready to review</em></div>
          </div>
          <div className="mc-preview__table-card"><div className="mc-preview__table-heading"><div><span className="mc-preview__table-icon"><Icon name="clean" size={15} /></span><strong>Recent records</strong></div><span>PREVIEW DATA</span></div><div className="mc-preview__table-scroll"><table><thead><tr><th>NAME</th><th>COMPANY</th><th>STATUS</th></tr></thead><tbody><tr><td><span className="mc-preview__row-avatar teal">JD</span> Jordan Davis</td><td>Northstar Group</td><td><span className="mc-preview__status"><i /> Clean</span></td></tr><tr><td><span className="mc-preview__row-avatar violet">AL</span> Alex Lee</td><td>Meridian Labs</td><td><span className="mc-preview__status"><i /> Clean</span></td></tr><tr><td><span className="mc-preview__row-avatar coral">SM</span> Sam Morgan</td><td>Atlas Partners</td><td><span className="mc-preview__status review"><i /> Review</span></td></tr></tbody></table></div></div>
        </>}
        <div className="mc-preview__bottom"><span><Icon name="shield" size={15} /> Your files stay in your browser</span><button type="button" onClick={openTool}>{tabCopy[active].action} <Icon name="arrow" size={15} /></button></div>
      </div>
    </div>
  </div>;
}

export default function HeroCarousel(props: HeroActions) {
  const workspaces = [
    { num: "01", label: "Data cleaning", description: "From messy to ready", href: "#cleaner" },
    { num: "02", label: "Excel automation", description: "Less manual work", href: "/excel-automation" },
    { num: "03", label: "Reconciliation", description: "Every detail accounted for", href: "/reconciliation-hub" },
    { num: "04", label: "Data toolbox", description: "More ways to work", href: "/data-toolbox" },
  ];
  return <section className="mc-next-hero" aria-label="MarqClean AI data operations platform">
    <div className="mc-next-hero__ambient" aria-hidden="true" />
    <div className="mc-next-hero__inner">
      <div className="mc-next-hero__copy">
        <div className="mc-next-hero__eyebrow"><span className="mc-next-hero__pulse" /> THE DATA OPERATIONS PLATFORM <span className="mc-next-hero__eyebrow-rule" /></div>
        <h1>Better data.<br /><span>Brighter decisions.</span></h1>
        <p>Bring clarity to complex data. Clean spreadsheets, automate Excel workflows, and reconcile records in one powerful, browser-first workspace.</p>
        <div className="mc-next-hero__actions"><button type="button" onClick={props.onCleanFile}>Start cleaning for free <Icon name="arrow" size={19} /></button><a href="#free-tools">Explore the platform <Icon name="arrow" size={17} /></a></div>
        <div className="mc-next-hero__proof"><span><Icon name="shield" size={17} /> Processed in your browser</span><span>No sign-up required</span><span>Free to get started</span></div>
      </div>
      <div className="mc-next-hero__visual"><div className="mc-next-hero__orbit orbit-one" aria-hidden="true" /><div className="mc-next-hero__orbit orbit-two" aria-hidden="true" /><WorkspacePreview {...props} /><div className="mc-next-hero__float"><span className="mc-next-hero__float-icon"><Icon name="clean" size={18} /></span><span><strong>Data, made dependable</strong><small>Clean · Validate · Reconcile</small></span><span className="mc-next-hero__float-check">✓</span></div></div>
    </div>
    <div className="mc-next-hero__footer"><div className="mc-next-hero__footer-inner"><span className="mc-next-hero__footer-label">EXPLORE YOUR WORKSPACE</span><nav aria-label="Product workspaces" className="mc-next-hero__workspaces">{workspaces.map((item) => <a key={item.num} href={item.href}><span>{item.num}</span><strong>{item.label}</strong><small>{item.description}</small><Icon name="arrow" size={15} /></a>)}</nav></div></div>
  </section>;
}
