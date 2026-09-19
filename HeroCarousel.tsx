import { useEffect, useRef, useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onSampleLeads: () => void;
  onOpenReconciliation: () => void;
};

function ProductPreview() {
  return (
    <div className="amzigo-product-preview">
      <div className="amzigo-preview-top">
        <div className="amzigo-preview-brand"><span className="amzigo-mini-mark">M</span> MarqClean AI</div>
        <span className="amzigo-preview-status">LIVE DATA QUALITY</span>
      </div>
      <div className="amzigo-preview-body">
        <aside>
          <div className="amzigo-side-active">Overview</div>
          <div>Data Cleaner</div>
          <div>Excel Automation</div>
          <div>Reconciliation</div>
          <div>Academy</div>
        </aside>
        <div className="amzigo-dashboard">
          <div className="amzigo-dash-head">
            <div><span>Good morning</span><strong>Your data, under control.</strong></div>
            <button>Export report</button>
          </div>
          <div className="amzigo-stat-grid">
            <div><span>Records cleaned</span><strong>24,680</strong><small>+18.4% this month</small></div>
            <div><span>Duplicates removed</span><strong>3,842</strong><small>98.7% confidence</small></div>
            <div><span>Files processed</span><strong>1,284</strong><small>All browser-first</small></div>
          </div>
          <div className="amzigo-chart-card">
            <div><div><span>Data quality score</span><strong>96.8%</strong></div><em>Last 30 days</em></div>
            <div className="amzigo-bars">
              {[44,58,51,72,68,84,78,92,88,96,91,98].map((h,i)=><span key={i} style={{height:h+"%"}} />)}
            </div>
          </div>
          <div className="amzigo-table-card">
            <div className="amzigo-table-title"><strong>Recent files</strong><span>View all</span></div>
            <div className="amzigo-row"><span>Q3_Leads.xlsx</span><b>Cleaned</b><small>4,821 rows</small></div>
            <div className="amzigo-row"><span>Bank_Statement.csv</span><b>Reconciled</b><small>2,104 rows</small></div>
            <div className="amzigo-row"><span>CRM_Export.csv</span><b>Validated</b><small>8,936 rows</small></div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function HeroCarousel({ onCleanFile, onSampleLeads, onOpenReconciliation }: HeroActions) {
  const [active, setActive] = useState(0);
  const rootRef = useRef<HTMLDivElement>(null);

  const tabs = [
    { label: "Data Cleaning", title: "Clean your data. Grow with confidence.", text: "Clean, standardise, validate and transform Excel and CSV files in seconds — without sending your data away.", cta: "Start Cleaning Free" },
    { label: "Excel Automation", title: "Automate the spreadsheet work.", text: "Turn repetitive Excel tasks into reliable workflows with formulas, formatting, conversions and intelligent data preparation.", cta: "Automate Excel" },
    { label: "Reconciliation", title: "Reconcile records without the manual grind.", text: "Compare statements, ledgers and master files, surface exceptions and export review-ready reconciliation results.", cta: "Open Reconciliation" },
    { label: "Data Toolbox", title: "Everything your data team needs, in one place.", text: "A practical browser-first toolkit for cleaning, conversion, analysis, document processing and everyday spreadsheet work.", cta: "Explore Data Toolbox" },
  ];

  useEffect(() => {
    const timer = window.setInterval(() => setActive((value) => (value + 1) % tabs.length), 7000);
    return () => window.clearInterval(timer);
  }, [tabs.length]);

  const current = tabs[active];

  return (
    <section ref={rootRef} className="amzigo-hero">
      <div className="amzigo-hero-orb amzigo-orb-one" />
      <div className="amzigo-hero-orb amzigo-orb-two" />
      <div className="amzigo-container amzigo-hero-inner">
        <div className="amzigo-eyebrow"><span /> AI-powered data operations for modern teams</div>
        <h1>{current.title}</h1>
        <p className="amzigo-hero-copy">{current.text}</p>
        <div className="amzigo-hero-actions">
          <button className="amzigo-btn-primary" onClick={onCleanFile}>{current.cta}<span>→</span></button>
          <button className="amzigo-btn-secondary" onClick={onSampleLeads}>Try a sample</button>
        </div>
        <div className="amzigo-trust-row">
          <span>✓ Browser-first processing</span>
          <span>✓ CSV & Excel ready</span>
          <span>✓ No credit card required</span>
        </div>
        <div className="amzigo-tabs" role="tablist" aria-label="MarqClean AI capabilities">
          {tabs.map((tab, index) => (
            <button key={tab.label} role="tab" aria-selected={active === index} className={active === index ? "is-active" : ""} onClick={() => {
              setActive(index);
              if (index === 2) onOpenReconciliation();
            }}>
              <span>{tab.label}</span>
              <i />
            </button>
          ))}
        </div>
        <div className="amzigo-product-stage">
          <div className="amzigo-stage-glow" />
          <ProductPreview />
        </div>
      </div>
    </section>
  );
}
