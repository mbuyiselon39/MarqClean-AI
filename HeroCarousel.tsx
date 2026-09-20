import { useEffect, useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onSampleLeads: () => void;
  onOpenReconciliation: () => void;
};

function ProductPreview() {
  return (
    <div className="mc-dark-product">
      <div className="mc-dark-product__bar">
        <div className="mc-dark-product__brand"><span className="mc-dark-mini">M</span> MarqClean AI</div>
        <span className="mc-dark-product__pill">LIVE DATA</span>
      </div>
      <div className="mc-dark-product__body">
        <aside>
          {["Dashboard","Smart Drop","Data Cleaner","Enrichment","Workflows","Export"].map((item, i) => <div key={item} className={i === 1 ? "is-active" : ""}><span>{["⌂","◈","⌁","◇","⌘","↓"][i]}</span>{item}</div>)}
        </aside>
        <div className="mc-dark-product__main">
          <div className="mc-dark-product__title"><span>Smart Drop</span><strong>Ready for your data.</strong></div>
          <div className="mc-dark-drop">↥<strong>Drag & drop your file here</strong><span>or click to browse · CSV, XLSX, PDF</span></div>
          <div className="mc-dark-actions">
            {["Clean & Format","Enrich with AI","Run Workflows","Export Clean Files"].map((x,i)=><div key={x}><span>{["⌁","◎","⌘","↓"][i]}</span><b>{x}</b><em>→</em></div>)}
          </div>
        </div>
        <div className="mc-dark-security"><span>♢</span><strong>100% Secure</strong><small>Client-side execution</small><p>Your files never leave this browser.</p></div>
      </div>
    </div>
  );
}

export default function HeroCarousel({ onCleanFile, onSampleLeads, onOpenReconciliation }: HeroActions) {
  const [active, setActive] = useState(0);
  const tabs = ["Data Cleaning","Excel Automation","Reconciliation","Data Toolbox"];
  useEffect(() => { const t = window.setInterval(() => setActive(v => (v + 1) % tabs.length), 6500); return () => window.clearInterval(t); }, []);
  const actions = [
    () => onCleanFile(),
    () => onCleanFile(),
    () => onOpenReconciliation(),
    () => onSampleLeads(),
  ];
  return (
    <section className="mc-hex-hero" aria-label="MarqClean AI data automation">
      <div className="mc-hex-hero__grid" />
      <div className="mc-hex-hero__spreadsheet" />
      <div className="mc-hex-container">
        <div className="mc-hex-copy">
          <div className="mc-hex-eyebrow"><span /> AI-POWERED DATA CLEANING</div>
          <h1>Turn Messy Spreadsheets<br />into <em>Clean, Powerful Data</em></h1>
          <p>MarqClean AI is the all-in-one, serverless platform that cleans, enriches, and transforms your data — right in your browser. No servers. No uploads. Just powerful AI tools, complete privacy, and professional-grade results.</p>
          <div className="mc-hex-benefits">
            <span>⌁ <b>Clean & Format</b><small>Data Automatically</small></span>
            <span>◎ <b>Enrich with AI</b><small>Web Scrapers</small></span>
            <span>⌘ <b>Automate Workflows</b><small>& Bulk Processing</small></span>
            <span>♢ <b>100% Client-Side</b><small>& Secure</small></span>
          </div>
          <div className="mc-hex-actions">
            <button onClick={onCleanFile}>Get Started Free <span>→</span></button>
            <button className="secondary" onClick={onSampleLeads}>▷ &nbsp; Watch Demo</button>
          </div>
        </div>
        <div className="mc-hex-stage">
          <ProductPreview />
        </div>
      </div>
      <div className="mc-hex-tabs" role="tablist">
        {tabs.map((tab, i) => <button key={tab} className={active === i ? "is-active" : ""} onClick={() => { setActive(i); actions[i](); }} role="tab" aria-selected={active === i}>{tab}</button>)}
      </div>
    </section>
  );
}
