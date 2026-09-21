import { useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onOpenReconciliation: () => void;
  onViewTools: () => void;
};

function WorkspacePreview() {
  return (
    <div className="mc-linear-preview" aria-label="MarqClean AI workspace preview">
      <div className="mc-linear-preview__chrome">
        <div className="mc-linear-preview__brand"><span className="mc-linear-dot" /> MarqClean AI</div>
        <div className="mc-linear-preview__search">⌘ K <span>Search tools, files, formulas…</span></div>
        <div className="mc-linear-preview__avatar">M</div>
      </div>
      <div className="mc-linear-preview__body">
        <aside>
          <small>WORKSPACE</small>
          {["Overview", "Smart Drop", "Cleaners", "Formatters", "Reconciliation"].map((item, i) => (
            <div key={item} className={i === 1 ? "is-active" : ""}><span>{["⌂", "↥", "✦", "◫", "⇄"][i]}</span>{item}</div>
          ))}
          <small className="mc-linear-preview__group">TOOLS</small>
          {["Excel Functions", "AI Tools"].map(item => <div key={item}><span>·</span>{item}</div>)}
        </aside>
        <main>
          <div className="mc-linear-preview__heading">
            <div><small>SMART DROP</small><strong>Ready for your data.</strong></div>
            <span className="mc-local-pill">● Processed locally</span>
          </div>
          <div className="mc-linear-preview__drop">
            <span className="mc-drop-icon">↥</span>
            <strong>Drop CSV, Excel or PDF</strong>
            <small>or choose a file from your computer</small>
          </div>
          <div className="mc-linear-preview__cards">
            {["Clean & format", "Reconcile files", "Run formulas"].map((x, i) => (
              <div key={x}><span>{["✦", "⇄", "ƒx"][i]}</span><b>{x}</b><em>→</em></div>
            ))}
          </div>
        </main>
      </div>
    </div>
  );
}

export default function HeroCarousel({ onCleanFile }: HeroActions) {
  const [active, setActive] = useState(0);
  const tabs = [
    { label: "Data Cleaning", href: "#cleaner" },
    { label: "Excel Automation", href: "/excel-automation" },
    { label: "Reconciliation", href: "/reconciliation-hub" },
    { label: "Data Toolbox", href: "/data-toolbox" },
  ];

  return (
    <section className="mc-minimal-hero" aria-label="MarqClean AI data preparation workspace">
      <div className="mc-minimal-hero__grid" aria-hidden="true" />
      <div className="mc-minimal-hero__wireframe" aria-hidden="true">
        <svg viewBox="0 0 900 520" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M50 420H850M50 350H850M50 280H850M50 210H850M50 140H850" stroke="currentColor" strokeOpacity=".08" />
          <path d="M90 430V100M230 430V100M370 430V100M510 430V100M650 430V100M790 430V100" stroke="currentColor" strokeOpacity=".08" />
          <path d="M90 360L230 310L370 335L510 235L650 265L790 155" stroke="currentColor" strokeOpacity=".18" strokeWidth="2" />
          <circle cx="790" cy="155" r="5" fill="currentColor" fillOpacity=".5" />
        </svg>
      </div>
      <div className="mc-minimal-hero__inner">
        <div className="mc-minimal-hero__copy">
          <div className="mc-minimal-hero__eyebrow"><span /> Browser-local data workspace</div>
          <h1>Clean data.<br /><em>Clear decisions.</em></h1>
          <p className="mc-hero-copy">Clean, transform, reconcile and analyse CSV, Excel and PDF data without sending your files to a server.</p>
          <div className="mc-minimal-hero__actions">
            <button onClick={onCleanFile}>Launch Workspace <span>↗</span></button>
            <a className="secondary" href="#free-tools">33 free tools</a>
          </div>
          <div className="mc-minimal-hero__trust"><span>✓</span> Processed locally in your browser <i /> No upload required</div>
        </div>
        <div className="mc-minimal-hero__product" aria-label="MarqClean AI workspace preview"><WorkspacePreview /></div>
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
