import { useState } from "react";

type HeroActions = {
  onCleanFile: () => void;
  onOpenReconciliation: () => void;
  onViewTools: () => void;
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

function WorkspacePreview() {
  return (
    <div className="mc-linear-preview" aria-label="MarqClean AI workspace preview">
      <div className="mc-linear-preview__chrome">
        <div className="mc-linear-preview__brand"><span className="mc-linear-dot" /> MarqClean AI</div>
        <div className="mc-linear-preview__search"><kbd>⌘K</kbd><span>Search tools, files, formulas</span></div>
        <div className="mc-linear-preview__avatar">M</div>
      </div>
      <div className="mc-linear-preview__body">
        <aside>
          <small>WORKSPACE</small>
          {[
            ["Overview", "home"],
            ["Smart Drop", "upload"],
            ["Cleaners", "clean"],
            ["Formatters", "format"],
            ["Reconciliation", "reconcile"],
          ].map(([item, icon], i) => (
            <div key={item} className={i === 1 ? "is-active" : ""}><span><LineIcon type={icon as "home" | "upload" | "clean" | "format" | "reconcile"} /></span>{item}</div>
          ))}
          <small className="mc-linear-preview__group">TOOLS</small>
          {[
            ["Excel Functions", "formula"],
            ["Data Tools", "format"],
          ].map(([item, icon]) => <div key={item}><span><LineIcon type={icon as "formula" | "format"} /></span>{item}</div>)}
        </aside>
        <main>
          <div className="mc-linear-preview__heading">
            <div><small>SMART DROP</small><strong>Ready for your data.</strong></div>
            <span className="mc-local-pill">Processed locally</span>
          </div>
          <div className="mc-linear-preview__drop">
            <span className="mc-drop-icon"><LineIcon type="upload" /></span>
            <strong>Drop CSV, Excel or PDF</strong>
            <small>or choose a file from your computer</small>
          </div>
          <div className="mc-linear-preview__cards">
            {[
              ["Clean & format", "clean"],
              ["Reconcile files", "reconcile"],
              ["Run formulas", "formula"],
            ].map(([label, icon]) => (
              <div key={label}><span><LineIcon type={icon as "clean" | "reconcile" | "formula"} /></span><b>{label}</b><em><LineIcon type="arrow" /></em></div>
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
