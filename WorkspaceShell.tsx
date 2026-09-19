import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";

// ---------------------------------------------------------------------------
// Shared enterprise workspace shell: unified dark theme, homepage global
// header, spreadsheet grid background, and a standardized module page header.
// Used by every internal tool so the platform feels like one product.
// ---------------------------------------------------------------------------

function go(path: string) {
  window.history.pushState({}, "", path);
  window.dispatchEvent(new PopStateEvent("popstate"));
}

const SIMPLE_NAV: Array<{ label: string; path: string }> = [
  { label: "Home", path: "/#top" },
  { label: "Features", path: "/#features" },
  { label: "Workflow", path: "/#workflow" },
];

const PRODUCTS_MENU: Array<{ label: string; path: string }> = [
  { label: "Quick Data & CSV Cleaner", path: "/data-cleaner" },
  { label: "Excel Automation", path: "/excel-automation" },
  { label: "Reconciliation Hub", path: "/reconciliation-hub" },
  { label: "Data Toolbox", path: "/data-toolbox" },
];

export function GlobalHeader() {
  const [isMobileNavOpen, setIsMobileNavOpen] = useState(false);
  const [isProductsOpen, setIsProductsOpen] = useState(false);
  const [isMobileProductsOpen, setIsMobileProductsOpen] = useState(false);
  const productsMenuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isProductsOpen) return;
    function handleOutside(event: MouseEvent) {
      if (productsMenuRef.current && !productsMenuRef.current.contains(event.target as Node)) {
        setIsProductsOpen(false);
      }
    }
    function handleKey(event: KeyboardEvent) {
      if (event.key === "Escape") setIsProductsOpen(false);
    }
    document.addEventListener("mousedown", handleOutside);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handleOutside);
      document.removeEventListener("keydown", handleKey);
    };
  }, [isProductsOpen]);

  return (
    <>
      <a
        className="sr-only focus:not-sr-only focus:fixed focus:left-4 focus:top-4 focus:z-50 focus:rounded-full focus:bg-white focus:px-4 focus:py-2 focus:text-sm focus:font-semibold focus:text-slate-950"
        href="#main-content"
      >
        Skip to content
      </a>
      <header className="ws-surface fixed inset-x-0 top-0 z-40 border-b border-white/10">
      <nav className="mx-auto flex max-w-[100rem] items-center justify-between gap-4 px-5 py-3.5 lg:px-8" aria-label="Primary navigation">
        <a
          href="/"
          className="mc-display flex items-center gap-2 text-lg font-extrabold tracking-tight text-white"
          onClick={(e) => { e.preventDefault(); go("/"); }}
        >
          <span className="inline-flex h-6 w-6 items-center justify-center rounded-md bg-gradient-to-br from-[#0D9488] to-[#14B8A6] text-xs font-black text-[#04121a]">M</span>
          MarqClean AI
        </a>
        <div className="hidden items-center gap-6 text-[13px] font-medium text-[#b0bacb] xl:flex">
          <a href={SIMPLE_NAV[0].path} className="transition hover:text-[#0D9488]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[0].path); }}>{SIMPLE_NAV[0].label}</a>
          <a href={SIMPLE_NAV[1].path} className="transition hover:text-[#0D9488]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[1].path); }}>{SIMPLE_NAV[1].label}</a>
          <a href={SIMPLE_NAV[2].path} className="transition hover:text-[#0D9488]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[2].path); }}>{SIMPLE_NAV[2].label}</a>
          <div className="relative" ref={productsMenuRef}>
            <button
              type="button"
              className="flex items-center gap-1 transition hover:text-[#0D9488] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0D9488]/50"
              aria-haspopup="true"
              aria-expanded={isProductsOpen}
              aria-controls="ws-products-menu"
              onClick={() => setIsProductsOpen((open) => !open)}
            >
              Products
              <svg className={`h-3.5 w-3.5 transition ${isProductsOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
            </button>
            {isProductsOpen ? (
              <div
                id="ws-products-menu"
                role="menu"
                aria-label="Products"
                className="absolute left-1/2 top-full mt-3 w-72 -translate-x-1/2 rounded-2xl border border-white/10 bg-[#0a0b16] p-2 text-sm shadow-2xl"
              >
                {PRODUCTS_MENU.map((item) => (
                  <a
                    key={item.label}
                    role="menuitem"
                    href={item.path}
                    className="block rounded-lg px-3 py-2.5 font-medium text-white/80 transition hover:bg-white/5 hover:text-white focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0D9488]/50"
                    onClick={(e) => { e.preventDefault(); setIsProductsOpen(false); go(item.path); }}
                  >
                    {item.label}
                  </a>
                ))}
              </div>
            ) : null}
          </div>
          <a href="/contact" className="transition hover:text-[#0D9488]" onClick={(e) => { e.preventDefault(); go("/contact"); }}>Contact</a>
        </div>
        <div className="flex items-center gap-2">
          <button
            className="ws-btn-primary hidden rounded-full px-5 py-2 text-sm font-semibold sm:inline-flex"
            onClick={() => go("/#cleaner")}
          >
            Start Free
          </button>
          <button
            type="button"
            className="inline-flex h-10 w-10 items-center justify-center rounded-full border border-white/15 text-white transition hover:border-[#0D9488]/60 hover:text-[#0D9488] xl:hidden"
            aria-label={isMobileNavOpen ? "Close navigation menu" : "Open navigation menu"}
            aria-expanded={isMobileNavOpen}
            aria-controls="ws-mobile-nav-panel"
            onClick={() => setIsMobileNavOpen((open) => !open)}
          >
            {isMobileNavOpen ? (
              <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
                <path d="M6 6l12 12M18 6L6 18" />
              </svg>
            ) : (
              <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
                <path d="M4 6h16M4 12h16M4 18h16" />
              </svg>
            )}
          </button>
        </div>
      </nav>
      {isMobileNavOpen ? (
        <div id="ws-mobile-nav-panel" className="border-t border-white/10 bg-[#0a0b16]/98 px-5 py-4 xl:hidden">
          <div className="flex flex-col gap-1 text-[15px] font-medium text-[#b0bacb]">
            {SIMPLE_NAV.map((item) => (
              <a
                key={item.label}
                href={item.path}
                className="rounded-lg px-3 py-2.5 transition hover:bg-white/5 hover:text-[#0D9488]"
                onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); go(item.path); }}
              >
                {item.label}
              </a>
            ))}
            <button
              type="button"
              className="flex items-center justify-between rounded-lg px-3 py-2.5 text-left transition hover:bg-white/5 hover:text-[#0D9488]"
              aria-expanded={isMobileProductsOpen}
              aria-controls="ws-mobile-products-menu"
              onClick={() => setIsMobileProductsOpen((open) => !open)}
            >
              Products
              <svg className={`h-4 w-4 transition ${isMobileProductsOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
            </button>
            {isMobileProductsOpen ? (
              <div id="ws-mobile-products-menu" className="ml-3 flex flex-col gap-1 border-l border-white/10 pl-3">
                {PRODUCTS_MENU.map((item) => (
                  <a
                    key={item.label}
                    href={item.path}
                    className="rounded-lg px-3 py-2 text-sm text-white/70 transition hover:bg-white/5 hover:text-white"
                    onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); setIsMobileProductsOpen(false); go(item.path); }}
                  >
                    {item.label}
                  </a>
                ))}
              </div>
            ) : null}
            <a
              className="rounded-lg px-3 py-2.5 transition hover:bg-white/5 hover:text-[#0D9488]"
              href="/contact"
              onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); go("/contact"); }}
            >
              Contact
            </a>
            <button
              className="ws-btn-primary mt-2 rounded-full px-5 py-2.5 text-sm font-semibold sm:hidden"
              onClick={() => { setIsMobileNavOpen(false); go("/#cleaner"); }}
            >
              Start Free
            </button>
          </div>
        </div>
      ) : null}
    </header>
    </>
  );
}

export type StatusTone = "idle" | "processing" | "success" | "warning" | "error";

const TONE_STYLES: Record<StatusTone, string> = {
  idle: "bg-white/5 text-[#b0bacb] border-white/12",
  processing: "bg-[#0D9488]/10 text-[#5EEAD4] border-[#0D9488]/40",
  success: "bg-[#FB6F58]/12 text-[#FFCCC2] border-[#FB6F58]/40",
  warning: "bg-[#f59e0b]/12 text-[#fbbf24] border-[#f59e0b]/40",
  error: "bg-[#ef4444]/12 text-[#f87171] border-[#ef4444]/40",
};

export function StatusBadge({ tone, label }: { tone: StatusTone; label: string }) {
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full border px-3 py-1 text-xs font-semibold ${TONE_STYLES[tone]}`}>
      <span className={`h-1.5 w-1.5 rounded-full ${tone === "processing" ? "animate-pulse bg-[#0D9488]" : tone === "success" ? "bg-[#FB6F58]" : tone === "warning" ? "bg-[#f59e0b]" : tone === "error" ? "bg-[#ef4444]" : "bg-[#b0bacb]"}`} />
      {label}
    </span>
  );
}

export function FormatBadge({ label }: { label: string }) {
  return <span className="ws-badge inline-flex items-center gap-1 rounded-md px-2 py-0.5 text-[11px] font-semibold uppercase tracking-wide">{label}</span>;
}

export function WorkspaceShell({
  icon,
  title,
  description,
  engine,
  formats,
  status,
  children,
}: {
  icon: string;
  title: string;
  description: string;
  engine?: string;
  formats?: string[];
  status?: { tone: StatusTone; label: string };
  children: ReactNode;
}) {
  return (
    <div className="ws-root ws-scope min-h-screen">
      <GlobalHeader />
      <div className="ws-grid pointer-events-none absolute inset-x-0 top-0 h-[38rem]" aria-hidden="true" />

      <div className="relative mx-auto max-w-[110rem] px-5 pb-20 pt-24 lg:px-8">
        {/* Module page header */}
        <div className="ws-surface flex flex-wrap items-start justify-between gap-4 rounded-2xl p-6">
          <div className="flex items-start gap-4">
            <span className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-[#0D9488]/20 to-[#14B8A6]/20 text-2xl ring-1 ring-inset ring-white/10">{icon}</span>
            <div>
              <h1 className="mc-display text-2xl font-bold tracking-tight text-white sm:text-3xl">{title}</h1>
              <p className="mt-1 max-w-3xl text-justify text-sm leading-6 text-[#b0bacb] [hyphens:auto]">{description}</p>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                {formats?.map((f) => <FormatBadge key={f} label={f} />)}
                {engine ? <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-semibold text-[#5EEAD4]">Engine: {engine}</span> : null}
              </div>
            </div>
          </div>
          {status ? <StatusBadge tone={status.tone} label={status.label} /> : null}
        </div>

        <div className="mt-6">{children}</div>
      </div>
    </div>
  );
}

// Responsive tool launchpad grid: shows every tool as a card, no scrolling to
// discover them. Desktop 4-col, tablet 2-col, mobile 1-col.
export function ToolLaunchpad<T extends string>({
  tools,
  active,
  onSelect,
}: {
  tools: Array<{ key: T; label: string; blurb: string; icon: string }>;
  active: T;
  onSelect: (key: T) => void;
}) {
  return (
    <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
      {tools.map((t) => (
        <button
          key={t.key}
          onClick={() => onSelect(t.key)}
          className={`group flex items-start gap-3 rounded-xl border p-4 text-left transition ${active === t.key ? "border-[#0D9488]/60 bg-[#0D9488]/10" : "border-white/8 bg-white/[0.03] hover:border-[#0D9488]/40 hover:bg-white/[0.06]"}`}
        >
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-white/5 text-lg ring-1 ring-inset ring-white/10">{t.icon}</span>
          <span className="min-w-0">
            <span className="block truncate text-sm font-semibold text-white">{t.label}</span>
            <span className="mt-0.5 block text-xs leading-5 text-[#8a94a8]">{t.blurb}</span>
          </span>
        </button>
      ))}
    </div>
  );
}

export function EmptyState({ title, note }: { title: string; note?: string }) {
  return (
    <div className="ws-surface rounded-2xl border-dashed p-10 text-center">
      <div className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-xl bg-white/5 text-2xl ring-1 ring-inset ring-white/10">🗂️</div>
      <h3 className="mc-display text-lg font-semibold text-white">{title}</h3>
      {note ? <p className="mx-auto mt-2 max-w-md text-sm text-[#b0bacb]">{note}</p> : null}
      <div className="mt-4 flex flex-wrap justify-center gap-2 text-xs text-[#b0bacb]">
        <span>✓ Data Validation</span>
        <span>✓ Reconciliation</span>
        <span>✓ Cleansing</span>
        <span>✓ Analysis</span>
      </div>
    </div>
  );
}
