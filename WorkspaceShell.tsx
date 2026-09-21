import { useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import Logo from "./Logo";
import ThemeToggle from "./ThemeToggle";

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
  { label: "Platform", path: "/#features" },
  { label: "How it works", path: "/#workflow" },
];

const PRODUCTS_MENU: Array<{ label: string; path: string }> = [
  { label: "Quick Data & CSV Cleaner", path: "/data-cleaner" },
  { label: "Local Data Engine", path: "/data-engine" },
  { label: "Excel Automation", path: "/excel-automation" },
  { label: "Reconciliation Hub", path: "/reconciliation-hub" },
  { label: "Data Toolbox", path: "/data-toolbox" },
];

export function GlobalHeader() {
  const [isMobileNavOpen, setIsMobileNavOpen] = useState(false);
  const [isWorkspacesOpen, setIsWorkspacesOpen] = useState(false);
  const [isMobileWorkspacesOpen, setIsMobileWorkspacesOpen] = useState(false);
  const productsMenuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isWorkspacesOpen) return;
    function handleOutside(event: MouseEvent) {
      if (productsMenuRef.current && !productsMenuRef.current.contains(event.target as Node)) {
        setIsWorkspacesOpen(false);
      }
    }
    function handleKey(event: KeyboardEvent) {
      if (event.key === "Escape") setIsWorkspacesOpen(false);
    }
    document.addEventListener("mousedown", handleOutside);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handleOutside);
      document.removeEventListener("keydown", handleKey);
    };
  }, [isWorkspacesOpen]);

  return (
    <>
      <a
        className="sr-only focus:not-sr-only focus:fixed focus:left-4 focus:top-4 focus:z-50 focus:rounded-md focus:bg-canvas focus:px-4 focus:py-2 focus:text-sm focus:font-medium focus:text-ink"
        href="#main-content"
      >
        Skip to content
      </a>
      <header className="ws-surface ws-global-header fixed inset-x-0 top-0 z-40 border-b">
      <nav className="mx-auto flex max-w-[100rem] items-center justify-between gap-4 px-5 py-3.5 lg:px-8" aria-label="Primary navigation">
        <a
          href="/"
          className="mc-display flex items-center gap-2 text-lg font-medium tracking-tight text-ink"
          onClick={(e) => { e.preventDefault(); go("/"); }}
        >
          <Logo className="shrink-0" />
        </a>
        <div className="hidden items-center gap-6 text-[13px] font-medium text-[rgb(var(--ink-2))] xl:flex">
          <a href={SIMPLE_NAV[0].path} className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[0].path); }}>{SIMPLE_NAV[0].label}</a>
          <a href={SIMPLE_NAV[1].path} className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[1].path); }}>{SIMPLE_NAV[1].label}</a>
          <a href={SIMPLE_NAV[2].path} className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); go(SIMPLE_NAV[2].path); }}>{SIMPLE_NAV[2].label}</a>
          <div className="relative" ref={productsMenuRef}>
            <button
              type="button"
              className="flex items-center gap-1 transition hover:text-[rgb(var(--accent))] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[rgb(var(--accent))]/50"
              aria-haspopup="true"
              aria-expanded={isWorkspacesOpen}
              aria-controls="ws-products-menu"
              onClick={() => setIsWorkspacesOpen((open) => !open)}
            >Workspaces<svg className={`h-3.5 w-3.5 transition ${isWorkspacesOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
            </button>
            {isWorkspacesOpen ? (
              <div
                id="ws-workspaces-menu"
                role="menu"
                aria-label="Workspaces"
                className="absolute left-1/2 top-full mt-3 w-80 -translate-x-1/2 rounded-lg border border-line bg-canvas p-2 text-sm"
              >
                {PRODUCTS_MENU.map((item) => (
                  <a
                    key={item.label}
                    role="menuitem"
                    href={item.path}
                    className="block rounded-lg px-3 py-2.5 font-medium text-ink-2 transition hover:bg-surface hover:text-[rgb(var(--accent))] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[rgb(var(--accent))]/30"
                    onClick={(e) => { e.preventDefault(); setIsWorkspacesOpen(false); go(item.path); }}
                  >
                    {item.label}
                  </a>
                ))}
              </div>
            ) : null}
          </div>
          <a href="/contact" className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); go("/contact"); }}>Contact</a>
        </div>
        <div className="flex items-center gap-2">
          <ThemeToggle compact />
          <button
            className="ws-btn-primary hidden rounded-md px-5 py-2 text-sm font-medium sm:inline-flex"
            onClick={() => go("/#cleaner")}
          >
            Start Free
          </button>
          <button
            type="button"
            className="inline-flex h-10 w-10 items-center justify-center rounded-md border border-line text-ink-2 transition hover:border-[rgb(var(--accent))]/60 hover:text-[rgb(var(--accent))] xl:hidden"
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
        <div id="ws-mobile-nav-panel" className="border-t border-line bg-canvas px-5 py-4 xl:hidden">
          <div className="flex flex-col gap-1 text-[15px] font-medium text-ink-2">
            {SIMPLE_NAV.map((item) => (
              <a
                key={item.label}
                href={item.path}
                className="rounded-lg px-3 py-2.5 transition hover:bg-surface hover:text-[rgb(var(--accent))]"
                onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); go(item.path); }}
              >
                {item.label}
              </a>
            ))}
            <button
              type="button"
              className="flex items-center justify-between rounded-lg px-3 py-2.5 text-left transition hover:bg-canvas/5 hover:text-[rgb(var(--accent))]"
              aria-expanded={isMobileWorkspacesOpen}
              aria-controls="ws-mobile-products-menu"
              onClick={() => setIsMobileWorkspacesOpen((open) => !open)}
            >Workspaces<svg className={`h-4 w-4 transition ${isMobileWorkspacesOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
            </button>
            {isMobileWorkspacesOpen ? (
              <div id="ws-mobile-products-menu" className="ml-3 flex flex-col gap-1 border-l border-line pl-3">
                {PRODUCTS_MENU.map((item) => (
                  <a
                    key={item.label}
                    href={item.path}
                    className="rounded-lg px-3 py-2 text-sm text-ink-2 transition hover:bg-surface hover:text-[rgb(var(--accent))]"
                    onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); setIsMobileWorkspacesOpen(false); go(item.path); }}
                  >
                    {item.label}
                  </a>
                ))}
              </div>
            ) : null}
            <a
              className="rounded-lg px-3 py-2.5 transition hover:bg-canvas/5 hover:text-[rgb(var(--accent))]"
              href="/contact"
              onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); go("/contact"); }}
            >
              Contact
            </a>
            <button
              className="ws-btn-primary mt-2 rounded-md px-5 py-2.5 text-sm font-medium sm:hidden"
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
  idle: "bg-canvas/5 text-[rgb(var(--ink-2))] border-line",
  processing: "bg-[rgb(var(--accent))]/10 text-[rgb(var(--accent-hover))] border-[rgb(var(--accent))]/40",
  success: "bg-[rgb(var(--success))]/10 text-[rgb(var(--success))] border-[rgb(var(--success))]/30",
  warning: "bg-[rgb(var(--ink-3))]/12 text-[rgb(var(--ink-3))] border-[rgb(var(--ink-3))]/40",
  error: "bg-[rgb(var(--error))]/12 text-[rgb(var(--error))] border-[rgb(var(--error))]/40",
};

export function StatusBadge({ tone, label }: { tone: StatusTone; label: string }) {
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-md border px-3 py-1 text-xs font-medium ${TONE_STYLES[tone]}`}>
      <span className={`h-1.5 w-1.5 rounded-md ${tone === "processing" ? "animate-pulse bg-[rgb(var(--accent))]" : tone === "success" ? "bg-[rgb(var(--success))]" : tone === "warning" ? "bg-[rgb(var(--ink-3))]" : tone === "error" ? "bg-[rgb(var(--error))]" : "bg-[rgb(var(--ink-2))]"}`} />
      {label}
    </span>
  );
}

export function FormatBadge({ label }: { label: string }) {
  return <span className="ws-badge inline-flex items-center gap-1 rounded-md px-2 py-0.5 text-[11px] font-medium uppercase tracking-wide">{label}</span>;
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
        <div className="ws-surface ws-module-header flex flex-wrap items-start justify-between gap-4 rounded-lg p-6">
          <div className="flex items-start gap-4">
            <span className="flex h-12 w-12 shrink-0 items-center justify-center rounded-lg bg-canvas /20 /20 text-2xl ring-1 ring-inset ring-white/10">{icon}</span>
            <div>
              <h1 className="mc-display text-2xl font-medium tracking-tight text-ink sm:text-3xl">{title}</h1>
              <p className="mt-1 max-w-3xl text-justify text-sm leading-6 text-[rgb(var(--ink-2))] [hyphens:auto]">{description}</p>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                {formats?.map((f) => <FormatBadge key={f} label={f} />)}
                {engine ? <span className="ws-badge rounded-md px-2 py-0.5 text-[11px] font-medium text-[rgb(var(--accent-hover))]">Engine: {engine}</span> : null}
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

// ---------------------------------------------------------------------------
// Unified workspace navigation: the former launchpad is now a real product
// sidebar with persistent collapse state, grouped tools and a global command
// palette. The existing tool contract remains compatible with every module.
// ---------------------------------------------------------------------------

type WorkspaceTool<T extends string> = {
  key: T;
  label: string;
  blurb: string;
  icon: string;
  category?: string;
};

function toolGroup(tool: WorkspaceTool<string>) {
  if (tool.category) return tool.category;
  const value = tool.key + " " + tool.label;
  if (/(bank|reconcil|statement|fund|verify|comparison)/i.test(value)) return "Control";
  if (/(formula|excel|xlookup|sumifs|gpt|automation)/i.test(value)) return "Automation";
  if (/(chart|statistic|analyst|analysis|toolbox)/i.test(value)) return "Intelligence";
  return "Data preparation";
}

function SidebarIcon({ children }: { children: ReactNode }) {
  return <span className="ws-sidebar-icon" aria-hidden="true">{children}</span>;
}

export function CommandPalette<T extends string>({
  tools,
  active,
  onSelect,
  open,
  onClose,
}: {
  tools: Array<WorkspaceTool<T>>;
  active: T;
  onSelect: (key: T) => void;
  open: boolean;
  onClose: () => void;
}) {
  const [query, setQuery] = useState("");
  const [cursor, setCursor] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  const results = tools.filter((tool) => {
    const haystack = (tool.label + " " + tool.blurb + " " + tool.key + " " + (tool.category ?? "")).toLowerCase();
    return haystack.includes(query.trim().toLowerCase());
  });

  useEffect(() => {
    if (!open) return;
    setQuery("");
    setCursor(0);
    requestAnimationFrame(() => inputRef.current?.focus());
  }, [open]);

  useEffect(() => {
    if (!open) return;
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      } else if (event.key === "ArrowDown") {
        event.preventDefault();
        setCursor((value) => Math.min(value + 1, Math.max(results.length - 1, 0)));
      } else if (event.key === "ArrowUp") {
        event.preventDefault();
        setCursor((value) => Math.max(value - 1, 0));
      } else if (event.key === "Enter" && results[cursor]) {
        event.preventDefault();
        onSelect(results[cursor].key);
        onClose();
      }
    }
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [open, onClose, onSelect, results, cursor]);

  useEffect(() => {
    if (cursor >= results.length) setCursor(Math.max(results.length - 1, 0));
  }, [cursor, results.length]);

  if (!open) return null;

  return (
    <div className="ws-command-overlay" role="presentation" onMouseDown={onClose}>
      <div className="ws-command" role="dialog" aria-modal="true" aria-label="Search workspace" onMouseDown={(event) => event.stopPropagation()}>
        <div className="ws-command__search">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></svg>
          <input ref={inputRef} value={query} onChange={(event) => { setQuery(event.target.value); setCursor(0); }} placeholder="Search tools, workflows and capabilities…" aria-label="Search tools" />
          <kbd>ESC</kbd>
        </div>
        <div className="ws-command__meta"><span>{results.length} workspace tools</span><span>↑ ↓ navigate · Enter open</span></div>
        <div className="ws-command__results" role="listbox" aria-label="Workspace tools">
          {results.map((tool, index) => (
            <button
              key={tool.key}
              type="button"
              role="option"
              aria-selected={active === tool.key}
              className={"ws-command__item " + (index === cursor ? "is-highlighted " : "") + (active === tool.key ? "is-active" : "")}
              onMouseEnter={() => setCursor(index)}
              onClick={() => { onSelect(tool.key); onClose(); }}
            >
              <span className="ws-command__item-icon">{tool.icon}</span>
              <span className="ws-command__item-copy"><strong>{tool.label}</strong><small>{tool.blurb}</small></span>
              <span className="ws-command__item-group">{toolGroup(tool)}</span>
            </button>
          ))}
          {!results.length ? <div className="ws-command__empty">No matching workspace tools. Try “clean”, “reconcile”, “formula” or “analysis”.</div> : null}
        </div>
      </div>
    </div>
  );
}

export function ToolLaunchpad<T extends string>({
  tools,
  active,
  onSelect,
}: {
  tools: Array<WorkspaceTool<T>>;
  active: T;
  onSelect: (key: T) => void;
}) {
  const [collapsed, setCollapsed] = useState(() => {
    try { return window.localStorage.getItem("marqclean:workspace-sidebar") === "collapsed"; } catch { return false; }
  });
  const [paletteOpen, setPaletteOpen] = useState(false);

  useEffect(() => {
    try { window.localStorage.setItem("marqclean:workspace-sidebar", collapsed ? "collapsed" : "expanded"); } catch {}
    document.body.dataset.mcWorkspaceSidebar = collapsed ? "collapsed" : "expanded";
    return () => { delete document.body.dataset.mcWorkspaceSidebar; };
  }, [collapsed]);

  useEffect(() => {
    function handleShortcut(event: KeyboardEvent) {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        setPaletteOpen((open) => !open);
      }
    }
    document.addEventListener("keydown", handleShortcut);
    return () => document.removeEventListener("keydown", handleShortcut);
  }, []);

  const groups = ["Data preparation", "Automation", "Control", "Intelligence"].map((name) => ({
    name,
    items: tools.filter((tool) => toolGroup(tool) === name),
  })).filter((group) => group.items.length);

  return (
    <>
      <aside className={"ws-sidebar " + (collapsed ? "is-collapsed" : "")} aria-label="MarqClean workspace">
        <div className="ws-sidebar__top">
          <button type="button" className="ws-sidebar__search" onClick={() => setPaletteOpen(true)} aria-label="Search workspace">
            <SidebarIcon><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></svg></SidebarIcon>
            <span>Search workspace</span><kbd>⌘K</kbd>
          </button>
          <button type="button" className="ws-sidebar__collapse" onClick={() => setCollapsed((value) => !value)} aria-label={collapsed ? "Expand workspace sidebar" : "Collapse workspace sidebar"} title={collapsed ? "Expand sidebar" : "Collapse sidebar"}>
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d={collapsed ? "m9 18 6-6-6-6" : "m15 18-6-6 6-6"} /></svg>
          </button>
        </div>

        <div className="ws-sidebar__scroll">
          <div className="ws-sidebar__workspace">
            <div className="ws-sidebar__workspace-mark"><Logo markOnly label="MarqClean AI workspace" /></div>
            <div className="ws-sidebar__workspace-copy"><strong>MarqClean AI</strong><small>Unified workspace</small></div>
          </div>

          {groups.map((group) => (
            <section className="ws-sidebar__group" key={group.name}>
              <div className="ws-sidebar__label">{group.name}</div>
              {group.items.map((tool) => (
                <button
                  type="button"
                  key={tool.key}
                  className={"ws-sidebar__item " + (active === tool.key ? "is-active" : "")}
                  onClick={() => onSelect(tool.key)}
                  title={collapsed ? tool.label : undefined}
                  aria-current={active === tool.key ? "page" : undefined}
                >
                  <SidebarIcon>{tool.icon}</SidebarIcon>
                  <span className="ws-sidebar__item-copy"><strong>{tool.label}</strong><small>{tool.blurb}</small></span>
                </button>
              ))}
            </section>
          ))}
        </div>

        <div className="ws-sidebar__footer">
          <div className="ws-sidebar__local"><span className="ws-local-dot" /><span><strong>Processed locally</strong><small>Your files stay in this browser</small></span></div>
          <button type="button" className="ws-sidebar__shortcut" onClick={() => setPaletteOpen(true)}>
            <span>Command palette</span><kbd>⌘K</kbd>
          </button>
        </div>
      </aside>

      <div className="ws-sidebar-mobile-bar">
        <button type="button" onClick={() => setPaletteOpen(true)} aria-label="Open command palette"><SidebarIcon><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></svg></SidebarIcon><span>Search workspace</span><kbd>⌘K</kbd></button>
      </div>

      <CommandPalette tools={tools} active={active} onSelect={onSelect} open={paletteOpen} onClose={() => setPaletteOpen(false)} />
    </>
  );
}

export function EmptyState({ title, note }: { title: string; note?: string }) {
  return (
    <div className="ws-surface rounded-lg border-dashed p-10 text-center">
      <div className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-lg bg-canvas/5 text-2xl ring-1 ring-inset ring-white/10">🗂️</div>
      <h3 className="mc-display text-lg font-medium text-ink">{title}</h3>
      {note ? <p className="mx-auto mt-2 max-w-md text-sm text-[rgb(var(--ink-2))]">{note}</p> : null}
      <div className="mt-4 flex flex-wrap justify-center gap-2 text-xs text-[rgb(var(--ink-2))]">
        <span>✓ Data Validation</span>
        <span>✓ Reconciliation</span>
        <span>✓ Cleansing</span>
        <span>✓ Analysis</span>
      </div>
    </div>
  );
}
