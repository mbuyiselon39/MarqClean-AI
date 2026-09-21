import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode, KeyboardEvent as ReactKeyboardEvent } from "react";
import Logo from "./Logo";
import ThemeToggle from "./ThemeToggle";

function go(path: string) {
  const target = new URL(path, window.location.origin);
  const sameDocument = target.pathname === window.location.pathname && target.search === window.location.search;

  if (sameDocument && target.hash) {
    // Hash navigation must update the real hash so the browser performs its
    // native anchor scroll. pushState alone does not trigger hashchange or
    // scroll the target into view.
    window.location.hash = target.hash.slice(1);
    requestAnimationFrame(() => {
      document.getElementById(target.hash.slice(1))?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  } else {
    window.history.pushState({}, "", target.pathname + target.search + target.hash);
    window.dispatchEvent(new PopStateEvent("popstate"));
    if (target.hash) {
      requestAnimationFrame(() => {
        document.getElementById(target.hash.slice(1))?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } else {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }
}

function Icon({ name, size = 20 }: { name: string; size?: number }) {
  const common = { width: size, height: size, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 1.8, strokeLinecap: "round" as const, strokeLinejoin: "round" as const, "aria-hidden": true };
  const paths: Record<string, ReactNode> = {
    search: <><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></>,
    chevron: <path d="m9 18 6-6-6-6" />,
    down: <path d="m6 9 6 6 6-6" />,
    grid: <><rect x="4" y="4" width="6" height="6" rx="1" /><rect x="14" y="4" width="6" height="6" rx="1" /><rect x="4" y="14" width="6" height="6" rx="1" /><rect x="14" y="14" width="6" height="6" rx="1" /></>,
    clean: <><path d="M4 7h16M7 4v16M4 17h16" /><path d="m15 15 5 5" /></>,
    formula: <><path d="M5 5h14v14H5z" /><path d="m8 9 3 3-3 3M13 15h3" /></>,
    reconcile: <><path d="M7 7h10l-2.5-2.5M17 17H7l2.5 2.5" /><path d="M17 7a5 5 0 0 1 0 10M7 17A5 5 0 0 1 7 7" /></>,
    toolbox: <><path d="M4 8h16v12H4z" /><path d="M8 8V5h8v3M9 13h6" /></>,
    academy: <><path d="m4 7 8-4 8 4-8 4-8-4Z" /><path d="M6 9v6c3 3 9 3 12 0V9M20 8v6" /></>,
    chart: <><path d="M4 19V5M4 19h16" /><path d="m7 15 4-4 3 2 5-7" /></>,
    document: <><path d="M6 3h8l4 4v14H6z" /><path d="M14 3v5h5M9 13h6M9 17h6" /></>,
    settings: <><path d="M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8Z" /><path d="M4 12H2m20 0h-2M12 4V2m0 20v-2m5.7-14.3 1.4-1.4M4.9 19.1l1.4-1.4m11.4 0 1.4 1.4M4.9 4.9l1.4 1.4" /></>,
    arrow: <><path d="M5 12h14" /><path d="m13 6 6 6-6 6" /></>,
    close: <><path d="m6 6 12 12M18 6 6 18" /></>,
    menu: <><path d="M4 6h16M4 12h16M4 18h16" /></>,
    book: <><path d="M4 5a3 3 0 0 1 3-3h13v18H7a3 3 0 0 0-3 3V5Z" /><path d="M7 20h13" /></>,
  };
  return <svg {...common}>{paths[name] ?? paths.grid}</svg>;
}

const SHORTCUT = typeof navigator !== "undefined" && /Mac|iPhone|iPad/i.test(navigator.platform) ? "⌘ K" : "Ctrl K";

const WORKSPACES = [
  { label: "Data Cleaning", path: "/#cleaner", icon: "clean", description: "Clean, standardise and validate spreadsheet data." },
  { label: "Excel Automation", path: "/excel-automation", icon: "formula", description: "Automate workbook formatting and formulas." },
  { label: "Reconciliation Hub", path: "/reconciliation-hub", icon: "reconcile", description: "Compare records and review exceptions." },
  { label: "Data Toolbox", path: "/data-toolbox", icon: "toolbox", description: "Advanced analysis and data preparation tools." },
  { label: "Excel Academy", path: "/excel-academy", icon: "academy", description: "Learn Excel with lessons and references." },
];

const NAV = [
  { label: "Home", path: "/" },
  { label: "About", path: "/about" },
  { label: "Platform", path: "/#features" },
  { label: "Tools", path: "/#free-tools" },
  { label: "How it works", path: "/#workflow" },
  { label: "Academy", path: "/excel-academy" },
  { label: "Contact", path: "/contact" },
];

type SearchResult = { group: string; label: string; description: string; path: string; icon: string };

const TOOL_RESULTS: SearchResult[] = [
  { group: "Tools", label: "Advanced Excel Functions", description: "Functions, formulas and workbook analysis.", path: "/data-toolbox", icon: "formula" },
  { group: "Tools", label: "Data Explorer", description: "Filter and inspect structured data.", path: "/data-toolbox", icon: "chart" },
  { group: "Tools", label: "Remove Duplicates", description: "Detect and remove duplicate records.", path: "/data-toolbox", icon: "clean" },
  { group: "Tools", label: "File Comparison Centre", description: "Compare Excel, CSV, PDF and Word files.", path: "/reconciliation-hub", icon: "reconcile" },
];

function isCurrent(path: string) {
  const pathname = window.location.pathname;
  if (path.startsWith("/#")) return pathname === "/";
  return pathname === path;
}

function handleWorkspaceMenuKey(event: ReactKeyboardEvent<HTMLAnchorElement>, close: () => void) {
  const items = Array.from(document.querySelectorAll<HTMLElement>(".ws-workspace-menu [role=menuitem]"));
  const index = items.indexOf(event.currentTarget);
  if (event.key === "ArrowDown") { event.preventDefault(); items[(index + 1) % items.length]?.focus(); }
  if (event.key === "ArrowUp") { event.preventDefault(); items[(index - 1 + items.length) % items.length]?.focus(); }
  if (event.key === "Escape") { event.preventDefault(); close(); document.querySelector<HTMLElement>(".ws-nav-dropdown > button")?.focus(); }
}

export function GlobalHeader() {
  const [mobileOpen, setMobileOpen] = useState(false);
  const [workspacesOpen, setWorkspacesOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [workspaceAccordion, setWorkspaceAccordion] = useState(false);
  const workspaceTimer = useRef<number | null>(null);

  useEffect(() => {
    const shortcut = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        setSearchOpen(true);
      }
      if (!event.metaKey && !event.ctrlKey && event.key === "/" && !["INPUT","TEXTAREA","SELECT"].includes((event.target as HTMLElement)?.tagName ?? "")) {
        event.preventDefault();
        setSearchOpen(true);
      }
      if (event.key === "Escape") {
        setWorkspacesOpen(false);
        setWorkspaceAccordion(false);
        setMobileOpen(false);
      }
    };
    document.addEventListener("keydown", shortcut);
    return () => document.removeEventListener("keydown", shortcut);
  }, []);

  useEffect(() => { const openSearch = () => setSearchOpen(true); const openWorkspaces = () => setWorkspacesOpen(true); window.addEventListener("marqclean:open-command-palette", openSearch); window.addEventListener("marqclean:open-workspaces", openWorkspaces); return () => { window.removeEventListener("marqclean:open-command-palette", openSearch); window.removeEventListener("marqclean:open-workspaces", openWorkspaces); }; }, []);
  useEffect(() => () => { if (workspaceTimer.current) window.clearTimeout(workspaceTimer.current); }, []);

  const closeAll = () => { setMobileOpen(false); setWorkspacesOpen(false); setWorkspaceAccordion(false); };

  return (
    <>
      <a className="ws-skip-link" href="#main-content">Skip to content</a>
      <header className="ws-global-header" aria-label="MarqClean AI site header">
        <nav className="ws-global-header__nav" aria-label="Primary navigation">
          <a href="/" className="ws-brand" onClick={(e) => { e.preventDefault(); go("/"); closeAll(); }} aria-label="MarqClean AI home"><Logo className="shrink-0" /></a>
          <div className="ws-global-header__links">
            {NAV.slice(0, 1).map((item) => <a key={item.label} href={item.path} aria-current={isCurrent(item.path) ? "page" : undefined} onClick={(e) => { e.preventDefault(); go(item.path); }}>{item.label}</a>)}
            <div className="ws-nav-dropdown" onMouseEnter={() => { if (workspaceTimer.current) clearTimeout(workspaceTimer.current); setWorkspacesOpen(true); }} onMouseLeave={() => { workspaceTimer.current = window.setTimeout(() => setWorkspacesOpen(false), 160); }}>
              <button type="button" aria-haspopup="menu" aria-expanded={workspacesOpen} onKeyDown={(e) => { if (e.key === "ArrowDown") { e.preventDefault(); setWorkspacesOpen(true); requestAnimationFrame(() => document.querySelector<HTMLElement>(".ws-workspace-menu [role=menuitem]")?.focus()); } if (e.key === "Escape") { e.preventDefault(); setWorkspacesOpen(false); } }} onClick={() => setWorkspacesOpen((v) => !v)}>Workspaces <Icon name="down" size={16} /></button>
              {workspacesOpen ? <div className="ws-workspace-menu" role="menu" aria-label="Workspaces">
{WORKSPACES.map((workspace) => <a role="menuitem" key={workspace.label} tabIndex={0} onKeyDown={(e) => handleWorkspaceMenuKey(e, () => setWorkspacesOpen(false))} href={workspace.path} aria-current={isCurrent(workspace.path) ? "page" : undefined} onClick={(e) => { e.preventDefault(); go(workspace.path); closeAll(); }}><span className="ws-menu-icon"><Icon name={workspace.icon} size={19} /></span><span><strong>{workspace.label}</strong><small>{workspace.description}</small></span></a>)}
              </div> : null}
            </div>
            {NAV.slice(1).map((item) => <a key={item.label} href={item.path} aria-current={isCurrent(item.path) ? "page" : undefined} onClick={(e) => { e.preventDefault(); go(item.path); }}>{item.label}</a>)}
          </div>
          <div className="ws-global-header__actions">
            <button type="button" className="ws-search-trigger" onClick={() => setSearchOpen(true)} aria-label={`Open search (${SHORTCUT})`}><Icon name="search" size={18} /><span>Search</span><kbd>{SHORTCUT}</kbd></button>
            <ThemeToggle compact />
            <button type="button" className="ws-btn-primary ws-start-free" onClick={() => go("/#cleaner")}>Start free</button>
            <button type="button" className="ws-mobile-menu" aria-label={mobileOpen ? "Close navigation" : "Open navigation"} aria-expanded={mobileOpen} aria-controls="ws-mobile-panel" onClick={() => setMobileOpen((v) => !v)}><Icon name={mobileOpen ? "close" : "menu"} size={20} /></button>
          </div>
        </nav>
        {mobileOpen ? <div id="ws-mobile-panel" className="ws-mobile-panel">
          {NAV.slice(0, 1).map((item) => <a key={item.label} href={item.path} onClick={(e) => { e.preventDefault(); go(item.path); closeAll(); }}>{item.label}</a>)}
          <button type="button" className="ws-mobile-accordion" aria-expanded={workspaceAccordion} onClick={() => setWorkspaceAccordion((v) => !v)}>Workspaces <Icon name={workspaceAccordion ? "down" : "chevron"} size={16} /></button>
          {workspaceAccordion ? <div className="ws-mobile-workspaces">{WORKSPACES.map((w) => <a key={w.label} href={w.path} onClick={(e) => { e.preventDefault(); go(w.path); closeAll(); }}><Icon name={w.icon} size={18} /><span><strong>{w.label}</strong><small>{w.description}</small></span></a>)}</div> : null}
          {NAV.slice(1).map((item) => <a key={item.label} href={item.path} onClick={(e) => { e.preventDefault(); go(item.path); closeAll(); }}>{item.label}</a>)}
          <button type="button" className="ws-mobile-search" onClick={() => { setMobileOpen(false); setSearchOpen(true); }}><Icon name="search" size={18} />Search <kbd>{SHORTCUT}</kbd></button>
          <button type="button" className="ws-btn-primary" onClick={() => { closeAll(); go("/#cleaner"); }}>Start free</button>
        </div> : null}
      </header>
      <CommandPalette open={searchOpen} onClose={() => setSearchOpen(false)} />
    </>
  );
}

export type StatusTone = "idle" | "processing" | "success" | "warning" | "error";

const TONE_STYLES: Record<StatusTone, string> = {
  idle: "ws-tone-idle", processing: "ws-tone-processing", success: "ws-tone-success", warning: "ws-tone-warning", error: "ws-tone-error",
};

export function StatusBadge({ tone, label }: { tone: StatusTone; label: string }) {
  return <span className={`ws-status-badge ${TONE_STYLES[tone]}`}><span className="ws-status-dot" />{label}</span>;
}
export function FormatBadge({ label }: { label: string }) { return <span className="ws-format-badge">{label}</span>; }

type WorkspaceTool<T extends string> = { key: T; label: string; blurb: string; icon: string; category?: string };

function toolIcon(key: string, _fallback: string) {
  if (/formula|dax|function/i.test(key)) return "formula";
  if (/reconcil|comparison|match|verify/i.test(key)) return "reconcile";
  if (/bank|statement/i.test(key)) return "document";
  if (/extract|web/i.test(key)) return "document";
  if (/chart|explorer|predict|clarity|model|analysis/i.test(key)) return "chart";
  if (/academy|lesson/i.test(key)) return "academy";
  if (/setting|admin/i.test(key)) return "settings";
  return "clean";
}

function workspaceGroups(tools: Array<WorkspaceTool<string>>) {
  const keys = tools.map((t) => t.key).join(" ");
  const academy = /curriculum|functions|cheat-sheet/.test(keys);
  if (academy) return [{ name: "Learn", items: tools }];
  if (/power-query|data-model|dax-reference|timesheet|extract/.test(keys)) {
    const dataPreparation = tools.filter((t) => /functions|power-query|data-model|dax-reference|merge|dedupe|fuzzy|clarity/.test(t.key));
    const analysis = tools.filter((t) => /explorer|predict/.test(t.key));
    const extraction = tools.filter((t) => /extract|timesheet/.test(t.key));
    return [
      { name: "Data preparation", items: dataPreparation },
      { name: "Analysis", items: analysis },
      { name: "Extraction", items: extraction },
    ].filter((g) => g.items.length);
  }
  const verification = tools.filter((t) => !/compliance|reporting|administration/.test(t.key));
  const compliance = tools.filter((t) => /compliance|documents/.test(t.key));
  const reporting = tools.filter((t) => /dashboard|reporting/.test(t.key));
  const administration = tools.filter((t) => /administration|sheet-manager/.test(t.key));
  return [
    { name: "Verification", items: verification },
    { name: "Compliance", items: compliance },
    { name: "Reporting", items: reporting },
    { name: "Administration", items: administration },
  ].filter((g) => g.items.length);
}

function SidebarIcon({ name }: { name: string }) { return <span className="ws-sidebar-icon"><Icon name={name} size={20} /></span>; }

export function CommandPalette({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [query, setQuery] = useState("");
  const [cursor, setCursor] = useState(0);
  const [recent, setRecent] = useState<SearchResult[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const results = useMemo(() => {
    const all = [
      ...WORKSPACES.map((w) => ({ group: "Workspaces", label: w.label, description: w.description, path: w.path, icon: w.icon })),
      ...TOOL_RESULTS,
      { group: "Lessons", label: "Excel Academy curriculum", description: "Guided lessons and progress.", path: "/excel-academy", icon: "academy" },
      { group: "Lessons", label: "Excel cheat sheet", description: "Shortcuts, patterns and error decoder.", path: "/excel-academy", icon: "book" },
      { group: "Functions", label: "Excel function reference", description: "Search functions and examples.", path: "/excel-academy", icon: "formula" },
    ];
    const q = query.trim().toLowerCase();
    const filtered = q ? all.filter((r) => (r.label + " " + r.description + " " + r.group).toLowerCase().includes(q)) : all;
    return q ? filtered : [...recent.map((r) => ({ ...r, group: "Recent" })), ...filtered.filter((r) => !recent.some((x) => x.path === r.path && x.label === r.label))];
  }, [query]);
  useEffect(() => {
    if (!open) return;
    setQuery(""); setCursor(0);
    try { const saved = JSON.parse(window.localStorage.getItem("marqclean:recent-searches") || "[]") as SearchResult[]; setRecent(saved.slice(0, 5)); } catch { setRecent([]); }
    requestAnimationFrame(() => inputRef.current?.focus());
  }, [open]);
  useEffect(() => { if (!open) return; const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); onClose(); } if (e.key === "ArrowDown") { e.preventDefault(); setCursor((c) => Math.min(c + 1, results.length - 1)); } if (e.key === "ArrowUp") { e.preventDefault(); setCursor((c) => Math.max(c - 1, 0)); } if (e.key === "Enter" && results[cursor]) { e.preventDefault(); openResult(results[cursor]); } }; document.addEventListener("keydown", onKey); return () => document.removeEventListener("keydown", onKey); }, [open, onClose, results, cursor]);
  useEffect(() => { if (cursor >= results.length) setCursor(Math.max(results.length - 1, 0)); }, [cursor, results.length]);
  const openResult = (item: SearchResult) => {
    try {
      const next = [item, ...recent.filter((r) => r.path !== item.path || r.label !== item.label)].slice(0, 5);
      window.localStorage.setItem("marqclean:recent-searches", JSON.stringify(next));
    } catch {}
    go(item.path); onClose();
  };
  if (!open) return null;
  const grouped = results.reduce<Record<string, SearchResult[]>>((acc, item) => { (acc[item.group] ||= []).push(item); return acc; }, {});
  return <div className="ws-command-overlay" onMouseDown={onClose}><div className="ws-command" role="dialog" aria-modal="true" aria-label="Search MarqClean AI" onMouseDown={(e) => e.stopPropagation()}>
    <div className="ws-command__search"><Icon name="search" size={18} /><input ref={inputRef} value={query} onChange={(e) => { setQuery(e.target.value); setCursor(0); }} placeholder="Search workspaces, tools, lessons and functions…" aria-label="Search MarqClean AI" /><kbd>ESC</kbd></div>
    <div className="ws-command__meta"><span>{results.length} results</span><span><kbd>↑</kbd> <kbd>↓</kbd> navigate · <kbd>Enter</kbd> open</span></div>
    <div className="ws-command__results">{Object.entries(grouped).map(([group, items]) => <section key={group}><h3>{group}</h3>{items.map((item) => { const index = results.indexOf(item); return <button type="button" key={item.group + item.label} className={index === cursor ? "is-highlighted" : ""} onMouseEnter={() => setCursor(index)} onClick={() => openResult(item)}><span className="ws-command__item-icon"><Icon name={item.icon} size={19} /></span><span className="ws-command__item-copy"><strong>{item.label}</strong><small>{item.description}</small></span><Icon name="arrow" size={17} /></button>; })}</section>)}</div>
  </div></div>;
}

export function ToolLaunchpad<T extends string>({ tools, active, onSelect }: { tools: Array<WorkspaceTool<T>>; active: T; onSelect: (key: T) => void }) {
  const [collapsed, setCollapsed] = useState(() => { try { return window.localStorage.getItem("marqclean:workspace-sidebar") === "collapsed"; } catch { return false; } });
  const [drawerOpen, setDrawerOpen] = useState(false);
  useEffect(() => { try { window.localStorage.setItem("marqclean:workspace-sidebar", collapsed ? "collapsed" : "expanded"); } catch {} }, [collapsed]);
  useEffect(() => { document.body.classList.toggle("ws-drawer-open", drawerOpen); return () => document.body.classList.remove("ws-drawer-open"); }, [drawerOpen]);
  useEffect(() => {
    if (!drawerOpen) return;
    const sidebar = document.querySelector<HTMLElement>(".ws-sidebar.is-open");
    requestAnimationFrame(() => sidebar?.querySelector<HTMLElement>("button, a")?.focus());
    const trap = (event: KeyboardEvent) => {
      if (event.key === "Escape") { event.preventDefault(); setDrawerOpen(false); return; }
      if (event.key !== "Tab" || !sidebar) return;
      const focusables = Array.from(sidebar.querySelectorAll<HTMLElement>('button:not([disabled]), a[href], input:not([disabled])'));
      if (!focusables.length) return;
      const first = focusables[0], last = focusables[focusables.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", trap);
    return () => document.removeEventListener("keydown", trap);
  }, [drawerOpen]);
  const groups = workspaceGroups(tools);
  const choose = (key: T) => { onSelect(key); setDrawerOpen(false); };
  const sidebar = <aside className={`ws-sidebar ${collapsed ? "is-collapsed" : ""} ${drawerOpen ? "is-open" : ""}`} aria-label="Workspace navigation">
    <div className="ws-sidebar__scroll">
      <div className="ws-sidebar__switcher">
        <div className="ws-sidebar__workspace-mark"><Logo markOnly label="MarqClean AI" /></div>
        <div className="ws-sidebar__workspace-copy"><strong>{/functions|power-query|data-model|extract|timesheet/.test(String(active)) ? "Data Toolbox" : /dashboard|bank|mailing|comparison|cleansing|reconciliation|compliance|documents|reporting|administration|sheet-manager/.test(String(active)) ? "Reconciliation Hub" : "Workspace"}</strong><small>Current workspace</small></div>
        <button type="button" className="ws-sidebar__switcher-button" aria-label="Open workspaces" onClick={() => window.dispatchEvent(new CustomEvent("marqclean:open-workspaces"))}><Icon name="down" size={16} /></button>
      </div>
      <div className="ws-sidebar__collapse-row"><button type="button" className="ws-sidebar__collapse" aria-expanded={!collapsed} aria-label={collapsed ? "Expand sidebar" : "Collapse sidebar"} title={collapsed ? "Expand sidebar" : "Collapse sidebar"} onClick={() => setCollapsed((v) => !v)}><Icon name="chevron" size={18} /></button></div>
      {groups.map((group) => <section className="ws-sidebar__group" key={group.name}><div className="ws-sidebar__label">{group.name}</div>{group.items.map((tool) => <button type="button" key={tool.key} className={`ws-sidebar__item ${active === tool.key ? "is-active" : ""}`} onClick={() => choose(tool.key as T)} aria-current={active === tool.key ? "page" : undefined} title={collapsed ? tool.label : undefined}><SidebarIcon name={toolIcon(tool.key, tool.icon)} /><span className="ws-sidebar__item-copy"><strong>{tool.label}</strong><small>{tool.blurb}</small></span></button>)}</section>)}
    </div>
    <div className="ws-sidebar__footer"><div className="ws-sidebar__local"><span className="ws-local-dot" /><span><strong>Runs locally</strong><small>Files never leave your browser</small></span></div><a href="/contact" onClick={(e) => { e.preventDefault(); go("/contact"); setDrawerOpen(false); }}>Help & feedback</a></div>
  </aside>;
  return <><button type="button" className="ws-drawer-menu" aria-label="Open workspace menu" onClick={() => setDrawerOpen(true)}><Icon name="menu" size={18} /> <span>Menu</span></button><div className="ws-sidebar-backdrop" aria-hidden={!drawerOpen} onClick={() => setDrawerOpen(false)} />{sidebar}</>;
}

export function PageHeader({ workspace, title, description, actions, onBack }: { workspace: string; title: string; description: string; actions?: ReactNode; onBack?: () => void }) {
  const goBack = () => {
    if (onBack) return onBack();
    if (window.history.length > 1 && document.referrer.includes(window.location.host)) window.history.back();
    else go("/");
  };
  return <div className="ws-page-header"><div className="ws-page-header__back"><button type="button" onClick={goBack}><Icon name="arrow" size={17} /> Back</button><nav aria-label="Breadcrumb"><a href="/" onClick={(e) => { e.preventDefault(); go("/"); }}>Home</a><span>/</span><span>{workspace}</span><span>/</span><strong>{title}</strong></nav></div><div className="ws-page-header__row"><div><h1>{workspace}</h1><p>{description}</p></div>{actions ? <div className="ws-page-header__actions">{actions}</div> : null}</div></div>;
}

export function WorkspaceShell({ icon: _icon, title, description, engine: _engine, formats: _formats, status: _status, children }: { icon: string; title: string; description: string; engine?: string; formats?: string[]; status?: { tone: StatusTone; label: string }; children: ReactNode }) {
  return <div className="ws-root ws-scope min-h-screen"><GlobalHeader /><div className="ws-page-frame"><PageHeader workspace={title} title={title} description={description} /><div className="ws-page-body">{children}</div></div></div>;
}

export function EmptyState({ title, note }: { title: string; note?: string }) {
  return <div className="ws-surface rounded-lg border-dashed p-10 text-center"><div className="mx-auto mb-4 grid h-12 w-12 place-items-center rounded-lg border border-line bg-surface text-ink-2"><Icon name="document" size={22} /></div><h3 className="text-lg font-semibold text-ink">{title}</h3>{note ? <p className="mx-auto mt-2 max-w-md text-sm text-ink-2">{note}</p> : null}<div className="mt-4 flex flex-wrap justify-center gap-2 text-xs text-ink-2"><span>Data validation</span><span>Reconciliation</span><span>Cleansing</span><span>Analysis</span></div></div>;
}
