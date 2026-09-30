import { useEffect, useState } from "react";

export type ThemePreference = "light" | "dark" | "system";

const STORAGE_KEY = "marqclean:theme";

function readPreference(): ThemePreference {
  if (typeof window === "undefined") return "system";
  let value: string | null = null;
  try {
    value = window.localStorage.getItem(STORAGE_KEY);
  } catch {
    return "system";
  }
  return value === "light" || value === "dark" || value === "system" ? value : "system";
}

function applyTheme(preference: ThemePreference) {
  const resolved = preference === "system"
    ? (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light")
    : preference;
  document.documentElement.setAttribute("data-theme", resolved);
}

export function initialiseTheme() {
  const preference = readPreference();
  applyTheme(preference);
}

export default function ThemeToggle({ compact = false }: { compact?: boolean }) {
  const [preference, setPreference] = useState<ThemePreference>(() => readPreference());

  useEffect(() => {
    applyTheme(preference);
    try {
      window.localStorage.setItem(STORAGE_KEY, preference);
    } catch {
      // Storage may be blocked; keep the current preference in memory.
    }
  }, [preference]);

  useEffect(() => {
    if (preference !== "system") return;
    const query = window.matchMedia("(prefers-color-scheme: dark)");
    const sync = () => applyTheme("system");
    query.addEventListener("change", sync);
    return () => query.removeEventListener("change", sync);
  }, [preference]);

  const next: Record<ThemePreference, ThemePreference> = {
    system: "light",
    light: "dark",
    dark: "system",
  };

  const label =
    preference === "system"
      ? "Theme: system"
      : preference === "light"
        ? "Theme: light"
        : "Theme: dark";

  return (
    <button
      type="button"
      className={`mc-theme-toggle ${compact ? "is-compact" : ""}`}
      onClick={() => setPreference(next[preference])}
      aria-label={`${label}. Click to switch.`}
      title={label}
    >
      {preference === "dark" ? (
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20.8 15.6A8.5 8.5 0 0 1 8.4 3.2 8.5 8.5 0 1 0 20.8 15.6Z" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg>
      ) : preference === "light" ? (
        <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4" fill="none" stroke="currentColor" strokeWidth="1.7" /><path d="M12 2v2M12 20v2M4.93 4.93l1.42 1.42M17.65 17.65l1.42 1.42M2 12h2M20 12h2M4.93 19.07l1.42-1.42M17.65 6.35l1.42-1.42" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" /></svg>
      ) : (
        <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2" fill="none" stroke="currentColor" strokeWidth="1.7" /><path d="M12 4v16" stroke="currentColor" strokeWidth="1.7" /></svg>
      )}
      {!compact ? <span>{preference}</span> : null}
    </button>
  );
}
