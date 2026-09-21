import { useEffect, useMemo, useState } from "react";
import { GlobalHeader, PageHeader } from "./WorkspaceShell";
import { LESSONS, TRACKS, byTrack, bySlug, neighbours, type Lesson } from "./lessons";
import { FUNCTIONS, FN_CATEGORIES, type FnRef } from "./functions";

type AcademyNav = "curriculum" | "functions" | "cheat-sheet";
const PROGRESS_KEY = "marqclean.academy.completedLessons";

function getCompleted(): Set<string> {
  try { const raw = window.localStorage.getItem(PROGRESS_KEY); return new Set(raw ? (JSON.parse(raw) as string[]) : []); } catch { return new Set(); }
}
function saveCompleted(set: Set<string>) {
  try { window.localStorage.setItem(PROGRESS_KEY, JSON.stringify([...set])); } catch { /* storage can be blocked */ }
}
function plural(count: number, singular: string, pluralWord = singular + "s") { return count === 1 ? `${count} ${singular}` : `${count} ${pluralWord}`; }
function icon(name: "book"|"check"|"search"|"copy"|"arrow"|"chevron") {
  const paths = {
    book: <><path d="M4 5a3 3 0 0 1 3-3h13v18H7a3 3 0 0 0-3 3V5Z"/><path d="M7 20h13"/></>,
    check: <path d="m4 12 5 5L20 6"/>,
    search: <><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/></>,
    copy: <><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"/></>,
    arrow: <><path d="M5 12h14"/><path d="m13 6 6 6-6 6"/></>,
    chevron: <path d="m9 18 6-6-6-6"/>,
  };
  return <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]}</svg>;
}

function setAcademyHash(value: string) {
  const hash = value ? `#${value}` : "";
  window.history.pushState({}, "", `/excel-academy${hash}`);
  window.dispatchEvent(new PopStateEvent("popstate"));
}

function readAcademyHash(): { nav: AcademyNav; lesson: string | null; fn: string | null } {
  const raw = window.location.hash.replace(/^#/, "");
  if (raw === "academy-functions") return { nav: "functions", lesson: null, fn: null };
  if (raw === "academy-cheat-sheet") return { nav: "cheat-sheet", lesson: null, fn: null };
  if (raw.startsWith("academy-lesson=")) return { nav: "curriculum", lesson: decodeURIComponent(raw.slice(16)), fn: null };
  if (raw.startsWith("academy-function=")) return { nav: "functions", lesson: null, fn: decodeURIComponent(raw.slice(18)) };
  return { nav: "curriculum", lesson: null, fn: null };
}

function LevelBadge({ level }: { level: Lesson["level"] }) { return <span className="ws-format-badge">{level}</span>; }

function LessonCard({ lesson, done, onOpen }: { lesson: Lesson; done: boolean; onOpen: (slug: string) => void }) {
  const state = done ? "Done" : "Not started";
  return <button type="button" className="academy-card" onClick={() => onOpen(lesson.slug)}>
    <div className="academy-card__top"><LevelBadge level={lesson.level} /><span className={done ? "academy-state is-done" : "academy-state"}>{done ? icon("check") : null}{state}</span></div>
    <h3>{lesson.title}</h3>
    <p>{lesson.summary}</p>
    <div className="academy-card__bottom"><span>{lesson.steps.length} steps</span><span>{lesson.quiz.length} question{lesson.quiz.length === 1 ? "" : "s"}</span><span>{lesson.syntax ? "Formula" : "Practical"}</span></div>
    <span className="academy-card__cta">{done ? "Review" : "Start"} {icon("arrow")}</span>
  </button>;
}

function Curriculum({ completed, onOpenLesson }: { completed: Set<string>; onOpenLesson: (slug: string) => void }) {
  const remaining = LESSONS.find((l) => !completed.has(l.slug));
  const percent = Math.round((completed.size / LESSONS.length) * 100);
  const [track, setTrack] = useState("All");
  const visibleTracks = track === "All" ? TRACKS : TRACKS.filter((t) => t.id === track);
  return <div className="academy-view">
    <section className="academy-progress-card">
      <div><span className="academy-eyebrow">Your learning progress</span><h2>{completed.size} of {LESSONS.length} lessons complete</h2><p>Your progress is saved on this device.</p></div>
      <div className="academy-progress-ring"><strong>{percent}%</strong><span>complete</span></div>
      <div className="academy-progress-bar" aria-label={`${percent}% complete`}><span style={{ width: `${percent}%` }} /></div>
      <button type="button" className="academy-reset" onClick={() => { saveCompleted(new Set()); window.location.reload(); }}>Reset progress</button>
    </section>
    {remaining ? <section className="academy-continue"><div><span className="academy-eyebrow">Continue learning</span><h2>{remaining.title}</h2><p>{remaining.summary}</p></div><button type="button" className="ws-btn-primary" onClick={() => onOpenLesson(remaining.slug)}>Continue {icon("arrow")}</button></section> : <section className="academy-continue"><div><span className="academy-eyebrow">Curriculum complete</span><h2>All {LESSONS.length} lessons are done.</h2><p>Review any lesson or use the function reference for a quick lookup.</p></div></section>}
    <div className="academy-track-filter" aria-label="Filter curriculum by track"><button type="button" className={track === "All" ? "is-active" : ""} onClick={() => setTrack("All")}>All tracks</button>{TRACKS.map((t) => <button type="button" key={t.id} className={track === t.id ? "is-active" : ""} onClick={() => setTrack(t.id)}>{t.name}</button>)}</div>
    <div className="academy-results-count">{LESSONS.filter((l) => track === "All" || l.track === track).length} lessons</div>
    {visibleTracks.map((t) => <section key={t.id} className="academy-track" id={`academy-track-${t.id}`}><div className="academy-track__heading"><div><span className="academy-eyebrow">Track</span><h2>{t.name}</h2><p>{t.blurb}</p></div><span className="ws-format-badge">{plural(byTrack(t.id).length, "lesson")}</span></div><div className="academy-card-grid">{byTrack(t.id).map((lesson) => <LessonCard key={lesson.slug} lesson={lesson} done={completed.has(lesson.slug)} onOpen={onOpenLesson} />)}</div></section>)}
  </div>;
}

function LessonView({ lesson, completed, onToggle, onOpenLesson, onBack }: { lesson: Lesson; completed: Set<string>; onToggle: (slug: string) => void; onOpenLesson: (slug: string) => void; onBack: () => void }) {
  const track = TRACKS.find((t) => t.id === lesson.track)!;
  const siblings = byTrack(lesson.track);
  const { prev, next, index } = neighbours(lesson.slug);
  const done = completed.has(lesson.slug);
  const percent = Math.round(((index + (done ? 1 : 0)) / LESSONS.length) * 100);
  return <div className="academy-lesson">
    <aside className="academy-outline">
      <div className="academy-outline__title"><span>{track.name}</span><button type="button" onClick={onBack}>All lessons</button></div>
      <div className="academy-outline__progress"><span style={{ width: `${Math.min(100, Math.round((siblings.filter((s) => completed.has(s.slug)).length / siblings.length) * 100))}%` }} /></div>
      <ol>{siblings.map((item) => <li key={item.slug}><button type="button" className={item.slug === lesson.slug ? "is-active" : ""} onClick={() => onOpenLesson(item.slug)}>{completed.has(item.slug) ? icon("check") : <span className="academy-outline__number">{siblings.indexOf(item)+1}</span>}<span>{item.title.split(" - ")[0]}</span></button></li>)}</ol>
    </aside>
    <article className="academy-lesson__content">
      <nav className="academy-breadcrumb" aria-label="Breadcrumb"><button type="button" onClick={onBack}>Academy</button><span>/</span><span>{track.name}</span><span>/</span><strong>{lesson.title}</strong></nav>
      <div className="academy-lesson__heading"><div><LevelBadge level={lesson.level}/><h1>{lesson.title}</h1><p>{lesson.summary}</p></div><div className="academy-lesson__progress"><span>Lesson {index + 1} of {LESSONS.length}</span><div><span style={{width:`${percent}%`}} /></div></div></div>
      <section className="academy-section"><span className="academy-eyebrow">Objectives</span><ul className="academy-objectives">{lesson.objectives.map((o) => <li key={o}>{icon("check")}<span>{o}</span></li>)}</ul></section>
      {lesson.syntax ? <section className="academy-section"><h2>Syntax</h2><pre>{lesson.syntax}</pre></section> : null}
      <section className="academy-section"><h2>Steps</h2><ol className="academy-steps">{lesson.steps.map((step, i) => <li key={step.title}><span>{i+1}</span><div><h3>{step.title}</h3>{step.detail ? <p>{step.detail}</p> : null}{step.formula ? <pre>{step.formula}</pre> : null}</div></li>)}</ol></section>
      {lesson.example ? <section className="academy-section"><h2>Worked example</h2><div className="academy-table-wrap"><table><thead><tr>{lesson.example.headers.map((h) => <th key={h}>{h}</th>)}</tr></thead><tbody>{lesson.example.rows.map((row, i) => <tr key={i}>{row.map((cell, j) => <td key={j}>{cell}</td>)}</tr>)}</tbody></table></div></section> : null}
      {lesson.formulas?.length ? <section className="academy-section"><h2>Formula patterns</h2><div className="academy-snippet-grid">{lesson.formulas.map((f) => <div key={f.label}><strong>{f.label}</strong><pre>{f.code}</pre></div>)}</div></section> : null}
      {lesson.shortcuts?.length ? <section className="academy-section"><h2>Shortcuts</h2><div className="academy-shortcuts">{lesson.shortcuts.map(([key, value]) => <div key={key}><kbd>{key}</kbd><span>{value}</span></div>)}</div></section> : null}
      {lesson.practice ? <PracticeTask task={lesson.practice} /> : null}
      <Quiz questions={lesson.quiz} />
      <section className="academy-complete"><div><span className="academy-eyebrow">Lesson status</span><h2>{done ? "Lesson complete" : "Ready to mark this lesson complete?"}</h2><p>{done ? "You can review it at any time." : "Marking a lesson complete updates your progress on this device."}</p></div><button type="button" className={done ? "ws-btn-secondary" : "ws-btn-primary"} onClick={() => onToggle(lesson.slug)}>{done ? "Mark as not complete" : "Mark complete"} {icon("check")}</button></section>
      <nav className="academy-next-prev"><button type="button" disabled={!prev} onClick={() => prev && onOpenLesson(prev.slug)}>← Previous</button><button type="button" disabled={!next} onClick={() => next && onOpenLesson(next.slug)}>Next lesson →</button></nav>
    </article>
  </div>;
}

function PracticeTask({ task }: { task: NonNullable<Lesson["practice"]> }) {
  const [show, setShow] = useState(false);
  return <section className="academy-section academy-practice"><span className="academy-eyebrow">Worked practice</span><h2>Try it yourself</h2><p>{task.task}</p>{show ? <pre>{task.answer}</pre> : <button type="button" className="ws-btn-secondary" onClick={() => setShow(true)}>Show model answer</button>}</section>;
}
function Quiz({ questions }: { questions: Lesson["quiz"] }) {
  const [picked, setPicked] = useState<Record<number, number>>({});
  const [checked, setChecked] = useState(false);
  const score = questions.reduce((n, q, i) => n + (picked[i] === q.answer ? 1 : 0), 0);
  return <section className="academy-section academy-quiz"><span className="academy-eyebrow">Knowledge check</span><h2>Quick quiz</h2>{questions.map((q, i) => <div key={q.q} className="academy-question"><p>{i+1}. {q.q}</p><div>{q.options.map((option, oi) => <button type="button" key={option} disabled={checked} className={picked[i] === oi ? "is-picked" : ""} onClick={() => setPicked((v) => ({...v,[i]:oi}))}>{option}</button>)}</div>{checked ? <small>{q.explain}</small> : null}</div>)}<div className="academy-quiz__actions">{checked ? <><strong>Score {score} / {questions.length}</strong><button type="button" className="ws-btn-secondary" onClick={() => {setChecked(false);setPicked({});}}>Try again</button></> : <button type="button" className="ws-btn-primary" onClick={() => setChecked(true)}>Check answers</button>}</div></section>;
}

function FunctionsView({ selected, onSelect }: { selected: FnRef | null; onSelect: (fn: FnRef | null) => void }) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("All");
  const rows = useMemo(() => { const q = query.trim().toLowerCase(); return FUNCTIONS.filter((f) => (category === "All" || f.category === category) && (!q || `${f.name} ${f.desc} ${f.syntax}`.toLowerCase().includes(q))); }, [query, category]);
  const selectedName = selected?.name ?? "";
  if (selected) return <section className="academy-function-detail"><button type="button" className="academy-text-button" onClick={() => onSelect(null)}>← All functions</button><div className="academy-detail-heading"><span className="ws-format-badge">{selected.category}</span><h1>{selected.name}</h1><p>{selected.desc}</p></div><div className="academy-function-block"><span className="academy-eyebrow">Syntax</span><pre>{selected.syntax}</pre></div><div className="academy-function-block"><span className="academy-eyebrow">Example</span><pre>{selected.example}</pre><p>Result: {selected.result}</p><button type="button" className="academy-copy" onClick={() => navigator.clipboard?.writeText(selected.example)}>Copy example {icon("copy")}</button></div>{selected.note ? <p className="academy-note">{selected.note}</p> : null}{selected.lesson ? <button type="button" className="ws-btn-secondary" onClick={() => { onSelect(null); setAcademyHash(`academy-lesson=${encodeURIComponent(selected.lesson!)}`); }}>Open related lesson {icon("arrow")}</button> : null}</section>;
  return <section className="academy-functions"><div className="academy-search-row"><label><span>Filter functions</span><div><span className="academy-search-icon">{icon("search")}</span><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search by function, syntax or use case" /><button type="button" aria-label="Clear function search" onClick={() => setQuery("")}>×</button></div></label><label><span>Category</span><select value={category} onChange={(e) => setCategory(e.target.value)}><option>All</option>{FN_CATEGORIES.map((c) => <option key={c}>{c}</option>)}</select></label></div><div className="academy-results-count">{rows.length} functions</div><div className="academy-function-grid">{rows.map((f) => <button type="button" key={f.name} className="academy-function-card" onClick={() => { onSelect(f); setAcademyHash(`academy-function=${encodeURIComponent(f.name)}`); }}><div><strong>{f.name}</strong><span>{f.category}</span></div><p>{f.desc}</p><code>{f.example}</code></button>)}</div></section>;
}

function CheatSheetView() {
  const lesson = bySlug("cheat-sheet");
  const errors: Array<[string,string,string]> = [["#NAME?","Misspelled function or unsupported name","Check spelling and Excel version"],["#VALUE!","Text where a number is expected","Clean the input with VALUE or TRIM"],["#REF!","A referenced cell was deleted","Restore or rewrite the reference"],["#DIV/0!","Dividing by zero or a blank",'Use IFERROR(B2/C2,"-")'],["#N/A","A lookup found nothing","Clean the lookup value and use IFNA"],["#SPILL!","Something blocks a dynamic array","Clear the output range"],["#NUM!","Invalid numeric calculation","Check the inputs"],["#CALC!","An array calculation returned no result","Check FILTER and its if_empty argument"]];
  return <section className="academy-cheat"><div className="academy-cheat-grid"><div><span className="academy-eyebrow">Shortcuts</span><h2>Work faster</h2><div className="academy-shortcuts">{lesson?.shortcuts?.map(([k,v]) => <div key={k}><kbd>{k}</kbd><span>{v}</span></div>)}</div></div><div><span className="academy-eyebrow">Patterns</span><h2>Reusable formulas</h2><div className="academy-snippet-grid">{lesson?.formulas?.map((f) => <div key={f.label}><strong>{f.label}</strong><pre>{f.code}</pre></div>)}</div></div></div><div className="academy-section"><span className="academy-eyebrow">Error decoder</span><h2>Understand common errors</h2><div className="academy-table-wrap"><table><thead><tr><th>Error</th><th>Usually means</th><th>Next step</th></tr></thead><tbody>{errors.map((r) => <tr key={r[0]}>{r.map((c,i)=><td key={i}>{c}</td>)}</tr>)}</tbody></table></div></div></section>;
}

export default function AcademyHub({ onExit }: { onExit: () => void }) {
  const initial = readAcademyHash();
  const [nav, setNav] = useState<AcademyNav>(initial.nav);
  const [activeLessonSlug, setActiveLessonSlug] = useState<string | null>(initial.lesson);
  const [selectedFn, setSelectedFn] = useState<FnRef | null>(() => initial.fn ? FUNCTIONS.find((f) => f.name === initial.fn) ?? null : null);
  const [completed, setCompleted] = useState<Set<string>>(() => getCompleted());

  useEffect(() => {
    const sync = () => { const next = readAcademyHash(); setNav(next.nav); setActiveLessonSlug(next.lesson); setSelectedFn(next.fn ? FUNCTIONS.find((f) => f.name === next.fn) ?? null : null); };
    window.addEventListener("popstate", sync);
    return () => window.removeEventListener("popstate", sync);
  }, []);

  function openLesson(slug: string) { setNav("curriculum"); setAcademyHash(`academy-lesson=${encodeURIComponent(slug)}`); setActiveLessonSlug(slug); }
  function openTab(next: AcademyNav) { setNav(next); setActiveLessonSlug(null); setSelectedFn(null); setAcademyHash(next === "functions" ? "academy-functions" : next === "cheat-sheet" ? "academy-cheat-sheet" : "academy-curriculum"); }
  function toggleComplete(slug: string) { setCompleted((prev) => { const next = new Set(prev); if (next.has(slug)) next.delete(slug); else next.add(slug); saveCompleted(next); return next; }); }

  const activeLesson = activeLessonSlug ? bySlug(activeLessonSlug) : null;
  return <div className="ws-root ws-scope min-h-screen"><GlobalHeader /><main id="main-content" className="academy-page">
    <div className="ws-page-frame">
      <PageHeader workspace="Excel Academy" title={activeLesson ? activeLesson.title : "Excel Academy"} description={activeLesson ? "Learn step by step with examples, practice and a knowledge check." : "A structured Excel learning space with lessons, a function reference and a practical cheat sheet."} onBack={activeLesson ? () => { setAcademyHash(""); setActiveLessonSlug(null); } : undefined} />
      {!activeLesson ? <nav className="academy-tabs" aria-label="Academy sections"><button type="button" className={nav === "curriculum" ? "is-active" : ""} aria-current={nav === "curriculum" ? "page" : undefined} onClick={() => openTab("curriculum")}>Curriculum</button><button type="button" className={nav === "functions" ? "is-active" : ""} aria-current={nav === "functions" ? "page" : undefined} onClick={() => openTab("functions")}>Function reference</button><button type="button" className={nav === "cheat-sheet" ? "is-active" : ""} aria-current={nav === "cheat-sheet" ? "page" : undefined} onClick={() => openTab("cheat-sheet")}>Cheat sheet</button></nav> : null}
      {activeLesson ? <LessonView lesson={activeLesson} completed={completed} onToggle={toggleComplete} onOpenLesson={openLesson} onBack={() => { setAcademyHash(""); setActiveLessonSlug(null); }} /> : nav === "curriculum" ? <Curriculum completed={completed} onOpenLesson={openLesson} /> : nav === "functions" ? <FunctionsView selected={selectedFn} onSelect={setSelectedFn} /> : <CheatSheetView />}
    </div>
  </main></div>;
}
