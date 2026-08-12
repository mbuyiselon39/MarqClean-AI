import { useMemo, useState } from "react";
import { GlobalHeader, ToolLaunchpad } from "../reconciliation/WorkspaceShell";
import { LESSONS, TRACKS, byTrack, bySlug, neighbours, type Lesson } from "./lessons";
import { FUNCTIONS, FN_CATEGORIES, type FnRef } from "./functions";

type AcademyNav = "curriculum" | "functions" | "cheat-sheet";

const NAV_ITEMS: Array<{ key: AcademyNav; label: string; blurb: string; icon: string }> = [
  { key: "curriculum", label: "Curriculum", blurb: `${LESSONS.length} guided lessons across ${TRACKS.length} tracks.`, icon: "🎓" },
  { key: "functions", label: "Function Reference", blurb: `Searchable A-Z reference of ${FUNCTIONS.length} Excel functions.`, icon: "🔎" },
  { key: "cheat-sheet", label: "Cheat Sheet", blurb: "Shortcuts, formula patterns, and an error-message decoder.", icon: "📋" },
];

const LEVEL_STYLE: Record<Lesson["level"], string> = {
  Beginner: "text-emerald-300 bg-emerald-400/10 border-emerald-400/30",
  Intermediate: "text-amber-300 bg-amber-400/10 border-amber-400/30",
  Advanced: "text-fuchsia-300 bg-fuchsia-400/10 border-fuchsia-400/30",
};

const ERROR_DECODER: Array<[string, string, string]> = [
  ["#NAME?", "Misspelled function, missing quotes, or a function your Excel version does not have", "Check the spelling and your Excel version"],
  ["#VALUE!", "Text where a number is expected", "Clean the data with VALUE or TRIM"],
  ["#REF!", "A referenced cell was deleted", "Undo, or rewrite the reference"],
  ["#DIV/0!", "Dividing by zero or a blank", '=IFERROR(B2/C2,"-")'],
  ["#N/A", "A lookup found nothing", '=IFNA(...,"Not found") and TRIM the lookup value'],
  ["#SPILL!", "Something blocks a dynamic array result", "Clear the cells below/right"],
  ["#NUM!", "Impossible maths (e.g. SQRT of a negative)", "Check the inputs"],
  ["#CALC!", "An array formula returned nothing", "Add the if_empty argument to FILTER"],
];

const PROGRESS_KEY = "marqclean.academy.completedLessons";

function getCompleted(): Set<string> {
  try {
    const raw = window.localStorage.getItem(PROGRESS_KEY);
    return new Set(raw ? (JSON.parse(raw) as string[]) : []);
  } catch {
    return new Set();
  }
}

function saveCompleted(set: Set<string>) {
  try {
    window.localStorage.setItem(PROGRESS_KEY, JSON.stringify(Array.from(set)));
  } catch {
    // localStorage unavailable (e.g. private browsing) — progress simply won't persist
  }
}

function Quiz({ questions }: { questions: Lesson["quiz"] }) {
  const [picked, setPicked] = useState<Record<number, number>>({});
  const [checked, setChecked] = useState(false);
  const score = questions.reduce((s, q, i) => s + (picked[i] === q.answer ? 1 : 0), 0);

  return (
    <div className="mc-glass rounded-2xl p-5">
      <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-[#5EEAD4]">Check your understanding</p>
      <div className="mt-4 space-y-5">
        {questions.map((q, i) => (
          <div key={q.q}>
            <p className="text-[15px] font-medium text-white">{i + 1}. {q.q}</p>
            <div className="mt-2 grid gap-2 sm:grid-cols-2">
              {q.options.map((option, oi) => {
                const isPicked = picked[i] === oi;
                const correct = checked && oi === q.answer;
                const wrong = checked && isPicked && oi !== q.answer;
                return (
                  <button
                    key={option}
                    onClick={() => !checked && setPicked({ ...picked, [i]: oi })}
                    className={`rounded-lg border px-3 py-2 text-left text-sm transition ${
                      correct
                        ? "border-emerald-400/70 bg-emerald-400/10 text-emerald-200"
                        : wrong
                          ? "border-red-400/70 bg-red-400/10 text-red-200"
                          : isPicked
                            ? "border-[#14B8A6] bg-[#0D9488]/15 text-white"
                            : "border-white/12 text-slate-300 hover:border-white/30"
                    }`}
                  >
                    {option}
                  </button>
                );
              })}
            </div>
            {checked ? <p className="mt-2 text-xs text-slate-400">💡 {q.explain}</p> : null}
          </div>
        ))}
      </div>
      <div className="mt-5 flex flex-wrap items-center gap-3">
        {!checked ? (
          <button onClick={() => setChecked(true)} className="mc-cta-cyan rounded-lg px-5 py-2 text-sm font-semibold">
            Check answers
          </button>
        ) : (
          <>
            <span className="rounded-lg bg-white/10 px-4 py-2 text-sm font-semibold text-white">Score {score} / {questions.length}</span>
            <button onClick={() => { setChecked(false); setPicked({}); }} className="rounded-lg border border-white/15 px-4 py-2 text-sm text-slate-300 hover:text-white">
              Try again
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function LessonDetail({ lesson, onOpenLesson, onBack, completed, onToggleComplete }: {
  lesson: Lesson;
  onOpenLesson: (slug: string) => void;
  onBack: () => void;
  completed: Set<string>;
  onToggleComplete: (slug: string) => void;
}) {
  const track = TRACKS.find((t) => t.id === lesson.track)!;
  const { prev, next, index } = neighbours(lesson.slug);
  const siblings = byTrack(lesson.track);
  const isDone = completed.has(lesson.slug);

  return (
    <div className="grid gap-10 lg:grid-cols-[220px_1fr]">
      <aside className="hidden lg:block">
        <div className="sticky top-24">
          <p className="text-xs font-semibold uppercase tracking-widest text-slate-500">{track.icon} {track.name}</p>
          <ul className="mt-3 space-y-1 border-l border-white/10 text-sm">
            {siblings.map((s) => (
              <li key={s.slug}>
                <button
                  onClick={() => onOpenLesson(s.slug)}
                  className={`-ml-px block w-full border-l-2 py-1.5 pl-3 text-left ${s.slug === lesson.slug ? "border-[#0D9488] font-semibold text-white" : "border-transparent text-slate-400 hover:border-white/25 hover:text-slate-200"}`}
                >
                  {s.title.split(" - ")[0]}
                </button>
              </li>
            ))}
          </ul>
          <button onClick={onBack} className="mt-4 inline-block text-xs text-[#5EEAD4] hover:text-white">← All tracks</button>
        </div>
      </aside>

      <article className="min-w-0">
        <nav className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
          <button onClick={onBack} className="hover:text-white">Curriculum</button>
          <span>/</span>
          <span className="text-slate-400">{track.name} · Lesson {index + 1} of {LESSONS.length}</span>
        </nav>

        <h1 className="mt-3 text-3xl font-black leading-tight text-white md:text-4xl">{lesson.title}</h1>
        <div className="mt-3 flex flex-wrap items-center gap-2 text-xs">
          <span className={`rounded-full border px-2.5 py-1 ${LEVEL_STYLE[lesson.level]}`}>{lesson.level}</span>
          <span className="rounded-full border border-white/15 bg-white/5 px-2.5 py-1 text-slate-300">{track.icon} {track.name}</span>
          {isDone ? <span className="rounded-full border border-emerald-400/40 bg-emerald-400/10 px-2.5 py-1 text-emerald-300">✓ Completed</span> : null}
        </div>
        <p className="mt-4 rounded-xl border-l-2 border-[#0D9488] bg-white/5 px-4 py-3 text-[15px] leading-7 text-slate-200">{lesson.summary}</p>

        <section className="mt-8 mc-glass rounded-2xl p-5">
          <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-[#5EEAD4]">What you will be able to do</p>
          <ul className="mt-3 space-y-1.5 text-[15px] text-slate-300">
            {lesson.objectives.map((o) => <li key={o} className="flex gap-2"><span className="text-emerald-400">✓</span>{o}</li>)}
          </ul>
        </section>

        {lesson.syntax ? (
          <section className="mt-6">
            <h2 className="text-xl font-bold text-white">Syntax</h2>
            <pre className="mt-2 overflow-x-auto rounded-xl border border-white/10 bg-black/45 p-4 font-mono text-sm text-emerald-300">{lesson.syntax}</pre>
          </section>
        ) : null}

        <section className="mt-8">
          <h2 className="text-xl font-bold text-white">Step by step</h2>
          <ol className="mt-4 space-y-4">
            {lesson.steps.map((s, i) => (
              <li key={s.title} className="mc-glass rounded-2xl p-5">
                <div className="flex gap-3">
                  <span className="grid h-7 w-7 shrink-0 place-items-center rounded-lg bg-[#0D9488]/25 text-sm font-bold text-[#5EEAD4]">{i + 1}</span>
                  <div className="min-w-0 flex-1">
                    <p className="font-semibold text-white">{s.title}</p>
                    {s.detail ? <p className="mt-1.5 text-[15px] leading-7 text-slate-300">{s.detail}</p> : null}
                    {s.formula ? <pre className="mt-3 overflow-x-auto whitespace-pre-wrap rounded-lg border border-white/10 bg-black/45 p-3 font-mono text-xs text-emerald-300">{s.formula}</pre> : null}
                  </div>
                </div>
              </li>
            ))}
          </ol>
        </section>

        {lesson.example ? (
          <section className="mt-8">
            <h2 className="text-xl font-bold text-white">Worked example</h2>
            {lesson.example.caption ? <p className="mt-1 text-sm text-slate-400">{lesson.example.caption}</p> : null}
            <div className="mt-3 overflow-x-auto rounded-xl border border-white/10">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr>{lesson.example.headers.map((h) => <th key={h} className="border-b border-white/10 bg-white/8 px-3 py-2 text-left font-semibold text-slate-200">{h}</th>)}</tr>
                </thead>
                <tbody>
                  {lesson.example.rows.map((r, i) => (
                    <tr key={i} className="odd:bg-white/[0.02]">
                      {r.map((c, j) => <td key={j} className="border-b border-white/5 px-3 py-2 align-top text-slate-300">{c}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        {lesson.formulas && lesson.formulas.length > 0 ? (
          <section className="mt-8">
            <h2 className="text-xl font-bold text-white">Ready-made snippets</h2>
            <div className="mt-3 space-y-3">
              {lesson.formulas.map((f) => (
                <div key={f.label}>
                  <p className="text-sm font-medium text-slate-300">{f.label}</p>
                  <pre className="mt-1 overflow-x-auto whitespace-pre-wrap rounded-lg border border-white/10 bg-black/45 p-3 font-mono text-xs text-emerald-300">{f.code}</pre>
                </div>
              ))}
            </div>
          </section>
        ) : null}

        {lesson.shortcuts && lesson.shortcuts.length > 0 ? (
          <section className="mt-8">
            <h2 className="text-xl font-bold text-white">Shortcuts</h2>
            <div className="mt-3 grid gap-2 sm:grid-cols-2">
              {lesson.shortcuts.map(([k, v]) => (
                <div key={k} className="flex items-center gap-3 rounded-lg border border-white/10 bg-white/[0.03] px-3 py-2">
                  <kbd className="rounded border border-white/15 bg-black/40 px-2 py-1 font-mono text-xs text-[#5EEAD4]">{k}</kbd>
                  <span className="text-sm text-slate-300">{v}</span>
                </div>
              ))}
            </div>
          </section>
        ) : null}

        {lesson.tips && lesson.tips.length > 0 ? (
          <section className="mt-8 mc-glass rounded-2xl p-5">
            <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-emerald-300">Trainer tips</p>
            <ul className="mt-3 space-y-2 text-[15px] text-slate-300">
              {lesson.tips.map((t) => <li key={t} className="flex gap-2"><span>💡</span>{t}</li>)}
            </ul>
          </section>
        ) : null}

        {lesson.mistakes && lesson.mistakes.length > 0 ? (
          <section className="mt-6 rounded-2xl border border-red-400/25 bg-red-500/[0.06] p-5">
            <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-red-300">Common mistakes</p>
            <ul className="mt-3 space-y-2 text-[15px] text-slate-300">
              {lesson.mistakes.map((m) => <li key={m} className="flex gap-2"><span className="text-red-400">✕</span>{m}</li>)}
            </ul>
          </section>
        ) : null}

        <div className="mt-10 space-y-6">
          {lesson.practice ? <PracticeTask task={lesson.practice} /> : null}
          <Quiz questions={lesson.quiz} />
          <button
            onClick={() => onToggleComplete(lesson.slug)}
            className={`rounded-lg px-5 py-2.5 text-sm font-semibold ${isDone ? "border border-emerald-400/60 bg-emerald-400/10 text-emerald-300" : "mc-cta-cyan"}`}
          >
            {isDone ? "✓ Lesson completed" : "Mark lesson as complete"}
          </button>
        </div>

        <nav className="mt-10 flex flex-wrap justify-between gap-3 border-t border-white/10 pt-6">
          {prev ? (
            <button onClick={() => onOpenLesson(prev.slug)} className="mc-glass max-w-[48%] rounded-xl p-4 text-left text-sm">
              <span className="text-xs text-slate-500">← Previous</span>
              <p className="mt-1 font-semibold text-white">{prev.title}</p>
            </button>
          ) : <span />}
          {next ? (
            <button onClick={() => onOpenLesson(next.slug)} className="mc-glass ml-auto max-w-[48%] rounded-xl p-4 text-right text-sm">
              <span className="text-xs text-slate-500">Next →</span>
              <p className="mt-1 font-semibold text-white">{next.title}</p>
            </button>
          ) : null}
        </nav>
      </article>
    </div>
  );
}

function PracticeTask({ task }: { task: NonNullable<Lesson["practice"]> }) {
  const [showAnswer, setShowAnswer] = useState(false);
  return (
    <div className="mc-glass rounded-2xl p-5">
      <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-[#5EEAD4]">Practice task</p>
      <p className="mt-2 text-[15px] text-slate-200">{task.task}</p>
      {showAnswer ? (
        <pre className="mt-3 overflow-x-auto rounded-lg border border-white/10 bg-black/40 p-3 font-mono text-xs text-emerald-300">{task.answer}</pre>
      ) : (
        <button onClick={() => setShowAnswer(true)} className="mt-3 rounded-lg border border-white/15 px-4 py-2 text-sm text-slate-300 hover:border-[#14B8A6] hover:text-white">
          Show model answer
        </button>
      )}
    </div>
  );
}

function CurriculumView({ onOpenLesson, completed }: { onOpenLesson: (slug: string) => void; completed: Set<string> }) {
  return (
    <div>
      <h1 className="text-3xl font-black text-white md:text-4xl">The full curriculum</h1>
      <p className="mt-3 max-w-3xl text-slate-300">
        {LESSONS.length} lessons across {TRACKS.length} tracks. Each one is a complete training module — objectives,
        numbered steps, syntax, a worked example, common mistakes, a practice task, and a quiz.
      </p>
      <p className="mt-2 text-sm text-[#5EEAD4]">{completed.size} of {LESSONS.length} lessons completed{completed.size > 0 ? ` (${Math.round((completed.size / LESSONS.length) * 100)}%)` : ""}.</p>

      <div className="mt-6 flex flex-wrap gap-2">
        {TRACKS.map((t) => (
          <a key={t.id} href={`#academy-track-${t.id}`} className="rounded-full border border-white/12 bg-white/5 px-3 py-1.5 text-sm text-slate-300 hover:border-[#14B8A6] hover:text-white">
            {t.icon} {t.name}
          </a>
        ))}
      </div>

      <div className="mt-12 space-y-14">
        {TRACKS.map((t) => (
          <section key={t.id} id={`academy-track-${t.id}`} className="scroll-mt-24">
            <div className="flex flex-wrap items-baseline gap-3">
              <h2 className="text-2xl font-bold text-white">{t.icon} {t.name}</h2>
              <span className="text-sm text-slate-500">{byTrack(t.id).length} lessons</span>
            </div>
            <p className="mt-1.5 max-w-3xl text-sm text-slate-400">{t.blurb}</p>
            <div className="mt-5 grid gap-4 md:grid-cols-2 lg:grid-cols-3">
              {byTrack(t.id).map((l, i) => (
                <button key={l.slug} onClick={() => onOpenLesson(l.slug)} className="mc-glass group flex h-full flex-col rounded-2xl p-5 text-left transition">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-mono text-xs text-slate-500">{String(i + 1).padStart(2, "0")}</span>
                    <span className="flex items-center gap-1.5">
                      {completed.has(l.slug) ? <span className="text-emerald-400">✓</span> : null}
                      <span className={`rounded-full border px-2 py-0.5 text-[11px] ${LEVEL_STYLE[l.level]}`}>{l.level}</span>
                    </span>
                  </div>
                  <h3 className="mt-2.5 font-bold leading-snug text-white group-hover:text-[#5EEAD4]">{l.title}</h3>
                  <p className="mt-2 flex-1 text-sm leading-6 text-slate-400">{l.summary}</p>
                  {l.syntax ? <p className="mt-3 truncate font-mono text-[11px] text-[#14B8A6]">{l.syntax}</p> : null}
                </button>
              ))}
            </div>
          </section>
        ))}
      </div>
    </div>
  );
}

function FunctionsView({ selected, onSelect }: { selected: FnRef | null; onSelect: (fn: FnRef | null) => void }) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("All");

  const rows = useMemo(() => {
    const q = query.trim().toLowerCase();
    return FUNCTIONS.filter((f) => {
      if (category !== "All" && f.category !== category) return false;
      if (!q) return true;
      return f.name.toLowerCase().includes(q) || f.desc.toLowerCase().includes(q) || f.syntax.toLowerCase().includes(q);
    });
  }, [query, category]);

  if (selected) {
    const related = selected.lesson ? bySlug(selected.lesson) : undefined;
    return (
      <div>
        <button onClick={() => onSelect(null)} className="text-sm text-[#5EEAD4] hover:text-white">← All functions</button>
        <h1 className="mt-3 font-mono text-3xl font-black text-white">{selected.name}</h1>
        <span className="mt-2 inline-block rounded-full border border-white/15 bg-white/5 px-2.5 py-1 text-xs text-slate-300">{selected.category}</span>
        <p className="mt-4 max-w-2xl text-[15px] leading-7 text-slate-300">{selected.desc}</p>
        <pre className="mt-4 overflow-x-auto rounded-xl border border-white/10 bg-black/45 p-4 font-mono text-sm text-emerald-300">{selected.syntax}</pre>
        <div className="mt-4 mc-glass rounded-2xl p-5">
          <p className="mc-mono text-xs font-semibold uppercase tracking-widest text-[#5EEAD4]">Example</p>
          <pre className="mt-2 overflow-x-auto whitespace-pre-wrap font-mono text-sm text-emerald-300">{selected.example}</pre>
          <p className="mt-2 text-sm text-slate-400">Result: {selected.result}</p>
        </div>
        {selected.note ? <p className="mt-4 text-sm text-amber-300">💡 {selected.note}</p> : null}
        {related ? (
          <button onClick={() => onSelect(null)} className="mt-6 inline-block rounded-lg border border-white/15 px-4 py-2 text-sm text-slate-300 hover:border-[#14B8A6] hover:text-white">
            View the full lesson: {related.title} →
          </button>
        ) : null}
      </div>
    );
  }

  return (
    <div>
      <h1 className="text-3xl font-black text-white md:text-4xl">Function Reference</h1>
      <p className="mt-3 max-w-3xl text-slate-300">A complete A-Z Excel function reference — {FUNCTIONS.length} functions, grouped the way Excel groups them.</p>
      <div className="mt-6 mc-glass flex flex-wrap gap-3 rounded-2xl p-4">
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search: xlookup, round, interest, count..."
          className="min-w-[240px] flex-1 rounded-lg border border-white/12 bg-black/30 px-3 py-2 text-sm text-white outline-none placeholder:text-slate-500 focus:border-[#14B8A6]"
        />
        <select value={category} onChange={(e) => setCategory(e.target.value)} className="rounded-lg border border-white/12 bg-black/30 px-3 py-2 text-sm text-white outline-none focus:border-[#14B8A6]">
          {["All", ...FN_CATEGORIES].map((c) => <option key={c} className="bg-[#0d0f18]">{c}</option>)}
        </select>
      </div>
      <p className="mt-3 text-sm text-slate-500">{rows.length} functions</p>
      <div className="mt-3 grid gap-3 md:grid-cols-2 lg:grid-cols-3">
        {rows.map((f) => (
          <button key={f.name} onClick={() => onSelect(f)} className="mc-glass group rounded-xl p-4 text-left">
            <div className="flex items-center justify-between gap-2">
              <span className="font-mono text-sm font-bold text-white group-hover:text-[#5EEAD4]">{f.name}</span>
              <span className="rounded-full border border-white/12 px-2 py-0.5 text-[10px] text-slate-400">{f.category}</span>
            </div>
            <p className="mt-2 text-sm leading-6 text-slate-400">{f.desc}</p>
            <p className="mt-2 truncate font-mono text-[11px] text-emerald-300">{f.example}</p>
          </button>
        ))}
      </div>
    </div>
  );
}

function CheatSheetView() {
  const lesson = bySlug("cheat-sheet");
  return (
    <div>
      <h1 className="text-3xl font-black text-white md:text-4xl">Excel Online Cheat Sheet</h1>
      <p className="mt-3 max-w-3xl text-slate-300">
        One page to keep open beside your work: the shortcuts that actually save time, the formula patterns you will
        reuse every week, and a decoder for every Excel error message.
      </p>

      <section className="mt-10">
        <h2 className="text-xl font-bold text-white">Keyboard shortcuts</h2>
        <div className="mt-4 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
          {(lesson?.shortcuts ?? []).map(([k, v]) => (
            <div key={k} className="flex items-center gap-3 rounded-lg border border-white/10 bg-white/[0.03] px-3 py-2">
              <kbd className="shrink-0 rounded border border-white/15 bg-black/40 px-2 py-1 font-mono text-xs text-[#5EEAD4]">{k}</kbd>
              <span className="text-sm text-slate-300">{v}</span>
            </div>
          ))}
        </div>
      </section>

      <section className="mt-12">
        <h2 className="text-xl font-bold text-white">Formula patterns</h2>
        <div className="mt-4 grid gap-3 md:grid-cols-2">
          {(lesson?.formulas ?? []).map((f) => (
            <div key={f.label} className="mc-glass rounded-xl p-4">
              <p className="text-sm font-medium text-slate-300">{f.label}</p>
              <pre className="mt-1.5 overflow-x-auto whitespace-pre-wrap font-mono text-xs text-emerald-300">{f.code}</pre>
            </div>
          ))}
        </div>
      </section>

      <section className="mt-12">
        <h2 className="text-xl font-bold text-white">Error message decoder</h2>
        <div className="mt-4 overflow-x-auto rounded-xl border border-white/10">
          <table className="w-full border-collapse text-sm">
            <thead>
              <tr>
                <th className="border-b border-white/10 bg-white/8 px-3 py-2 text-left font-semibold text-slate-200">Error</th>
                <th className="border-b border-white/10 bg-white/8 px-3 py-2 text-left font-semibold text-slate-200">Usually means</th>
                <th className="border-b border-white/10 bg-white/8 px-3 py-2 text-left font-semibold text-slate-200">Fix</th>
              </tr>
            </thead>
            <tbody>
              {ERROR_DECODER.map(([code, meaning, fix]) => (
                <tr key={code} className="odd:bg-white/[0.02]">
                  <td className="border-b border-white/5 px-3 py-2 align-top font-mono font-semibold text-rose-300">{code}</td>
                  <td className="border-b border-white/5 px-3 py-2 align-top text-slate-300">{meaning}</td>
                  <td className="border-b border-white/5 px-3 py-2 align-top font-mono text-xs text-emerald-300">{fix}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

export default function AcademyHub({ onExit }: { onExit: () => void }) {
  const [nav, setNav] = useState<AcademyNav>("curriculum");
  const [activeLessonSlug, setActiveLessonSlug] = useState<string | null>(null);
  const [selectedFn, setSelectedFn] = useState<FnRef | null>(null);
  const [completed, setCompleted] = useState<Set<string>>(() => getCompleted());

  function openLesson(slug: string) {
    setActiveLessonSlug(slug);
  }
  function toggleComplete(slug: string) {
    setCompleted((prev) => {
      const next = new Set(prev);
      if (next.has(slug)) next.delete(slug);
      else next.add(slug);
      saveCompleted(next);
      return next;
    });
  }

  const activeLesson = activeLessonSlug ? bySlug(activeLessonSlug) : null;

  return (
    <div className="ws-root ws-scope relative min-h-screen">
      <GlobalHeader />
      <div className="mx-auto max-w-7xl px-5 pb-20 pt-28 lg:px-8">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="mc-display text-xl font-bold text-white">Free Excel Academy</h2>
            <p className="mt-1 max-w-3xl text-sm leading-6 text-[#b0bacb]">
              A structured, free Excel course built into MarqClean AI — {LESSONS.length} guided lessons, a full function
              reference, and a cheat sheet. Runs entirely in your browser; your progress is saved on this device.
            </p>
          </div>
          <button className="ws-btn-secondary rounded-full px-4 py-2 text-sm font-medium" onClick={onExit}>Back to MarqClean AI</button>
        </div>

        <div className="mt-6">
          <p className="mc-mono text-[11px] uppercase tracking-[0.2em] text-[#8a94a8]">Modules Available ({NAV_ITEMS.length})</p>
          <div className="mt-3">
            <ToolLaunchpad tools={NAV_ITEMS} active={nav} onSelect={(key) => { setNav(key); setActiveLessonSlug(null); setSelectedFn(null); }} />
          </div>
        </div>

        <main id="main-content" className="mt-8 min-w-0">
          {nav === "curriculum" ? (
            activeLesson ? (
              <LessonDetail lesson={activeLesson} onOpenLesson={openLesson} onBack={() => setActiveLessonSlug(null)} completed={completed} onToggleComplete={toggleComplete} />
            ) : (
              <CurriculumView onOpenLesson={openLesson} completed={completed} />
            )
          ) : nav === "functions" ? (
            <FunctionsView selected={selectedFn} onSelect={setSelectedFn} />
          ) : (
            <CheatSheetView />
          )}
        </main>
      </div>
    </div>
  );
}
