import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { GlobalHeader, PageHeader } from "./WorkspaceShell";
import {
  CHEAT_CATEGORIES,
  CHEAT_ENTRIES,
  XLSX_LESSONS,
  lessonBySlug,
  lessonNeighbours,
  renderFormula,
  type CheatCategory,
  type XlDiagram,
  type XlFunctionDoc,
  type XlLesson,
  type XlSection,
  type XlTable,
} from "./excel-learning";
import { buildWorkbook } from "./xlsx-builder";

const ROUTE = "excel-learning";
const CHEAT_SLUG = "formula-cheat-sheet";
const PROGRESS_KEY = "marqclean.excel-learning.completed";
const SEPARATOR_KEY = "marqclean.excel-learning.separator";
const SITE_URL = "https://marqcleanai.vertexsg.co.za";

const DISCLAIMER =
  "MarqClean AI supports data preparation and reconciliation. It does not replace professional accounting, audit, legal, tax, credit, medical or compliance advice. Users must review outputs before relying on them.";

const SERVICES = [
  { label: "Book Excel training", href: "/contact", detail: "A session built around your own files and the reports you actually produce." },
  { label: "Request a dashboard consultation", href: "/dashboarding", detail: "Scoped design and build for a repeatable management report." },
  { label: "Get help cleaning your data", href: "/data-automation", detail: "A documented, repeatable preparation workflow instead of a monthly manual fix." },
];

const LANDING_FAQ = [
  {
    q: "Is this course free?",
    a: "Yes. Every lesson, quiz and practice workbook is free and needs no account. The practice workbooks are generated in your browser when you download them.",
  },
  {
    q: "What Excel version do I need?",
    a: "Everything in Lessons 1 to 4, 6, 7 and 10 works in any modern version. Where a function or feature needs a specific release — XLOOKUP, IFS, the built-in Histogram chart, Power Query — the lesson says so explicitly.",
  },
  {
    q: "How long does the course take?",
    a: "Roughly two and a half hours across the ten lessons if you work through the practice exercises. Each lesson is self-contained, so you can jump to the one you need.",
  },
  {
    q: "Do my files leave my browser?",
    a: "No file upload is involved anywhere in this section. The practice workbooks are built in your browser at the moment you click download, and your progress is stored in this browser's local storage only.",
  },
  {
    q: "How is this different from the Excel Academy?",
    a: "The Excel Academy is a broad A-Z curriculum with a full function reference. Excel for Data Analysis is a focused ten-lesson track that follows one dataset from raw export through to a reported result, with a downloadable practice workbook per lesson.",
  },
  {
    q: "Are the formulas verified?",
    a: "Every calculated result in these lessons is computed from the practice dataset and checked arithmetically. Version-availability statements follow documented Microsoft behaviour. If you find a discrepancy in your version, tell us and we will correct it.",
  },
];

/* ------------------------------------------------------------------ *
 * .xlsx generation — built in the browser, no upload, no dependency
 * ------------------------------------------------------------------ */

function downloadWorkbook(lesson: XlLesson) {
  const bytes = buildWorkbook(lesson.workbookSheets);
  const blob = new Blob([bytes as unknown as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `marqclean-excel-for-data-analysis-lesson-${lesson.n}-${lesson.slug}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 250);
}

/* ------------------------------------------------------------------ *
 * Routing and metadata
 * ------------------------------------------------------------------ */

type SubPage = { kind: "landing" } | { kind: "lesson"; slug: string } | { kind: "cheat" } | { kind: "missing" };

function readSubPage(): SubPage {
  const raw = (window.location.pathname || "/")
    .replace(/^\/+|\/+$/g, "")
    .toLowerCase();
  const parts = raw.split("/").filter(Boolean);
  if (parts.length <= 1) return { kind: "landing" };
  const sub = parts[1];
  if (sub === CHEAT_SLUG) return { kind: "cheat" };
  return lessonBySlug(sub) ? { kind: "lesson", slug: sub } : { kind: "missing" };
}

function goTo(sub: string) {
  const path = sub ? `/${ROUTE}/${sub}` : `/${ROUTE}`;
  window.history.pushState({}, "", path);
  window.dispatchEvent(new PopStateEvent("popstate"));
  window.scrollTo({ top: 0 });
}

function setMeta(name: string, content: string, isProperty = false) {
  const attribute = isProperty ? "property" : "name";
  let element = document.head.querySelector<HTMLMetaElement>(`meta[${attribute}="${name}"]`);
  if (!element) {
    element = document.createElement("meta");
    element.setAttribute(attribute, name);
    document.head.appendChild(element);
  }
  element.setAttribute("content", content);
}

function setCanonical(href: string) {
  let link = document.head.querySelector<HTMLLinkElement>('link[rel="canonical"]');
  if (!link) {
    link = document.createElement("link");
    link.rel = "canonical";
    document.head.appendChild(link);
  }
  link.href = href;
}

function setSchema(id: string, payload: unknown) {
  let script = document.getElementById(id) as HTMLScriptElement | null;
  if (!script) {
    script = document.createElement("script");
    script.id = id;
    script.type = "application/ld+json";
    document.head.appendChild(script);
  }
  script.textContent = JSON.stringify(payload);
}

function removeSchema(id: string) {
  document.getElementById(id)?.remove();
}

/** Clear schema written by other routes so only one page's markup is present. */
function clearForeignSchema() {
  removeSchema("marqclean-service-schema");
}

const breadcrumbBase = [
  { name: "Home", url: `${SITE_URL}/` },
  { name: "Excel Learning", url: `${SITE_URL}/${ROUTE}` },
];

export function applyMetadata(page: SubPage) {
  clearForeignSchema();
  removeSchema("marqclean-learning-schema");
  removeSchema("marqclean-cheat-schema");

  const organisation = {
    "@type": "Organization",
    "@id": `${SITE_URL}/#organization`,
    name: "Vertex Stream Technologies",
    url: SITE_URL,
  };

  if (page.kind === "lesson") {
    const lesson = lessonBySlug(page.slug)!;
    const url = `${SITE_URL}/${ROUTE}/${lesson.slug}`;
    // metaTitle is the complete title tag, kept under 60 characters rather
    // than padded with a brand suffix that would push it over the SERP limit.
    document.title = lesson.metaTitle;
    setMeta("description", lesson.metaDescription);
    setMeta("robots", "index, follow");
    setMeta("og:title", lesson.metaTitle, true);
    setMeta("og:description", lesson.metaDescription, true);
    setMeta("og:type", "article", true);
    setMeta("og:url", url, true);
    setMeta("twitter:title", lesson.metaTitle);
    setMeta("twitter:description", lesson.metaDescription);
    setCanonical(url);
    setSchema("marqclean-learning-schema", {
      "@context": "https://schema.org",
      "@graph": [
        {
          "@type": "Article",
          headline: lesson.title,
          description: lesson.summary,
          url,
          inLanguage: "en-ZA",
          isPartOf: { "@id": `${SITE_URL}/${ROUTE}#course` },
          publisher: organisation,
        },
        {
          "@type": "LearningResource",
          "@id": `${url}#resource`,
          name: lesson.title,
          description: lesson.summary,
          url,
          educationalLevel: lesson.level,
          learningResourceType: "Lesson",
          inLanguage: "en-ZA",
          isPartOf: { "@id": `${SITE_URL}/${ROUTE}#course` },
          teaches: lesson.objectives,
          timeRequired: `PT${lesson.minutes}M`,
          hasPart: {
            "@type": "Quiz",
            name: `${lesson.title} knowledge check`,
            educationalUse: "practice",
            numberOfItems: lesson.quiz.length,
          },
        },
        {
          "@type": "BreadcrumbList",
          itemListElement: [...breadcrumbBase, { name: lesson.title, url }].map((item, index) => ({
            "@type": "ListItem",
            position: index + 1,
            name: item.name,
            item: item.url,
          })),
        },
        {
          "@type": "FAQPage",
          mainEntity: lesson.faq.map((entry) => ({
            "@type": "Question",
            name: entry.q,
            acceptedAnswer: { "@type": "Answer", text: entry.a },
          })),
        },
      ],
    });
    return;
  }

  if (page.kind === "cheat") {
    const heading = "Excel Formula Cheat Sheet";
    const title = "Excel Formula Cheat Sheet: Syntax, Examples & Versions";
    const description =
      "Searchable Excel formula cheat sheet: syntax, examples, results and version notes for the functions taught in Excel for Data Analysis.";
    const url = `${SITE_URL}/${ROUTE}/${CHEAT_SLUG}`;
    document.title = title;
    setMeta("description", description);
    setMeta("robots", "index, follow");
    setMeta("og:title", title, true);
    setMeta("og:description", description, true);
    setMeta("og:type", "website", true);
    setMeta("og:url", url, true);
    setCanonical(url);
    removeSchema("marqclean-learning-schema");
    setSchema("marqclean-cheat-schema", {
      "@context": "https://schema.org",
      "@type": "Reference",
      name: heading,
      description,
      url,
      inLanguage: "en-ZA",
      about: CHEAT_CATEGORIES,
      publisher: organisation,
    });
    return;
  }

  const heading = "Excel for Data Analysis: From Basics to Advanced";
  const title = "Excel for Data Analysis: 10 Free Lessons | MarqClean AI";
  const description =
    "A free 10-lesson Excel course for data analysis: basics, functions, lookups, data cleaning, PivotTables, charts, Power Query and interview questions.";
  const url = `${SITE_URL}/${ROUTE}`;
  document.title = title;
  setMeta("description", description);
  setMeta("robots", "index, follow");
  setMeta("og:title", title, true);
  setMeta("og:description", description, true);
  setMeta("og:type", "website", true);
  setMeta("og:url", url, true);
  setCanonical(url);
  removeSchema("marqclean-cheat-schema");
  setSchema("marqclean-learning-schema", {
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "Course",
        "@id": `${url}#course`,
        name: heading,
        description,
        url,
        inLanguage: "en-ZA",
        isAccessibleForFree: true,
        numberOfCredits: 0,
        provider: organisation,
        publisher: organisation,
        educationalLevel: "Beginner to Advanced",
        hasCourseInstance: {
          "@type": "CourseInstance",
          courseMode: "online",
          courseWorkload: "PT2H30M",
        },
        syllabusSections: XLSX_LESSONS.map((lesson, index) => ({
          "@type": "Syllabus",
          position: index + 1,
          name: `Lesson ${lesson.n}: ${lesson.title}`,
          description: lesson.summary,
          timeRequired: `PT${lesson.minutes}M`,
        })),
      },
      {
        "@type": "CollectionPage",
        name: heading,
        url,
        isPartOf: { "@id": `${SITE_URL}/#website` },
        about: { "@id": `${url}#course` },
      },
      {
        "@type": "BreadcrumbList",
        itemListElement: breadcrumbBase.map((item, index) => ({
          "@type": "ListItem",
          position: index + 1,
          name: item.name,
          item: item.url,
        })),
      },
      {
        "@type": "FAQPage",
        mainEntity: LANDING_FAQ.map((entry) => ({
          "@type": "Question",
          name: entry.q,
          acceptedAnswer: { "@type": "Answer", text: entry.a },
        })),
      },
    ],
  });
}

/* ------------------------------------------------------------------ *
 * Shared UI
 * ------------------------------------------------------------------ */

function Icon({ name }: { name: "copy" | "check" | "arrow" | "download" | "search" | "book" | "chart" | "spark" }) {
  const paths = {
    copy: (
      <>
        <rect x="8" y="8" width="12" height="12" rx="2" />
        <path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" />
      </>
    ),
    check: <path d="m4 12 5 5L20 6" />,
    arrow: (
      <>
        <path d="M5 12h14" />
        <path d="m13 6 6 6-6 6" />
      </>
    ),
    download: (
      <>
        <path d="M12 3v12" />
        <path d="m7 11 5 5 5-5" />
        <path d="M5 21h14" />
      </>
    ),
    search: (
      <>
        <circle cx="11" cy="11" r="7" />
        <path d="m20 20-4-4" />
      </>
    ),
    book: (
      <>
        <path d="M4 5a3 3 0 0 1 3-3h13v18H7a3 3 0 0 0-3 3V5Z" />
        <path d="M7 20h13" />
      </>
    ),
    chart: (
      <>
        <path d="M4 19V5M4 19h16" />
        <path d="m7 15 4-4 3 2 5-7" />
      </>
    ),
    spark: <path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5 18 18M18 6l-2.5 2.5M8.5 15.5 6 18" />,
  };
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      {paths[name]}
    </svg>
  );
}

function CopyButton({ text, separator, label = "Copy" }: { text: string; separator: ";" | ","; label?: string }) {
  const [state, setState] = useState<"idle" | "done" | "error">("idle");
  const timer = useRef<number | null>(null);
  const value = renderFormula(text, separator);

  useEffect(() => () => { if (timer.current) window.clearTimeout(timer.current); }, []);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(value);
      setState("done");
    } catch {
      setState("error");
    }
    if (timer.current) window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => setState("idle"), 2200);
  };

  return (
    <button type="button" className="xl-copy" onClick={() => { void copy(); }}>
      {state === "done" ? <Icon name="check" /> : <Icon name="copy" />}
      <span>{state === "done" ? "Copied" : state === "error" ? "Press Ctrl+C" : label}</span>
      <span className="sr-only" aria-live="polite">
        {state === "done" ? `${value} copied to clipboard` : state === "error" ? "Clipboard unavailable, select the text and press Ctrl+C" : ""}
      </span>
    </button>
  );
}

function CodeBlock({ code, separator, caption }: { code: string; separator: ";" | ","; caption?: string }) {
  const rendered = renderFormula(code, separator);
  return (
    <div className="xl-code">
      {caption ? <span className="xl-code__caption">{caption}</span> : null}
      <div className="xl-code__row">
        <pre tabIndex={0}>
          <code>{rendered}</code>
        </pre>
        <CopyButton text={code} separator={separator} />
      </div>
    </div>
  );
}

function SeparatorControl({ separator, onChange }: { separator: ";" | ","; onChange: (value: ";" | ",") => void }) {
  return (
    <div className="xl-separator" role="group" aria-label="Formula list separator">
      <span className="xl-separator__label">Argument separator</span>
      <button type="button" aria-pressed={separator === ";"} onClick={() => onChange(";")}>
        Semicolon
      </button>
      <button type="button" aria-pressed={separator === ","} onClick={() => onChange(",")}>
        Comma
      </button>
    </div>
  );
}

function DataTable({ table }: { table: XlTable }) {
  return (
    <div className="xl-table-wrap">
      {table.caption ? <p className="xl-table-caption">{table.caption}</p> : null}
      <table>
        <thead>
          <tr>
            {table.headers.map((header, index) => (
              <th key={`${header}-${index}`} scope="col">
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {table.rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) =>
                cellIndex === 0 ? (
                  <th key={cellIndex} scope="row">
                    {cell}
                  </th>
                ) : (
                  <td key={cellIndex} className={table.resultColumns?.includes(cellIndex) ? "is-result" : undefined}>
                    {cell}
                  </td>
                ),
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Diagram({ kind, caption }: { kind: XlDiagram; caption: string }) {
  const stroke = "rgb(var(--accent))";
  const line = "rgb(var(--line-strong))";
  const text = "rgb(var(--ink-2))";

  if (kind === "etl") {
    return (
      <figure className="xl-diagram">
        <svg viewBox="0 0 720 190" role="img" aria-label="Power Query workflow: Extract from a source, Transform in the editor, Load to a worksheet">
          {[
            { x: 20, title: "Extract", body: "Connect to CSV, Excel,\na folder, a database\nor the web" },
            { x: 260, title: "Transform", body: "Promote headers,\nchange types, split,\nfilter, unpivot" },
            { x: 500, title: "Load", body: "Table, PivotTable\nor the Data Model" },
          ].map((stage, index) => (
            <g key={stage.title}>
              <rect x={stage.x} y="30" width="200" height="120" rx="10" fill="rgb(var(--surface))" stroke={index === 1 ? stroke : line} strokeWidth={index === 1 ? 2 : 1} />
              <text x={stage.x + 100} y="62" textAnchor="middle" fontSize="17" fontWeight="700" fill="rgb(var(--ink))">
                {stage.title}
              </text>
              {stage.body.split("\n").map((lineText, lineIndex) => (
                <text key={lineIndex} x={stage.x + 100} y={88 + lineIndex * 19} textAnchor="middle" fontSize="12.5" fill={text}>
                  {lineText}
                </text>
              ))}
              {index < 2 ? (
                <>
                  <line x1={stage.x + 200} y1="90" x2={stage.x + 258} y2="90" stroke={line} strokeWidth="2" />
                  <path d={`M${stage.x + 250} 84 l10 6 l-10 6 z`} fill={line} />
                </>
              ) : null}
            </g>
          ))}
          <path d="M600 150 q60 30 -120 30 q-180 0 -330 -22" fill="none" stroke={stroke} strokeWidth="1.5" strokeDasharray="5 5" />
          <text x="360" y="176" textAnchor="middle" fontSize="12" fill={stroke}>
            Refresh replays every step
          </text>
        </svg>
        <figcaption>{caption}</figcaption>
      </figure>
    );
  }

  if (kind === "lookup") {
    const steps = [
      "1. Know the value you have",
      "2. Clean it — TRIM spaces, fix types",
      "3. Find the table that holds the answer",
      "4. Choose the direction and the function",
      "5. Decide what happens when nothing matches",
    ];
    return (
      <figure className="xl-diagram">
        <svg viewBox="0 0 720 150" role="img" aria-label="Five-step lookup workflow, from the value you have to handling a missing match">
          {steps.map((step, index) => {
            const x = 20 + index * 140;
            return (
              <g key={step}>
                <rect x={x} y="35" width="126" height="76" rx="10" fill="rgb(var(--surface))" stroke={index === 1 || index === 4 ? stroke : line} strokeWidth={index === 1 || index === 4 ? 2 : 1} />
                {step.split(". ").map((part, partIndex) => (
                  <text key={partIndex} x={x + 63} y={partIndex === 0 ? 66 : 88} textAnchor="middle" fontSize={partIndex === 0 ? 15 : 12} fontWeight={partIndex === 0 ? 700 : 400} fill={partIndex === 0 ? "rgb(var(--ink))" : text}>
                    {partIndex === 0 ? part : part}
                  </text>
                ))}
                {index < 4 ? <path d={`M${x + 128} 73 l10 0`} stroke={line} strokeWidth="2" /> : null}
              </g>
            );
          })}
          <text x="360" y="136" textAnchor="middle" fontSize="12" fill={stroke}>
            Most lookup failures happen at step 2 or step 5
          </text>
        </svg>
        <figcaption>{caption}</figcaption>
      </figure>
    );
  }

  if (kind === "merge-append") {
    return (
      <figure className="xl-diagram">
        <svg viewBox="0 0 720 250" role="img" aria-label="Append stacks two tables vertically and adds rows. Merge joins two tables horizontally on a key and adds columns.">
          <text x="180" y="24" textAnchor="middle" fontSize="16" fontWeight="700" fill="rgb(var(--ink))">
            Append — adds rows
          </text>
          <rect x="60" y="40" width="110" height="60" rx="6" fill="rgb(var(--surface))" stroke={line} />
          <rect x="60" y="106" width="110" height="60" rx="6" fill="rgb(var(--surface))" stroke={line} />
          <rect x="190" y="40" width="110" height="126" rx="6" fill="rgb(var(--accent-tint))" stroke={stroke} strokeWidth="2" />
          <text x="115" y="75" textAnchor="middle" fontSize="12" fill={text}>10 rows</text>
          <text x="115" y="141" textAnchor="middle" fontSize="12" fill={text}>10 rows</text>
          <text x="245" y="98" textAnchor="middle" fontSize="13" fontWeight="700" fill="rgb(var(--ink))">20 rows</text>
          <text x="245" y="118" textAnchor="middle" fontSize="12" fill={text}>6 columns</text>
          <path d="M176 100 h10" stroke={line} strokeWidth="2" />

          <text x="540" y="24" textAnchor="middle" fontSize="16" fontWeight="700" fill="rgb(var(--ink))">
            Merge — adds columns
          </text>
          <rect x="420" y="40" width="110" height="126" rx="6" fill="rgb(var(--surface))" stroke={line} />
          <rect x="550" y="40" width="80" height="60" rx="6" fill="rgb(var(--surface))" stroke={line} />
          <rect x="420" y="186" width="210" height="44" rx="6" fill="rgb(var(--accent-tint))" stroke={stroke} strokeWidth="2" />
          <text x="475" y="100" textAnchor="middle" fontSize="12" fill={text}>10 rows</text>
          <text x="590" y="75" textAnchor="middle" fontSize="12" fill={text}>3 rows</text>
          <text x="525" y="207" textAnchor="middle" fontSize="13" fontWeight="700" fill="rgb(var(--ink))">10 rows · 7 columns</text>
          <text x="525" y="223" textAnchor="middle" fontSize="11.5" fill={text}>joined on the key</text>
        </svg>
        <figcaption>{caption}</figcaption>
      </figure>
    );
  }

  return (
    <figure className="xl-diagram">
      <svg viewBox="0 0 720 210" role="img" aria-label="Chart choice: change over time uses a line, comparison between categories uses a bar, distribution uses a histogram, relationships use a scatter">
        {[
          { x: 20, q: "Change over time?", a: "Line", path: "M14 62 L34 52 L54 56 L74 40 L94 30" },
          { x: 200, q: "Compare categories?", a: "Bar", path: "" },
          { x: 380, q: "Show a distribution?", a: "Histogram", path: "" },
          { x: 560, q: "Two variables related?", a: "Scatter", path: "" },
        ].map((card) => (
          <g key={card.a}>
            <rect x={card.x} y="24" width="140" height="150" rx="10" fill="rgb(var(--surface))" stroke={line} />
            <text x={card.x + 70} y="48" textAnchor="middle" fontSize="11.5" fill={text}>
              {card.q}
            </text>
            {card.a === "Line" ? <path d={card.path} transform={`translate(${card.x},0)`} fill="none" stroke={stroke} strokeWidth="2.5" /> : null}
            {card.a === "Bar"
              ? [40, 58, 30, 48].map((height, index) => (
                  <rect key={index} x={card.x + 16 + index * 28} y={140 - height} width="18" height={height} fill={index === 1 ? stroke : "rgb(var(--line-strong))"} />
                ))
              : null}
            {card.a === "Histogram"
              ? [18, 46, 22, 12, 4].map((height, index) => (
                  <rect key={index} x={card.x + 14 + index * 23} y={140 - height} width="23" height={height} fill={index === 1 ? stroke : "rgb(var(--line-strong))"} />
                ))
              : null}
            {card.a === "Scatter"
              ? [
                  [22, 128],
                  [40, 112],
                  [58, 96],
                  [76, 82],
                  [96, 62],
                  [112, 50],
                ].map(([cx, cy], index) => <circle key={index} cx={card.x + cx} cy={cy} r="4.5" fill={index > 2 ? stroke : "rgb(var(--line-strong))"} />)
              : null}
            <line x1={card.x + 14} y1="141" x2={card.x + 126} y2="141" stroke={line} strokeWidth="1.5" />
            <text x={card.x + 70} y="163" textAnchor="middle" fontSize="14" fontWeight="700" fill="rgb(var(--ink))">
              {card.a}
            </text>
          </g>
        ))}
      </svg>
      <figcaption>{caption}</figcaption>
    </figure>
  );
}

function FunctionDoc({ doc, separator }: { doc: XlFunctionDoc; separator: ";" | "," }) {
  return (
    <article className="xl-function">
      <div className="xl-function__head">
        <h3>{doc.name}</h3>
        <span className="ws-format-badge">{doc.category}</span>
      </div>
      <p className="xl-function__purpose">{doc.purpose}</p>
      <CodeBlock code={doc.syntax} separator={separator} caption="Syntax" />
      {doc.version ? <p className="xl-version"><strong>Availability:</strong> {doc.version}</p> : null}
      <div className="xl-examples">
        <span className="xl-eyebrow">Worked examples</span>
        <table>
          <thead>
            <tr>
              <th scope="col">What you want</th>
              <th scope="col">Formula</th>
              <th scope="col">Result</th>
              <th scope="col">
                <span className="sr-only">Copy this formula</span>
              </th>
            </tr>
          </thead>
          <tbody>
            {doc.examples.map((example) => (
              <tr key={example.label}>
                <th scope="row">{example.label}</th>
                <td>
                  <code>{renderFormula(example.formula, separator)}</code>
                </td>
                <td className="is-result">{example.result}</td>
                <td className="xl-examples__copy">
                  <CopyButton text={example.formula} separator={separator} label="" />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="xl-two-col">
        <div className="xl-box xl-box--key">
          <span className="xl-eyebrow">Key points</span>
          <ul>
            {doc.keyPoints.map((point) => (
              <li key={point}>{point}</li>
            ))}
          </ul>
        </div>
        <div className="xl-box xl-box--warn">
          <span className="xl-eyebrow">Common mistakes</span>
          <ul>
            {doc.mistakes.map((mistake) => (
              <li key={mistake}>{mistake}</li>
            ))}
          </ul>
        </div>
      </div>
    </article>
  );
}

function Section({ section, separator }: { section: XlSection; separator: ";" | "," }) {
  switch (section.kind) {
    case "prose":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          {section.body.map((paragraph) => (
            <p key={paragraph.slice(0, 40)}>{paragraph}</p>
          ))}
        </section>
      );
    case "functions":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          {section.intro ? <p>{section.intro}</p> : null}
          {section.functions.map((doc) => (
            <FunctionDoc key={doc.name} doc={doc} separator={separator} />
          ))}
        </section>
      );
    case "table":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          <DataTable table={section.table} />
          {section.note ? <p className="xl-note">{section.note}</p> : null}
        </section>
      );
    case "steps":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          <ol className="xl-steps">
            {section.steps.map((step, index) => (
              <li key={step.title}>
                <span aria-hidden="true">{index + 1}</span>
                <div>
                  <h3>{step.title}</h3>
                  <p>{step.detail}</p>
                </div>
              </li>
            ))}
          </ol>
        </section>
      );
    case "diagram":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          <Diagram kind={section.diagram} caption={section.caption} />
        </section>
      );
    case "callout":
      return (
        <section className="xl-section" id={section.id} aria-labelledby={`${section.id}-h`}>
          <h2 id={`${section.id}-h`}>{section.heading}</h2>
          <ul className="xl-points">
            {section.points.map((point) => (
              <li key={point}>{point}</li>
            ))}
          </ul>
        </section>
      );
    case "practice":
      return <PracticeExercise heading={section.heading} task={section.task} answer={section.answer} id={section.id} />;
    default:
      return null;
  }
}

function PracticeExercise({ heading, task, answer, id }: { heading: string; task: string; answer: string; id: string }) {
  const [revealed, setRevealed] = useState(false);
  return (
    <section className="xl-section xl-practice" id={id} aria-labelledby={`${id}-h`}>
      <h2 id={`${id}-h`}>{heading}</h2>
      <p>{task}</p>
      <button type="button" className="ws-btn-secondary" aria-expanded={revealed} onClick={() => setRevealed((value) => !value)}>
        {revealed ? "Hide model answer" : "Show model answer"}
      </button>
      {revealed ? <p className="xl-answer">{answer}</p> : null}
    </section>
  );
}

function Quiz({ lesson, onPass }: { lesson: XlLesson; onPass: () => void }) {
  const [picked, setPicked] = useState<Record<number, number>>({});
  const [checked, setChecked] = useState(false);
  const score = lesson.quiz.reduce((total, question, index) => total + (picked[index] === question.answer ? 1 : 0), 0);
  const passed = checked && score === lesson.quiz.length;
  const announced = useRef(false);

  useEffect(() => {
    if (passed && !announced.current) {
      announced.current = true;
      onPass();
    }
  }, [passed, onPass]);

  useEffect(() => {
    setPicked({});
    setChecked(false);
    announced.current = false;
  }, [lesson.slug]);

  const answered = Object.keys(picked).length;

  return (
    <section className="xl-section xl-quiz" id="quiz" aria-labelledby="quiz-h">
      <h2 id="quiz-h">Knowledge check</h2>
      <p>{lesson.quiz.length} questions. Answer all of them, then check — the explanations are where the detail is.</p>
      {lesson.quiz.map((question, index) => (
        <fieldset key={question.q} className="xl-question" disabled={checked}>
          <legend>
            {index + 1}. {question.q}
          </legend>
          {question.options.map((option, optionIndex) => (
            <label key={option} className={picked[index] === optionIndex ? "is-picked" : undefined}>
              <input
                type="radio"
                name={`${lesson.slug}-q${index}`}
                checked={picked[index] === optionIndex}
                onChange={() => setPicked((value) => ({ ...value, [index]: optionIndex }))}
              />
              <span>{option}</span>
            </label>
          ))}
          {checked ? (
            <p className={picked[index] === question.answer ? "xl-feedback is-right" : "xl-feedback is-wrong"}>
              <strong>{picked[index] === question.answer ? "Correct. " : "Not quite. "}</strong>
              {question.explain}
            </p>
          ) : null}
        </fieldset>
      ))}
      <div className="xl-quiz__actions">
        {checked ? (
          <>
            <p className="xl-score" role="status">
              You scored {score} out of {lesson.quiz.length}.
              {passed ? " All correct — this lesson is marked complete." : " Read the explanations and try again."}
            </p>
            <button
              type="button"
              className="ws-btn-secondary"
              onClick={() => {
                setPicked({});
                setChecked(false);
              }}
            >
              Try again
            </button>
          </>
        ) : (
          <button type="button" className="ws-btn-primary" disabled={answered < lesson.quiz.length} onClick={() => setChecked(true)}>
            Check my answers
          </button>
        )}
        {checked ? null : (
          <p className="xl-quiz__hint" role="status">
            {answered} of {lesson.quiz.length} answered
          </p>
        )}
      </div>
    </section>
  );
}

function ServiceCta() {
  return (
    <section className="xl-services" aria-labelledby="xl-services-h">
      <h2 id="xl-services-h">Need help applying this to your own data?</h2>
      <p>
        These lessons teach the technique. If you need the result on your files, MarqClean AI runs scoped engagements alongside the free tools.
      </p>
      <div className="xl-services__grid">
        {SERVICES.map((service) => (
          <a key={service.href} className="xl-service" href={service.href}>
            <strong>{service.label}</strong>
            <span>{service.detail}</span>
            <span className="xl-service__go">
              {service.label} <Icon name="arrow" />
            </span>
          </a>
        ))}
      </div>
    </section>
  );
}

function Breadcrumbs({ trail }: { trail: { name: string; href: string }[] }) {
  return (
    <nav className="xl-breadcrumb" aria-label="Breadcrumb">
      <ol>
        {trail.map((item, index) => (
          <li key={item.href}>
            {index === trail.length - 1 ? (
              <span aria-current="page">{item.name}</span>
            ) : (
              <>
                <a
                  href={item.href}
                  onClick={(event) => {
                    event.preventDefault();
                    goTo(item.href.replace(`/${ROUTE}`, "").replace(/^\//, ""));
                  }}
                >
                  {item.name}
                </a>
                <span aria-hidden="true">/</span>
              </>
            )}
          </li>
        ))}
      </ol>
    </nav>
  );
}

function TrustBar() {
  return (
    <ul className="xl-trust">
      <li>Files processed locally in your browser</li>
      <li>No account required</li>
      <li>Review before export</li>
      <li>Your original file stays unchanged</li>
      <li>100 MB browser-safe limit</li>
      <li>POPIA and GDPR conscious</li>
    </ul>
  );
}

function Disclaimer() {
  return <p className="xl-disclaimer">{DISCLAIMER}</p>;
}

/* ------------------------------------------------------------------ *
 * Pages
 * ------------------------------------------------------------------ */

function LessonCard({ lesson, done, onOpen }: { lesson: XlLesson; done: boolean; onOpen: (slug: string) => void }) {
  return (
    <a className="xl-card" href={`/${ROUTE}/${lesson.slug}`} onClick={(event) => { event.preventDefault(); onOpen(lesson.slug); }}>
      <div className="xl-card__top">
        <span className="xl-card__number">Lesson {lesson.n}</span>
        <span className={done ? "xl-card__state is-done" : "xl-card__state"}>{done ? "Complete" : lesson.level}</span>
      </div>
      <h3>{lesson.title}</h3>
      <p>{lesson.summary}</p>
      <div className="xl-card__bottom">
        <span>{lesson.minutes} min</span>
        <span>{lesson.quiz.length}-question quiz</span>
        <span>Practice workbook</span>
      </div>
      <span className="xl-card__cta">
        {done ? "Review lesson" : "Start lesson"} <Icon name="arrow" />
      </span>
    </a>
  );
}

function Landing({ completed, onOpen, separator, onSeparator }: {
  completed: Set<string>;
  onOpen: (slug: string) => void;
  separator: ";" | ",";
  onSeparator: (value: ";" | ",") => void;
}) {
  const next = XLSX_LESSONS.find((lesson) => !completed.has(lesson.slug));
  const percent = Math.round((completed.size / XLSX_LESSONS.length) * 100);
  const totalMinutes = XLSX_LESSONS.reduce((sum, lesson) => sum + lesson.minutes, 0);

  return (
    <>
      <PageHeader
        workspace="Excel Learning"
        title="Excel for Data Analysis: From Basics to Advanced"
        description="Ten lessons that take one messy dataset from a raw export through to a result you can report. Free, with a practice workbook for every lesson."
      />
      <Breadcrumbs trail={[{ name: "Home", href: "/" }, { name: "Excel Learning", href: `/${ROUTE}` }]} />

      <section className="xl-hero" aria-labelledby="xl-outcome-h">
        <div>
          <span className="xl-eyebrow">What you will be able to do</span>
          <h2 id="xl-outcome-h">Turn an export into an answer, without guessing</h2>
          <p>
            Most Excel teaching stops at syntax. This course follows one dataset the whole way: clean it, look things up in it, summarise it, chart it, and
            defend the numbers when someone asks how you got them. Every formula is shown with its result, and every result is calculated from the practice
            file you download.
          </p>
          <div className="xl-hero__actions">
            <button type="button" className="ws-btn-primary" onClick={() => onOpen(next ? next.slug : XLSX_LESSONS[0].slug)}>
              {next ? `Start Lesson ${next.n}: ${next.title.split(":")[0]}` : "Review Lesson 1"} <Icon name="arrow" />
            </button>
            <button type="button" className="ws-btn-secondary" onClick={() => goTo(CHEAT_SLUG)}>
              Open the formula cheat sheet
            </button>
          </div>
          <SeparatorControl separator={separator} onChange={onSeparator} />
        </div>
        <div className="xl-hero__meta">
          <dl>
            <div>
              <dt>Lessons</dt>
              <dd>{XLSX_LESSONS.length}</dd>
            </div>
            <div>
              <dt>Total time</dt>
              <dd>About {Math.round(totalMinutes / 60 * 10) / 10} hours</dd>
            </div>
            <div>
              <dt>Functions covered</dt>
              <dd>{CHEAT_ENTRIES.length}</dd>
            </div>
            <div>
              <dt>Practice workbooks</dt>
              <dd>{XLSX_LESSONS.length} .xlsx files</dd>
            </div>
          </dl>
        </div>
      </section>

      <TrustBar />

      <section className="xl-section" aria-labelledby="xl-audience-h">
        <h2 id="xl-audience-h">Who this is for</h2>
        <p>
          Analysts, finance and operations staff, bookkeepers, and anyone who inherits a spreadsheet and has to make it mean something. It assumes you can open
          Excel and type into a cell, and nothing more. If you are preparing for an analyst interview, Lesson 10 collects the questions that come up most often.
        </p>
      </section>

      <section className="xl-progress" aria-labelledby="xl-progress-h">
        <div>
          <span className="xl-eyebrow">Your progress</span>
          <h2 id="xl-progress-h">
            {completed.size} of {XLSX_LESSONS.length} lessons complete
          </h2>
          <p>Saved in this browser only. Nothing is sent anywhere and no account is involved.</p>
        </div>
        <div className="xl-progress__bar" role="progressbar" aria-valuenow={percent} aria-valuemin={0} aria-valuemax={100} aria-label="Course progress">
          <span style={{ width: `${percent}%` }} />
        </div>
        <strong>{percent}%</strong>
      </section>

      <section className="xl-section" aria-labelledby="xl-lessons-h">
        <h2 id="xl-lessons-h">The ten lessons</h2>
        <p>Each lesson stands alone, so you can jump straight to the one you need. They are ordered so that each one builds on the dataset introduced in Lesson 1.</p>
        <div className="xl-card-grid">
          {XLSX_LESSONS.map((lesson) => (
            <LessonCard key={lesson.slug} lesson={lesson} done={completed.has(lesson.slug)} onOpen={onOpen} />
          ))}
        </div>
      </section>

      <section className="xl-section xl-cheat-cta" aria-labelledby="xl-cheat-h">
        <div>
          <span className="xl-eyebrow">Reference</span>
          <h2 id="xl-cheat-h">Formula cheat sheet</h2>
          <p>
            Every function from the course on one searchable page: purpose, syntax, an example, its result and the Excel versions that support it. Search by
            function name, category or the problem you are trying to solve, and copy any formula in your own argument separator.
          </p>
          <button type="button" className="ws-btn-primary" onClick={() => goTo(CHEAT_SLUG)}>
            Open the formula cheat sheet <Icon name="arrow" />
          </button>
        </div>
        <div className="xl-cheat-cta__preview">
          {CHEAT_CATEGORIES.map((category) => (
            <span key={category}>{category}</span>
          ))}
        </div>
      </section>

      <ServiceCta />

      <section className="xl-section" aria-labelledby="xl-tools-h">
        <h2 id="xl-tools-h">Do it in the browser instead</h2>
        <p>
          Every technique in this course has a tool that runs it on your own file, locally in the browser, with a review step before you export anything. No
          account, and the file is not uploaded to a server.
        </p>
        <ul className="xl-tool-links">
          <li>
            <a href="/">Clean a messy CSV or Excel file</a> — the cleaning sequence from Lesson 6, applied to a whole file.
          </li>
          <li>
            <a href="/remove-duplicates-excel">Remove duplicates</a> — with the trimming that has to happen first.
          </li>
          <li>
            <a href="/fuzzy-name-matching">Fuzzy match records</a> — for the near-matches that no lookup function can resolve.
          </li>
          <li>
            <a href="/sumifs-countifs-tool">SUMIFS and COUNTIFS</a> — the conditional totals from Lesson 4, generated for you.
          </li>
          <li>
            <a href="/xlookup-online">XLOOKUP and INDEX/MATCH</a> — the lookups from Lesson 5.
          </li>
          <li>
            <a href="/excel-file-comparison">Compare two Excel files</a> — what changed between this month and last.
          </li>
          <li>
            <a href="/bank-statement-converter">Convert a PDF bank statement</a> — into Excel or CSV you can actually total.
          </li>
          <li>
            <a href="/excel-academy">Excel Academy</a> — the broader A-Z curriculum and full function reference.
          </li>
        </ul>
      </section>

      <section className="xl-section" aria-labelledby="xl-faq-h">
        <h2 id="xl-faq-h">Questions about the course</h2>
        {LANDING_FAQ.map((entry) => (
          <details key={entry.q} className="xl-faq">
            <summary>{entry.q}</summary>
            <p>{entry.a}</p>
          </details>
        ))}
      </section>

      <section className="xl-final" aria-labelledby="xl-final-h">
        <h2 id="xl-final-h">Start with the dataset, not the syntax</h2>
        <p>Download the Lesson 1 practice workbook and follow along. Ten lessons, about two and a half hours, one dataset all the way through.</p>
        <div className="xl-final__actions">
          <button type="button" className="ws-btn-primary" onClick={() => onOpen(XLSX_LESSONS[0].slug)}>
            Start Lesson 1 <Icon name="arrow" />
          </button>
          <button type="button" className="ws-btn-secondary" onClick={() => downloadWorkbook(XLSX_LESSONS[0])}>
            <Icon name="download" /> Download the practice workbook
          </button>
        </div>
        <Disclaimer />
      </section>
    </>
  );
}

function LessonPage({ lesson, completed, onToggle, onOpen, separator, onSeparator }: {
  lesson: XlLesson;
  completed: Set<string>;
  onToggle: (slug: string) => void;
  onOpen: (slug: string) => void;
  separator: ";" | ",";
  onSeparator: (value: ";" | ",") => void;
}) {
  const { prev, next } = lessonNeighbours(lesson.slug);
  const done = completed.has(lesson.slug);
  const [workbookState, setWorkbookState] = useState<"idle" | "done">("idle");
  const tocRef = useRef<HTMLDivElement>(null);

  const toc = useMemo(
    () =>
      lesson.sections
        .filter((section) => section.kind !== "practice")
        .map((section) => ({ id: section.id, heading: section.heading })),
    [lesson.slug],
  );

  const download = () => {
    downloadWorkbook(lesson);
    setWorkbookState("done");
    window.setTimeout(() => setWorkbookState("idle"), 4000);
  };

  const markComplete = useCallback(() => {
    if (!completed.has(lesson.slug)) onToggle(lesson.slug);
  }, [completed, lesson.slug, onToggle]);

  return (
    <>
      <PageHeader workspace={`Lesson ${lesson.n} of ${XLSX_LESSONS.length}`} title={lesson.title} description={lesson.summary} onBack={() => goTo("")} />
      <Breadcrumbs
        trail={[
          { name: "Home", href: "/" },
          { name: "Excel Learning", href: `/${ROUTE}` },
          { name: lesson.title, href: `/${ROUTE}/${lesson.slug}` },
        ]}
      />

      <div className="xl-lesson-grid">
        <div className="xl-lesson-sidebar" ref={tocRef}>
          <nav aria-label="Lesson contents">
            <span className="xl-eyebrow">On this page</span>
            <ol>
              <li>
                <a href="#objectives">Learning objectives</a>
              </li>
              {toc.map((entry) => (
                <li key={entry.id}>
                  <a href={`#${entry.id}`}>{entry.heading}</a>
                </li>
              ))}
              <li>
                <a href="#workbook">Practice workbook</a>
              </li>
              <li>
                <a href="#quiz">Knowledge check</a>
              </li>
              <li>
                <a href="#lesson-faq">Questions</a>
              </li>
            </ol>
          </nav>
          <div className="xl-lesson-sidebar__progress">
            <span className="xl-eyebrow">Course progress</span>
            <strong>
              Lesson {lesson.n} of {XLSX_LESSONS.length}
            </strong>
            <div
              className="xl-progress__bar"
              role="progressbar"
              aria-valuenow={lesson.n}
              aria-valuemin={1}
              aria-valuemax={XLSX_LESSONS.length}
              aria-label={`Lesson ${lesson.n} of ${XLSX_LESSONS.length}`}
            >
              <span style={{ width: `${(lesson.n / XLSX_LESSONS.length) * 100}%` }} />
            </div>
            <span className="xl-lesson-sidebar__state">{done ? "Completed on this device" : "Not yet completed"}</span>
          </div>
        </div>

        <article className="xl-lesson-content">
          <div className="xl-lesson-meta">
            <span className="ws-format-badge">{lesson.level}</span>
            <span>{lesson.minutes} minute read</span>
            <span>{lesson.quiz.length}-question quiz</span>
            <span>
              Lesson {lesson.n} of {XLSX_LESSONS.length}
            </span>
          </div>
          <SeparatorControl separator={separator} onChange={onSeparator} />

          <section className="xl-section" id="objectives" aria-labelledby="objectives-h">
            <h2 id="objectives-h">Learning objectives</h2>
            <ul className="xl-points">
              {lesson.objectives.map((objective) => (
                <li key={objective}>{objective}</li>
              ))}
            </ul>
          </section>

          {lesson.sections.map((section) => (
            <Section key={section.id} section={section} separator={separator} />
          ))}

          <section className="xl-section xl-workbook" id="workbook" aria-labelledby="workbook-h">
            <h2 id="workbook-h">Download the practice workbook</h2>
            <p>
              The workbook is generated in your browser when you click download — it is not fetched from a server and no upload is involved. It contains the{" "}
              {lesson.workbookSheets.map((sheet) => sheet.name).join(", ")} sheet{lesson.workbookSheets.length === 1 ? "" : "s"} used in this lesson.
            </p>
            <button type="button" className="ws-btn-primary" onClick={download}>
              <Icon name="download" /> Download Lesson {lesson.n} workbook (.xlsx)
            </button>
            <span className="sr-only" role="status" aria-live="polite">
              {workbookState === "done" ? `Lesson ${lesson.n} practice workbook downloaded` : ""}
            </span>
            {workbookState === "done" ? <p className="xl-note">Downloaded — check your downloads folder.</p> : null}
          </section>

          <div className="xl-tool-cta">
            <span className="xl-eyebrow">Related MarqClean AI tool</span>
            <h2>Do this on your own file</h2>
            <p>{lesson.tool.note}</p>
            <a className="ws-btn-primary" href={lesson.tool.path}>
              {lesson.tool.label} <Icon name="arrow" />
            </a>
          </div>

          <Quiz lesson={lesson} onPass={markComplete} />

          <section className="xl-section xl-complete" aria-labelledby="xl-complete-h">
            <h2 id="xl-complete-h">{done ? "Lesson complete" : "Mark this lesson complete"}</h2>
            <p>{done ? "Recorded in this browser. You can review it any time." : "Progress is stored in this browser only."}</p>
            <button type="button" className={done ? "ws-btn-secondary" : "ws-btn-primary"} onClick={() => onToggle(lesson.slug)}>
              {done ? "Mark as not complete" : "Mark complete"}
            </button>
          </section>

          <ServiceCta />

          <nav className="xl-next-prev" aria-label="Lesson navigation">
            {prev ? (
              <a href={`/${ROUTE}/${prev.slug}`} onClick={(event) => { event.preventDefault(); onOpen(prev.slug); }}>
                <span>Previous</span>
                <strong>
                  Lesson {prev.n}: {prev.title}
                </strong>
              </a>
            ) : (
              <span />
            )}
            {next ? (
              <a className="is-next" href={`/${ROUTE}/${next.slug}`} onClick={(event) => { event.preventDefault(); onOpen(next.slug); }}>
                <span>Next</span>
                <strong>
                  Lesson {next.n}: {next.title}
                </strong>
              </a>
            ) : (
              <a className="is-next" href={`/${ROUTE}/${CHEAT_SLUG}`} onClick={(event) => { event.preventDefault(); goTo(CHEAT_SLUG); }}>
                <span>Finish</span>
                <strong>Open the formula cheat sheet</strong>
              </a>
            )}
          </nav>

          <section className="xl-section" id="lesson-faq" aria-labelledby="lesson-faq-h">
            <h2 id="lesson-faq-h">Questions about this lesson</h2>
            {lesson.faq.map((entry) => (
              <details key={entry.q} className="xl-faq">
                <summary>{entry.q}</summary>
                <p>{entry.a}</p>
              </details>
            ))}
          </section>

          <section className="xl-section" aria-labelledby="xl-related-h">
            <h2 id="xl-related-h">Related lessons</h2>
            <ul className="xl-related">
              {XLSX_LESSONS.filter((item) => item.slug !== lesson.slug)
                .slice(0, 4)
                .map((item) => (
                  <li key={item.slug}>
                    <a href={`/${ROUTE}/${item.slug}`} onClick={(event) => { event.preventDefault(); onOpen(item.slug); }}>
                      <span>Lesson {item.n}</span> {item.title}
                    </a>
                  </li>
                ))}
            </ul>
          </section>

          <Disclaimer />
        </article>
      </div>
    </>
  );
}

function CheatSheet({ separator, onSeparator, onOpenLesson }: {
  separator: ";" | ",";
  onSeparator: (value: ";" | ",") => void;
  onOpenLesson: (slug: string) => void;
}) {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState<"All" | CheatCategory>("All");

  const rows = useMemo(() => {
    const needle = query.trim().toLowerCase();
    return CHEAT_ENTRIES.filter((entry) => {
      if (category !== "All" && entry.category !== category) return false;
      if (!needle) return true;
      return `${entry.name} ${entry.purpose} ${entry.category} ${entry.syntax}`.toLowerCase().includes(needle);
    });
  }, [query, category]);

  return (
    <>
      <PageHeader
        workspace="Excel Learning"
        title="Excel Formula Cheat Sheet"
        description="Every function from Excel for Data Analysis, searchable by name, category or the problem you are trying to solve."
        onBack={() => goTo("")}
      />
      <Breadcrumbs
        trail={[
          { name: "Home", href: "/" },
          { name: "Excel Learning", href: `/${ROUTE}` },
          { name: "Formula Cheat Sheet", href: `/${ROUTE}/${CHEAT_SLUG}` },
        ]}
      />

      <div className="xl-cheat-controls">
        <div className="xl-search">
          <label htmlFor="xl-cheat-search">
            <span className="sr-only">Search functions by name, category or use case</span>
            <Icon name="search" />
          </label>
          <input
            id="xl-cheat-search"
            type="search"
            value={query}
            placeholder="Search by function, category or problem — try 'duplicate' or 'not found'"
            onChange={(event) => setQuery(event.target.value)}
          />
          {query ? (
            <button type="button" onClick={() => setQuery("")} aria-label="Clear the function search">
              Clear
            </button>
          ) : null}
        </div>
        <div className="xl-filters" role="group" aria-label="Filter by category">
          {(["All", ...CHEAT_CATEGORIES] as const).map((option) => (
            <button key={option} type="button" aria-pressed={category === option} onClick={() => setCategory(option as "All" | CheatCategory)}>
              {option}
            </button>
          ))}
        </div>
        <SeparatorControl separator={separator} onChange={onSeparator} />
      </div>

      <p className="xl-results-count" role="status">
        {rows.length} {rows.length === 1 ? "function" : "functions"}
        {category !== "All" ? ` in ${category}` : ""}
        {query ? ` matching “${query}”` : ""}
      </p>

      {rows.length === 0 ? (
        <div className="xl-empty">
          <h2>No function matches that search</h2>
          <p>
            This cheat sheet covers the functions taught in the ten lessons. For the full A-Z Excel function reference, including financial, date and dynamic
            array functions, use the <a href="/excel-academy">Excel Academy function reference</a>.
          </p>
          <button type="button" className="ws-btn-secondary" onClick={() => { setQuery(""); setCategory("All"); }}>
            Clear the filters
          </button>
        </div>
      ) : (
        <div className="xl-cheat-list">
          {rows.map((entry) => {
            const lesson = XLSX_LESSONS.find((candidate) =>
              candidate.sections.some((section) => section.kind === "functions" && section.functions.some((fn) => fn.name === entry.name)),
            );
            return (
              <article key={entry.name} className="xl-cheat-entry" id={`fn-${entry.name.toLowerCase()}`}>
                <div className="xl-cheat-entry__head">
                  <h2>{entry.name}</h2>
                  <span className="ws-format-badge">{entry.category}</span>
                </div>
                <p>{entry.purpose}</p>
                <CodeBlock code={entry.syntax} separator={separator} caption="Syntax" />
                {entry.examples[0] ? (
                  <div className="xl-cheat-entry__example">
                    <CodeBlock code={entry.examples[0].formula} separator={separator} caption="Example" />
                    <p className="xl-cheat-entry__result">
                      <strong>Result:</strong> {entry.examples[0].result}
                    </p>
                  </div>
                ) : null}
                <dl className="xl-cheat-entry__meta">
                  <div>
                    <dt>Availability</dt>
                    <dd>{entry.version ?? "All versions of Excel"}</dd>
                  </div>
                  {lesson ? (
                    <div>
                      <dt>Full lesson</dt>
                      <dd>
                        <button type="button" onClick={() => onOpenLesson(lesson.slug)}>
                          Lesson {lesson.n}: {lesson.title.split(":")[0]}
                        </button>
                      </dd>
                    </div>
                  ) : null}
                </dl>
              </article>
            );
          })}
        </div>
      )}

      <ServiceCta />
      <Disclaimer />
    </>
  );
}

function NotFound() {
  return (
    <>
      <PageHeader workspace="Excel Learning" title="That lesson does not exist" description="The page you asked for is not part of this course." onBack={() => goTo("")} />
      <section className="xl-section">
        <h2>All ten lessons</h2>
        <ul className="xl-related">
          {XLSX_LESSONS.map((lesson) => (
            <li key={lesson.slug}>
              <a href={`/${ROUTE}/${lesson.slug}`} onClick={(event) => { event.preventDefault(); goTo(lesson.slug); }}>
                <span>Lesson {lesson.n}</span> {lesson.title}
              </a>
            </li>
          ))}
        </ul>
        <button type="button" className="ws-btn-primary" onClick={() => goTo("")}>
          Back to the course overview
        </button>
      </section>
    </>
  );
}

/* ------------------------------------------------------------------ *
 * Shell
 * ------------------------------------------------------------------ */

function readCompleted(): Set<string> {
  try {
    const raw = window.localStorage.getItem(PROGRESS_KEY);
    return new Set(raw ? (JSON.parse(raw) as string[]) : []);
  } catch {
    return new Set();
  }
}

function readSeparator(): ";" | "," {
  try {
    return window.localStorage.getItem(SEPARATOR_KEY) === "," ? "," : ";";
  } catch {
    return ";";
  }
}

export default function ExcelLearning() {
  const [page, setPage] = useState<SubPage>(() => readSubPage());
  const [completed, setCompleted] = useState<Set<string>>(() => readCompleted());
  const [separator, setSeparator] = useState<";" | ",">(() => readSeparator());

  useEffect(() => {
    const sync = () => setPage(readSubPage());
    window.addEventListener("popstate", sync);
    return () => window.removeEventListener("popstate", sync);
  }, []);

  useEffect(() => {
    applyMetadata(page);
  }, [page]);

  const toggle = (slug: string) => {
    setCompleted((previous) => {
      const next = new Set(previous);
      if (next.has(slug)) next.delete(slug);
      else next.add(slug);
      try {
        window.localStorage.setItem(PROGRESS_KEY, JSON.stringify([...next]));
      } catch {
        /* storage may be blocked; progress simply will not persist */
      }
      return next;
    });
  };

  const changeSeparator = (value: ";" | ",") => {
    setSeparator(value);
    try {
      window.localStorage.setItem(SEPARATOR_KEY, value);
    } catch {
      /* preference simply will not persist */
    }
  };

  const open = (slug: string) => goTo(slug);

  return (
    <div className="ws-root ws-scope min-h-screen">
      <GlobalHeader />
      <main id="main-content" className="xl-page">
        <div className="ws-page-frame">
          {page.kind === "lesson" && lessonBySlug(page.slug) ? (
            <LessonPage
              lesson={lessonBySlug(page.slug)!}
              completed={completed}
              onToggle={toggle}
              onOpen={open}
              separator={separator}
              onSeparator={changeSeparator}
            />
          ) : page.kind === "cheat" ? (
            <CheatSheet separator={separator} onSeparator={changeSeparator} onOpenLesson={open} />
          ) : page.kind === "missing" ? (
            <NotFound />
          ) : (
            <Landing completed={completed} onOpen={open} separator={separator} onSeparator={changeSeparator} />
          )}
        </div>
      </main>
    </div>
  );
}
