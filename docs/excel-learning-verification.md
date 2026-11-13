# Excel for Data Analysis — verification and corrections record

Updated: 2026-11-13
Applies to: `excel-learning.ts`, `ExcelLearning.tsx`, `xlsx-builder.ts`

This file exists so that every number, formula result and version claim published in the
"Excel for Data Analysis" learning hub can be traced to the check that supports it. It is
the evidence behind the claim on the landing page that the formulas are verified.

---

## 1. Source material status — read this first

The brief for this section referred to **eleven attached images of handwritten notes** (a
cover page plus pages 1 to 10). **No images were present in the repository or the
conversation workspace** at the time of build. This was confirmed by searching the whole
workspace for image files: the only images present are the repository's own
`public/icon-512.png` and `public/og-card.png`.

Consequence:

- The ten lesson topics, their order and their sub-topic lists follow the outline given in
  the written brief exactly. That part is faithful to the request.
- The **workings, examples, wording and any non-standard content in the actual notes have
  not been transcribed**, because the notes were never available to read.
- One specific correction named in the brief (the missing-values mean) *has* been applied,
  because the brief itself stated both the wrong and the right value. See section 4.

**Before publishing, re-check each lesson against your notes.** Anything in your notes that
is not in the published lesson — a different dataset, an extra example, a different
ordering, a South African-specific illustration — needs to be added by hand or the images
supplied so it can be transcribed.

---

## 2. How the numbers were verified

Every calculated result published in the lessons is derived from one shared practice
dataset, defined once in `excel-learning.ts`:

| Sheet | Columns | Rows |
| --- | --- | --- |
| `Sales` | Region, Product, Rep, Units, Unit Price, Sales | 10 data rows, A1:F11 |
| `Payments` | Client, Amount (two amounts deliberately blank) | 4 data rows, A1:B5 |
| `Prices` | Product, Code, Unit Price | 3 data rows, A1:C4 |
| `Dirty` | Client Name, Email, Amount (deliberately messy) | 5 data rows, A1:C5 |
| `Monthly` | Month, Revenue | 6 data rows, A1:B7 |

Verification was done two independent ways.

### 2.1 Cross-check against the repository's own engine

The published results were recomputed using the production functions this site already
ships in `excel-functions.ts` — `sumifs`, `countifs`, `xlookup`, `indexMatch`,
`uniqueValues` — run over the same practice data. All 30 assertions passed. This means the
figures a reader sees in a lesson agree with the figures the site's own tools produce on
the same input.

Representative results:

| Claim in lesson | Published value | Engine result |
| --- | --- | --- |
| `=SUMIF(A2:A11;"North";F2:F11)` (L4) | 8100 | 8,100 |
| `=SUMIF(A2:A11;"South";F2:F11)` (L4) | 10200 | 10,200 |
| `=SUMIF(A2:A11;"East";F2:F11)` (L4) | 4250 | 4,250 |
| `=SUMIF(A2:A11;"West";F2:F11)` (L4) | 6200 | 6,200 |
| `=COUNTIF(A2:A11;"North")` (L4) | 3 | 3 |
| `=COUNTIF(D2:D11;">10")` (L4) | 3 | 3 |
| `=SUMIF(B2:B11;"<>Widget";F2:F11)` (L4) | 14000 | 14,000 |
| `=SUMIFS(F2:F11;A2:A11;"North";B2:B11;"Widget")` (L4) | 2500 | 2,500 |
| `=SUMIFS(F2:F11;F2:F11;">2000";D2:D11;">5")` (L4) | 26000 | 26,000 |
| `=COUNTIFS(A2:A11;"South";B2:B11;"Gadget")` (L4) | 2 | 2 |
| `=XLOOKUP("Sensor";A2:A4;C2:C4)` (L5) | 150 | 150 |
| `=XLOOKUP("G-200";B2:B4;A2:A4)` (L5) | Gadget | Gadget |
| `=INDEX(C2:C4;MATCH("Gadget";A2:A4;0))` (L5) | 400 | 400 |
| `=AVERAGE(B2:B5)` on Payments (L2, L10) | 21500 | 21500 |
| `=COUNT(B2:B5)` / `=COUNTBLANK(B2:B5)` (L2) | 2 / 2 | 2 / 2 |
| Distinct clients after TRIM+UPPER (L6) | 3 | 3 |

Other figures verified arithmetically: `SUM(F2:F11)` = 28750, `AVERAGE` = 2875,
`MIN` = 750, `MAX` = 5000, `SUM(D2:D11)` = 94, `AVERAGEIF` North = 2700,
`AVERAGEIFS` Widget = 2950, PivotTable grand total = 28750 with the 4 × 2 grid
East 4250 / North 8100 / South 10200 / West 6200, Gadget 14000 / Widget 14750,
monthly total = 145000 with average 24166.67 and June-over-January +12500 (+67.6%),
histogram bins at width 5 = 2 / 5 / 2 / 1 / 0.

### 2.2 Practice workbooks round-tripped through a real spreadsheet parser

`xlsx-builder.ts` generates the `.xlsx` downloads in the browser. Each lesson's workbook
was built, then read back with SheetJS (`xlsx`, already a project dependency) and every
cell compared against the source array. Ten workbooks, every cell, all matched. This
confirms the downloads open as valid workbooks with the right sheet names and data.

---

## 3. Errors found and fixed during the build

These were caught by the checks above, not by the notes. They are recorded because they
are the kind of error that would otherwise ship.

| Location | Wrong value written first | Correct value | Cause |
| --- | --- | --- | --- |
| L4 SUMIFS example, sales > 2000 and units > 5 | 20700 | **26000** | Arithmetic slip |
| L4 COUNTIFS example, Widget orders over 10 units | 2 | **3** | Missed one matching row (units 14, 20, 12) |
| L6 Remove Duplicates on the raw name column | "3 duplicates removed, 5 rows become 3" | **1 duplicate removed, 5 rows become 4** | Excel's Remove Duplicates is case-insensitive but does **not** trim spaces, so the space-padded variant survives. The corrected version is the better teaching point and is what the lesson now says. |
| L6 before/after table, "distinct names Excel can tell apart" | 5 → 3 | **4 → 3** | Same cause as above; five literal strings collapse to four under case-insensitive comparison |
| L7 practice answer, PivotTable filtered to Rep = Lerato | "South 7800, total 7800" | **South 10200, total 10200** | All three South rows belong to Lerato, so the filter removes nothing from South |
| L7 result table, East × Gadget cell | "0" | **(blank)** | A PivotTable shows an empty cell by default, not zero. The lesson now explains the difference and where to change it. |
| L8 histogram bin table, the 15–20 row | "20 falls in the bin below" | **20** | Self-contradictory caption; bins are upper-inclusive so 20 belongs in 15–20 |
| L8 practice answer | "crossing the 25000 target in March" | **below target through April, clears it in May** | March is 24800, which is below 25000 |
| L6 title tag | 61 characters | **55 characters** | Exceeded the 60-character limit set in the brief |
| Landing title tag | 63 characters | **55 characters** | Same |
| Cheat-sheet title tag | 70 characters | **54 characters** | Same |
| All lesson title tags | `metaTitle` + `" | MarqClean AI"` (up to 75 chars) | **`metaTitle` alone** | Appending a brand suffix pushed every lesson over the 60-character limit |

---

## 4. The missing-values correction named in the brief

The brief asked for this to be corrected and reported explicitly.

**On the Payments sheet the known amounts are 25,000 and 18,000. Their mean is 21,500, not
20,833.**

- `(25000 + 18000) / 2 = 21500`
- `20833` is the mean of **three** values. `(25000 + 18000 + 19500) / 3 = 20833.33`

So 20,833 is consistent with an example that also contained a third value of 19,500 — most
likely a longer version of the same illustration. That origin is an inference, not something
confirmed against your notes, which were not available.

Where the correction appears:

- **Lesson 2** — the AVERAGE worked example returns 21500 and the practice answer works
  through what happens if a blank is replaced with zero (the average falls to 14333.33).
- **Lesson 6** — missing-values handling, with COUNTBLANK returning 2.
- **Lesson 10, Q4** — the model answer states 21500 and explicitly explains why it is not
  20833, since that distinction is itself a good interview answer.

---

## 5. Excel behaviour and version claims

These describe how Excel behaves rather than arithmetic, so they were checked against
published Microsoft documentation and current references rather than computed.

| Claim | Where used | Basis |
| --- | --- | --- |
| XLOOKUP is in Microsoft 365, Excel 2021 and Excel 2024; **not** Excel 2019 or 2016 | L5, L10, cheat sheet | Microsoft's XLOOKUP support page states it is unavailable in Excel 2016 and 2019 |
| IFS is in Excel 2019, 2021, 2024 and Microsoft 365; not Excel 2016 or earlier | L3 | Introduced for Office 365 subscribers in 2016, shipped in the perpetual release with Excel 2019 |
| IFS returns `#N/A` when no condition is TRUE and there is no catch-all; up to 127 pairs | L3 | Documented IFS behaviour |
| Histogram is a built-in chart type from Excel 2016; earlier versions need the Analysis ToolPak or FREQUENCY | L8 | New statistical chart types (Histogram, Box & Whisker, Pareto, Waterfall, Treemap, Sunburst, Funnel) were introduced with Excel 2016 |
| Power Query is built into Excel 2016+ as Get & Transform; free add-in for Excel 2010 (Professional Plus) and 2013 | L9 | Documented availability; the add-in is no longer receiving feature updates |
| Flash Fill is Excel 2013 and later | L6 | Widely documented; **not independently re-verified during this build** |
| SUMIFS / COUNTIFS / AVERAGEIFS are Excel 2007 and later | L4 | Widely documented; **not independently re-verified during this build** |
| Remove Duplicates is case-insensitive but does not trim spaces | L6 | Behaviour relied on by the corrected worked example above |
| AVERAGE ignores blanks; a zero is counted | L2, L6, L10 | Documented AVERAGE behaviour |
| COUNTBLANK and COUNTA both count a cell whose formula returns an empty string | L2 | Documented behaviour for both functions |
| Excel histogram bins are lower-exclusive, upper-inclusive | L8 | Documented bin behaviour; the worked example depends on it |

Anything marked **not independently re-verified** should be spot-checked in your own Excel
build before publishing. Everything else in this table has a cited basis above.

**Locale note.** All formulas are authored with a semicolon argument separator, matching
the existing `functions.ts` reference on this site and the default list separator in many
South African Windows regional settings. Every lesson and the cheat sheet carry a control
that re-renders and copies every formula with commas instead. This is a deliberate
accessibility choice for a mixed ZA/international audience, not an inconsistency.

---

## 6. What was checked automatically, and how to re-run it

The checks used during this build were run as throwaway scripts against the real modules
(they were removed afterwards to keep them out of the repository). They covered:

1. `npx tsc -p tsconfig.json --noEmit` — clean.
2. `npm run build` — clean; the learning hub is a lazily loaded chunk of about 52 kB gzipped
   and does not affect the payload of any other page.
3. The smoke checks in `.github/workflows/quality.yml` — all pass.
4. Workbook generation round-tripped through SheetJS for all ten lessons — every cell matched.
5. Published results recomputed with `excel-functions.ts` over the practice data — 30/30.
6. Server-side render of the landing page, all ten lessons, the cheat sheet and a
   deliberately invalid lesson path — all rendered, with content assertions on headings,
   progress counters, tables, copy buttons, quizzes, workbooks and CTAs.
7. JSON-LD emitted for every page parsed as valid JSON, with the expected `@type`s, ten
   `syllabusSections` on the Course, correct canonical URLs, and stale schema from other
   routes cleared.
8. Title tags under 60 characters and meta descriptions under 155 characters for all
   twelve indexable pages.

If you want these to run in CI permanently, they are straightforward to reinstate as a
`scripts/verify-excel-learning.mjs` entry plus a step in `quality.yml`.

---

## 7. Open items before publishing

- **Supply the eleven note images** so the lessons can be checked against, and reconciled
  with, your actual material. Until then the content is built from the written outline.
- Confirm the Flash Fill and `*IFS` version statements in your target Excel versions.
- Confirm the service CTA destinations (`/contact`, `/dashboarding`, `/data-automation`)
  are the ones you want lessons to send people to.
- Decide on the relationship between this hub and the existing `/excel-academy`. They
  overlap on VLOOKUP, SUMIF and PivotTables and will compete for the same queries. The
  landing page currently differentiates them explicitly, but a longer-term decision to
  merge or to make one canonical for the other would be better for search.
