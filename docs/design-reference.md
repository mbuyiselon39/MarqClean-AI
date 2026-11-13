# MarqClean AI — design system

Updated: 2026-11-13
Status: the single source of truth for colour, type, space, radius, elevation,
motion and focus. Tokens are defined **once**, in the `:root` and
`[data-theme="dark"]` blocks at the top of `index.css`.

`.github/workflows/quality.yml` asserts this file exists, so the CI gate and the
design system cannot drift apart silently.

## Product principles

MarqClean AI is a browser-first data operations platform. The interface
communicates precision and trust without suggesting features that are not
present: core file processing is local, output is reviewable, and the workspaces
are free to use. The homepage showcases an interactive **illustrative**
workspace; its sample metrics are labelled rather than represented as live
customer data.

## Colour

### Palette

The rebrand moves the brand hue from evergreen to **cobalt (hue 232)**. The
reason is functional, not decorative: MarqClean's primary colour sits next to
green "success" and red "error" states in every workspace. An evergreen brand
accent competes with the success semantic and makes data-quality feedback
harder to read at a glance. Cobalt is far from both ends of the semantic axis.

| Token | Light (RGB) | Dark (RGB) | Role |
| --- | --- | --- | --- |
| `--canvas` | `246 246 249` | `10 13 20` | page background |
| `--surface` | `255 255 255` | `19 22 32` | cards, panels, inputs |
| `--surface-2` | `239 241 245` | `27 31 44` | nested / inset surfaces |
| `--surface-3` | `230 233 239` | `36 41 56` | pressed / selected fills |
| `--surface-inverse` | `21 26 40` | `237 238 243` | inverted bands (CI-asserted) |
| `--on-inverse` | `255 255 255` | `21 26 40` | text on inverted bands (CI-asserted) |
| `--ink` | `16 18 25` | `237 238 243` | primary text |
| `--ink-2` | `52 59 81` | `200 205 218` | secondary text |
| `--ink-3` | `76 87 118` | `182 188 206` | tertiary text, labels |
| `--ink-4` | `91 105 143` | `163 171 194` | quietest text — still AA |
| `--accent` | `31 53 193` | `115 133 250` | brand, links, primary actions |
| `--accent-tint` | `237 239 253` | `31 36 71` | accent fills behind accent text |
| `--success` | `21 128 78` | `102 214 162` | data-quality pass |
| `--error` | `217 38 50` | `241 126 134` | data-quality fail |
| `--warning` | `165 101 29` | `238 187 99` | caution |
| `--info` | `37 117 208` | `117 174 240` | neutral notice |
| `--line` | `221 224 233` | `44 50 68` | decorative separators |
| `--line-strong` | `198 202 215` | `63 69 90` | hover borders, dividers |
| `--line-control` | `128 136 158` | `97 106 133` | **control** boundaries |

`--magenta`, `--magenta-hover`, `--soft-pink` and `--soft-pink-2` are retained
as aliases of the cobalt accent so rules written against the old vocabulary keep
resolving. They are not an independent palette.

### Contrast (WCAG 2.2)

Every value below is computed with the WCAG relative-luminance formula from the
token values in `index.css` — not eyeballed.

- **44 / 44 text pairs pass AA (4.5:1)** across both themes. This includes the
  quietest pair, `--ink-4` on `--surface-2`.
- **6 / 6 control-boundary pairs pass 3:1** (`--line-control` against
  `--surface`, `--canvas` and `--surface-2`, both themes).

Three pairs failed the first pass and were corrected by solving for the
threshold rather than adjusting by eye:

| Problem | Fix |
| --- | --- |
| `--line-control` on `--surface-2` was 2.93:1 light / 2.92:1 dark | light `128 136 158` → 3.13:1; dark `97 106 133` → 3.06:1 |
| Dark `--accent` on `--accent-tint` was 4.39:1 | dark `--accent` → `115 133 250` = 4.61:1 |
| Light `--success` on `--canvas` was 4.48:1 | light `--success` → `21 128 78` = 4.60:1 |

**Two border tokens deliberately fail 3:1.** `--line` (1.2–1.4:1) and
`--line-strong` (1.5–2.1:1) are decorative only. WCAG 1.4.11 applies to the
boundary needed to *identify* a component, so every input, select and button
uses `--line-control` at rest, and `--line-strong` appears only on hover states
and dividers. Do not promote either token to a resting control border.

## Typography

Self-hosted variable **Manrope** (display) and **DM Sans** (body/interface).
Both are copied from npm into `public/fonts/` by `scripts/fetch-fonts.mjs`,
which also rewrites the preload block in `index.html`. Combined budget:
**61.8 KB against the 150 KB CI limit** (DM Sans 36,932 B, Manrope 24,836 B).
"Clash Grotesk" and "Satoshi" survive only as fallback names in the
`font-family` stacks, which the CI gate checks for.

Ten type steps, and nothing else. The previous sheet carried **60 distinct
font-size values**, 39 of them below the 11px readability floor (7, 8, 9, 10 px)
and 18 at half-pixel sizes that rasterise inconsistently across browsers.

| Token | Size | Line |
| --- | --- | --- |
| `--text-2xs` | 11px | 16px |
| `--text-xs` | 12px | 18px |
| `--text-sm` | 14px | 21px |
| `--text-base` | 16px | 25px |
| `--text-md` | 18px | 27px |
| `--text-lg` | 20px | 28px |
| `--text-xl` | 24px | 32px |
| `--text-2xl` | 30px | 38px |
| `--text-3xl` | 38px | 44px |
| `--text-4xl` | `clamp(38px, 2.6rem + 2.4vw, 56px)` | 1.05 |
| `--text-5xl` | `clamp(48px, 3rem + 4vw, 72px)` | 1.02 |

`--text-floor: 11px` is the minimum. The CI gate rejects any literal
`font-size` below it, so the floor is enforced rather than aspirational.

Display sizes are fluid. The old sheet set `h1 { font-size: 60px }` with no
responsive step-down, which overflowed narrow viewports; `--text-4xl` scales
instead. Body copy is 16px — the previous 15px default is below the comfortable
reading size for the dense documentation this product publishes.

## Space, radius and breakpoints

- **Space** (`--space-1…--space-10`): `4, 8, 12, 16, 24, 32, 48, 64, 96, 128px`. Prefer the tokens over
  literals so vertical rhythm stays consistent across the four style families.
- **Radius** (`--radius-xs` 6 / `-sm` 10 / `-md` 14 / `-lg` 20 / `-pill` 999):
  the old sheet mixed 4, 6, 8, 10, 12, 14, 16 and 999px across families, so a
  card looked different on each page.
- **Breakpoints** (exposed via `@theme`): 480 / 640 / 768 / 1024 / 1280 / 1536.
  Media queries use the six canonical `max-width` values **479, 639, 767, 899,
  1023, 1279** (each one below a scale step). The sheet previously used
  **18 distinct** breakpoints — 420, 520, 535, 600, 620, 639, 640, 700, 740,
  767, 820, 840, 860, 900, 1023, 1120, 1199, 1200.

## Elevation

Depth is a **closed scale of four**. A shadow may be `none` or one of
`--shadow-1…4`; the CI gate rejects anything else.

| Token | Use |
| --- | --- |
| `--shadow-1` | resting cards, inputs, secondary buttons |
| `--shadow-2` | card hover, primary button hover |
| `--shadow-3` | popovers, dropdowns, tooltips |
| `--shadow-4` | command palette, modals, mobile menu, sidebar |

The gate previously banned every shadow outright. That produced a flat,
undifferentiated UI where a floating command palette and a resting card were
indistinguishable — so depth leaked back in as ad-hoc `border` and `background`
hacks instead. The scale is deliberately small enough that depth stays a
decision, not a habit: if a surface needs a level that is not on this list, the
hierarchy is wrong, not the token set.

## Motion

`--dur-fast` 120ms (colour, borders), `--dur-base` 200ms (hover lift, tabs),
`--dur-slow` 320ms (overlays), with `--ease-out` and `--ease-in-out`. Under
`prefers-reduced-motion: reduce` all durations collapse to 1ms and the hover
`translateY` is removed, so motion preference is honoured globally rather than
per-component.

## Focus

```css
:where(button, input, select, textarea, a, [tabindex], summary, [role="menuitem"])
  :focus-visible {
  outline: var(--ring-width) solid rgb(var(--accent));
  outline-offset: var(--ring-offset);
}
```

Focus is the one affordance that is never softened. It is 3px at ≥3:1 against
both `--surface` and `--accent-tint`, and a separate `forced-colors` rule
guarantees a visible ring in Windows High Contrast mode.

## Touch targets

Interactive controls are held to **44 × 44 CSS px** minimum via
`--control-height`. Dense data grids and table rows are exempt (they are not
primary navigation) but their tap areas are padded to at least 32px.

## Structure: the four style families

`index.css` currently contains four historically stacked class families:
`.mc-` (287 rules), `.ws-` (208), `.xl-` (183) and `.academy-` (55), plus 11
stragglers. This — not the palette — is the main source of visual incoherence.

The "Component language" block at the end of `index.css` neutralises the
disagreement without a risky rewrite: it re-asserts one radius, one elevation,
one border and one transition treatment for every card, panel, overlay, button,
badge, tab and table across all four families, and wins by source order rather
than `!important` (which the CI gate bans). Collapsing the families into a
single namespace is the correct end state and is tracked separately; the
override block is the safe intermediate step that makes the rebrand land
without touching 743 rules of JSX class vocabulary at once.

## Verification

```bash
npm ci
npx tsc --noEmit
npm run build
node scripts/fetch-fonts.mjs   # only if public/fonts/ is missing
```

The build self-hosts its fonts and makes no external font requests.
`npm run build` emits ~152 KB of CSS (~27.6 KB gzipped) covering all 73 routes.
Contrast figures above are reproducible by running the WCAG formula against the
token blocks in `index.css`.
