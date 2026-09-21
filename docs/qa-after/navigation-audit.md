# QA after — navigation/workspace audit

## Implemented

- One shared global header is used by marketing and internal pages.
- Workspaces dropdown restored with keyboard-accessible menu semantics and mobile accordion.
- Exactly one header search trigger opens the global search palette; Ctrl K / ⌘ K and / also open it.
- Search palette now groups Workspaces, Tools, Lessons and Functions and stores recent selections locally.
- Workspace sidebar uses a 272px desktop column, a 72px collapsed rail, sticky/independent scrolling, pinned local-processing status, and a responsive off-canvas drawer below 1024px.
- Drawer has body scroll lock, Escape handling and focus trapping.
- Toolbox groups: Data preparation / Analysis / Extraction.
- Reconciliation groups: Verification / Compliance / Reporting / Administration.
- Sidebar uses line icons and active aria-current state.
- Shared PageHeader added to internal, legal/contact and SEO tool pages.
- Toolbox and Reconciliation tab changes push hash state on the same existing route, allowing browser Back/Forward and deep links without introducing new routes.
- Academy no longer uses the app sidebar on its overview. It uses Curriculum / Function reference / Cheat sheet page tabs.
- Academy lessons have a course outline, progress, objectives, steps, examples, quiz, completion state and Previous/Next controls.
- Academy progress is persisted with guarded localStorage access.
- Academy lesson/function selections are represented in the existing /excel-academy URL via hash state.
- Hero workspace preview now has keyboard-operable tabs, a real file picker/drop zone wired to the existing cleaning handler, global search integration, and real workflow action buttons.
- Start Free remains an inline pill.
- Workspace/internal typography is compact and sentence case.

## Browser verification still required

- Pixel screenshots at all requested widths and Windows scaling levels.
- Runtime scrollWidth/clientWidth measurement at each requested viewport.
- Lighthouse Accessibility >=95.
- 200% browser zoom visual reflow.
- Final production screenshots for the requested pages in both themes.
