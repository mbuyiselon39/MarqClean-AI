# MarqClean AI — interface direction

Updated: 2026-09-30

## Product principles

MarqClean AI is a browser-first data operations platform. The interface should communicate precision and trust without suggesting features that are not present: core file processing is local, output is reviewable, and the workspaces remain free to use. The homepage showcases an interactive **illustrative** workspace; its sample metrics are labelled rather than represented as live customer data.

## Visual system

- **Palette:** warm off-white canvas, white working surfaces, deep petrol brand hero, restrained evergreen actions, and soft mint emphasis. The dark theme uses deep navy surfaces and accessible mint actions.
- **Typography:** self-hosted Manrope for display and DM Sans for body/interface text. Variable Latin WOFF2 files are copied from npm dependencies during builds, avoiding remote runtime and build-time font requests. Legacy font-family names remain only as fallbacks.
- **Shape:** 8–16px control/card radii, fine borders, minimal surface lift, no decorative photography or fabricated social proof.
- **Navigation:** compact global header, searchable command palette, direct links to the platform and tools, workspace switcher, responsive mobile menu and clear launch action.
- **Hero:** plain-language value proposition paired with an interactive product preview. Preview tabs, search, file selection, drag-and-drop and tool launches retain their real interactions. Sample rows and metrics are explicitly illustrative.
- **Tool surfaces:** existing functional modules and routes retain their processing, preview and download logic. Design tokens and component framing provide visual consistency without replacing the data engines.

## Responsive and accessible behavior

The hero collapses to one column on tablets and the preview navigation scrolls horizontally on small screens. Workspace and command palette navigation remain keyboard-accessible. Buttons and links retain visible focus states; motion respects `prefers-reduced-motion`. The theme switcher supports light, dark and operating-system preferences.

## Verification

Run `npm ci`, `npx tsc --noEmit`, and `npm run build`. The build self-hosts its font files, so it does not need access to an external font API. Core app route, preview host, and local font paths can be checked through the Vite preview.
