# QA before — navigation/workspace audit

## Browser capture status
The requested 320/390/768/1024/1280/1440/1920/2560px screenshots and 125%/150% Windows scaling captures could not be produced in the repository execution environment because there is no browser-rendering/DevTools session available. No synthetic screenshots are stored here.

## Source causes confirmed before the refactor

- The global header differed between the marketing home and internal pages.
- Internal pages rendered a fixed 272px sidebar without reserving an equivalent content column.
- The sidebar contained two customer-visible search/command-palette affordances in addition to the header search.
- The workspace switcher was a static MarqClean AI block rather than a real workspace menu.
- Toolbox and Reconciliation navigation used generic group inference rather than the required workspace-specific information architecture.
- Academy reused the application sidebar and displayed three navigation layers simultaneously.
- Academy lesson/function state was local component state, so lesson/function selections were not represented by deep-linkable browser history.
- Several internal pages used marketing-style display typography and duplicate title/back blocks.
- Internal pages used emoji/Unicode icons in workspace navigation.
- The hero preview was a static mock: its sidebar, drop area and action rows did not execute the corresponding product workflows.
- The hero preview search affordance did not own the same global search modal as the site header.
