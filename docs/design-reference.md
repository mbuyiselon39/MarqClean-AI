# MarqClean AI — live design reference

Updated: 2026-09-21

## Measurement note

The requested 1440px and 390px Chrome DevTools computed-style capture was not available in the repository automation environment, so this file does **not** invent computed values. The live beehiiv page was inspected for structure/content and cross-checked against a current independent style extraction. Where a computed DevTools value could not be directly captured, the table marks it as **not captured** and records the usable reference value instead.

The live page currently presents an uppercase display-led dark product-launch layout, with the hero, product areas, pricing, success story, integrations, FAQ, CTA and footer sections in that order. https://www.beehiiv.com/

A current Refero extraction reports Clash Grotesk/Satoshi, a 1200px max width, 8px base spacing, 96px section gaps, 32px card padding, 6px rectangular radii, and a 72px display scale. These are reference observations, not a substitute for a DevTools computed-style capture. See the source: https://styles.refero.design/style/350b1557-56f0-4361-8c8b-b7a88081982b 

## Reference table

| Element | 1440 computed | 390 computed | Reference / implementation value | Difference from starting token |
|---|---|---|---|---|
| Hero H1 | Not captured | Not captured | Clash Grotesk 700; 72px desktop reference; 40px mobile implementation; line-height 1 | Live reference indicates a larger 72px display scale than the 60px starting token |
| Hero subtext | Not captured | Not captured | Satoshi 400; 18px; ~1.55 line-height; max 60ch | Aligns with starting body-lg range |
| Primary CTA | Not captured | Not captured | Satoshi 700; 44-48px minimum; indigo fill; pill in hero | Radius differs: hero uses pill as requested |
| Secondary CTA | Not captured | Not captured | Satoshi 500-700; 44px; transparent + hairline | Uses 6px radius |
| Trusted label | Not captured | Not captured | Satoshi 700; 12px; uppercase; 0.08em | Matches requested eyebrow treatment |
| Section H2, uppercase | Not captured | Not captured | Clash Grotesk 700; 48px desktop / 32px mobile; uppercase via CSS | Live reference supports 48px |
| Section H2, sentence-case source | Not captured | Not captured | Source copy remains sentence case; CSS transforms display headings uppercase | Keeps semantic copy readable in source |
| Feature-tab card | Not captured | Not captured | Surface-1, 1px hairline, 6px/12px panel radius, 44-48px control height | Starting token retained |
| Feature sub-item H4 + body | Not captured | Not captured | Satoshi 700 20px / 14-16px body | Matches requested 20-24px sub-heading |
| Pricing / free-access card | Not captured | Not captured | Surface-1, 1px hairline, 12px large panel; no fabricated tiers | Replaced paid-plan imitation with factual free-access block |
| Testimonial / proof card | Not captured | Not captured | No fabricated proof; MarqClean uses product/output evidence instead | No third-party social proof copied |
| Stat number | Not captured | Not captured | Clash Grotesk 700; 48-60px; indigo/magenta only as emphasis | Existing strip reconciled to 63 function actions / 33 tools / 6 core formats / 0 server uploads |
| FAQ question/answer | Not captured | Not captured | Satoshi 600 question; 16px body; 60px minimum row; 1px dividers | Matches 56-64px starting requirement |
| Footer heading/link | Not captured | Not captured | Satoshi 700 12px uppercase heading; 14px links; 5 columns | Matches requested footer structure |
| Header nav link | Not captured | Not captured | Satoshi 600-700; 14px; muted ink → white hover | Uses single nav layer |
| Container width | Not captured | Not captured | 1240px implementation; 24px mobile / 32px desktop gutters | Within requested 1200-1280px range; 1200px is the external live reference |

## Beehiiv live structure used only as design-language reference

- Dark near-black canvas with violet surfaces.
- Clash Grotesk display typography and Satoshi interface/body typography.
- Thin borders and surface lift instead of conventional shadows.
- Large uppercase section headlines.
- Product UI mockups as visual evidence rather than decorative lifestyle imagery.
- Indigo/magenta used as restrained accent punctuation.
- Product sections, pricing, proof, integrations, FAQ, CTA and footer create the page rhythm.

No beehiiv copy, logos, customer names, statistics, testimonials, avatars or images were added to MarqClean AI.

## Font licensing note

Fontshare states that its fonts are free for personal and commercial use, and its license page distinguishes open-source fonts from closed-source fonts governed by the ITF Free Font License. Satoshi is currently listed by Fontshare as a closed-source font. The implementation downloads the unmodified WOFF2 web assets from Fontshare during the build and serves those files from MarqClean AI; no font files are modified or subsetted. Review the current ITF license terms before any future font-file redistribution or packaging change: https://fontshare.com/licenses/sil-ofl and https://www.fontshare.com/?q=Satoshi

## Verification still requiring a real browser

- Chrome DevTools computed styles at exactly 1440px and 390px.
- Dark/light visual screenshots.
- Lighthouse Performance >= 90 and Accessibility >= 95.
- Pixel-level left-edge alignment at 1440 / 820 / 390.
- Final production screenshot capture for home, tool page, contact, Data Toolbox, Reconciliation Hub and Excel Automation.
