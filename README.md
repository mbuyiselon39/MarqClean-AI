# MarqClean AI

> **Clean data. Clear decisions.**

MarqClean AI is a browser-first data operations platform for cleaning, validating, transforming, matching and reconciling business data.

It is designed for finance, operations, compliance, marketing and data teams that need reliable spreadsheet and document workflows without sending working data to a traditional application backend.

## Live application

**[Open MarqClean AI](https://marqcleanai.vertexsg.co.za/)**

## What it does

- **Data cleaning** — standardise and prepare messy datasets for downstream use.
- **Validation** — identify missing, invalid, inconsistent and suspicious values.
- **Excel & CSV processing** — work with common business spreadsheet formats directly in the browser.
- **PDF processing** — extract and prepare structured information from supported PDF workflows.
- **Matching & reconciliation** — compare source and target datasets and surface matches, mismatches and duplicate keys.
- **Duplicates & exceptions** — make data-quality problems visible and actionable.
- **Filtering & review** — inspect records and isolate quality issues before producing an output.
- **Local-first processing** — keep supported data processing in the browser rather than depending on an application server.
- **Export-ready workflows** — prepare cleaned and reconciled results for operational use.

## Architecture

MarqClean AI is built as a static/serverless web application:

- React
- TypeScript
- Vite
- Tailwind CSS
- ExcelJS / SheetJS
- Papa Parse
- PDF.js
- DuckDB-WASM for browser-side analytical processing
- Zod for validation
- Cloudflare Pages for static hosting

No traditional application server is required for the core data-processing workflow.

## Data handling

The platform is designed around browser-first processing. Files are loaded into the application for processing and analysis in the user's browser, with supported transformations performed locally.

**Always follow your organisation's information-security, privacy and data-retention requirements before loading sensitive information into any web application.**

## Quality and reliability

The repository includes automated checks for:

- TypeScript type safety
- Production builds
- Client-side routing
- SPA fallback configuration
- Production hostname consistency
- Sitemap and robots hostname consistency
- Cloudflare Pages asset-size limits
- Reconciliation duplicate-key handling

The main branch is intended to remain deployment-ready.

## Development

### Requirements

- Node.js 24+
- npm

### Install

~~~bash
npm install
~~~

### Run locally

~~~bash
npm run dev
~~~

### Production build

~~~bash
npm run build
~~~

## Project structure

~~~text
.
├── components/        # Application UI and workflow components
├── public/            # Static assets and SPA deployment configuration
├── workers/           # Browser-side processing workers
├── App.tsx            # Application shell and client-side routing
├── engine.ts          # Data quality, matching and reconciliation logic
├── data-model.ts      # Core data structures
├── duckdbEngine.ts    # Browser-side DuckDB integration
├── localDataEngines.ts# Local spreadsheet/data processing
└── vite.config.ts     # Vite build configuration
~~~

## Deployment

The application is configured for **Cloudflare Pages** with an SPA fallback so client-side routes can be served correctly from direct URLs.

Production domain:

**https://marqcleanai.vertexsg.co.za/**

## Status

MarqClean AI is under active development. Features and supported file-processing workflows may evolve as the platform is expanded.

---

**MarqClean AI**  
*Clean data. Clear decisions.*