import { AnimatePresence, motion } from "framer-motion";
import { strToU8, zipSync } from "fflate";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { parsePhoneNumberFromString } from "libphonenumber-js";
import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";

const ReconciliationHub = lazy(() => import("./reconciliation/ReconciliationHub"));
const DataToolbox = lazy(() => import("./reconciliation/DataToolbox"));
const ClientFunds = lazy(() => import("./reconciliation/ClientFunds"));
const AcademyHub = lazy(() => import("./academy/AcademyHub"));
const HeroCarousel = lazy(() => import("./HeroCarousel"));
import SmartDropzone, { type SmartPipelineAction, type SmartRecommendation } from "./SmartDropzone";
import Logo from "./Logo";
import ThemeToggle from "./ThemeToggle";
import { extractPdfTextInWorker } from "./pdfWorkerClient";

type RawRow = Record<string, unknown>;
type CleanRow = Record<string, string>;

type IndustryMatch = {
  industry: string;
  confidence: "High" | "Medium" | "Needs review";
};

type CleanStats = {
  totalRows: number;
  cleanedRows: number;
  headerFixes: number;
  fieldRepairs: number;
  industriesAdded: number;
  missingFlags: number;
  duplicates: number;
};

type CleanResult = {
  fileName: string;
  rows: CleanRow[];
  headers: string[];
  stats: CleanStats;
  originalHeaders: string[];
};

type FooterPageKey = "about" | "privacy" | "terms" | "cookies" | "contact" | "accessibility";
type ToolPageKey =
  | "clean-csv-file-online"
  | "csv-to-excel-cleaner"
  | "excel-data-cleaner-for-marketers"
  | "lead-list-cleaner"
  | "remove-duplicates-from-csv"
  | "fix-csv-capitalization"
  | "categorize-company-industries-csv"
  | "crm-data-cleanup-tool"
  | "email-list-data-cleaner"
  | "microsoft-excel-lead-cleaner"
  | "csv-to-json-converter"
  | "csv-to-xml-converter"
  | "ai-csv-editor-assistant"
  | "csv-statistics-and-charts"
  | "gpt-for-spreadsheets"
  | "ai-document-generator"
  | "ai-business-analyst"
  | "bank-statement-converter"
  | "pdf-to-excel-converter"
  | "pdf-bank-statement-to-csv"
  | "data-reconciliation-tool"
  | "excel-file-comparison"
  | "document-validation-tool"
  | "pdf-vs-excel-verification"
  | "mailing-list-verification"
  | "merge-spreadsheets"
  | "remove-duplicates-excel"
  | "fuzzy-name-matching"
  | "web-scraping-to-excel"
  | "timesheet-tool"
  | "advanced-excel-functions"
  | "xlookup-online"
  | "sumifs-countifs-tool";
type AppPageKey = "home" | "not-found" | "reconciliation-hub" | "data-toolbox" | "excel-automation" | "excel-academy" | FooterPageKey | ToolPageKey;

type FooterPage = {
  title: string;
  description: string;
  sections: Array<{
    heading: string;
    body: string[];
  }>;
};

type SeoToolPage = {
  slug: ToolPageKey;
  category: string;
  cardTitle: string;
  icon: string;
  title: string;
  h1: string;
  description: string;
  keyword: string;
  h2: string;
  sections: Array<{
    h3: string;
    body: string;
  }>;
};

const COMPANY_NAME = "Vertex Stream Technologies";
const GROUP_NAME = "Vertex Stream Group";
const PRODUCT_NAME = "MarqClean AI";

const WORKSPACE_MODULES: Record<"leads" | "converter" | "formulas" | "bank", { icon: string; title: string; description: string; formats: string[]; engine: string }> = {
  leads: {
    icon: "🧹",
    title: "Quick Data & CSV Cleaner",
    description: "Clean, standardize, validate and transform Excel, CSV and spreadsheet data. Remove duplicates, fix formatting, normalize names and validate records.",
    formats: ["XLSX", "XLS", "CSV"],
    engine: "MarqClean AI Data Quality Engine",
  },
  converter: {
    icon: "🔄",
    title: "CSV to Excel Converter",
    description: "Split delimited text into columns using the same delimiter and text-qualifier logic as the Excel Text to Columns wizard, then export a clean Excel workbook.",
    formats: ["CSV", "TXT", "XLSX"],
    engine: "MarqClean AI Structure Engine",
  },
  formulas: {
    icon: "📊",
    title: "Excel Formulas and Formatting",
    description: "Detect numeric, currency, percentage and date columns, then generate a formatted Excel workbook with totals, number formatting and real formulas.",
    formats: ["CSV", "XLSX"],
    engine: "MarqClean AI Formula Engine",
  },
  bank: {
    icon: "🏦",
    title: "Bank Ledger X",
    description: "Convert PDF bank statements into clean Excel, CSV and QIF files with normalized dates and amounts, ready for ledger reconciliation.",
    formats: ["PDF", "XLSX", "CSV", "QIF"],
    engine: "MarqClean AI Reconciliation Engine",
  },
};
const SITE_URL = "https://www.marqcleanai.vertexsg.co.za";
const BASE_META_DESCRIPTION = `${PRODUCT_NAME} is an AI-powered data automation platform for cleaning, validating, transforming and reconciling Excel, CSV, PDF and financial datasets. Built for operations, compliance, finance and data teams.`;
const MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024;
const MAX_FILE_SIZE_LABEL = "100 MB";
const SUPPORTED_EXTENSIONS = new Set(["csv", "xlsx", "xls"]);
const CONVERTER_SUPPORTED_EXTENSIONS = new Set(["csv", "txt", "xlsx"]);

type DelimiterOption = "tab" | "semicolon" | "comma" | "space" | "other";
type TextQualifier = "\"" | "'" | "";
type ColumnType = "numeric" | "currency" | "percent" | "date" | "text";
type AggregateType = "sum" | "average" | "count" | "min" | "max";
type AggregateChoice = AggregateType | "none";

type WizardSettings = {
  delimiters: Record<DelimiterOption, boolean>;
  otherChar: string;
  treatConsecutiveAsOne: boolean;
  textQualifier: TextQualifier;
};

const DELIMITER_CHAR_MAP: Record<Exclude<DelimiterOption, "other">, string> = {
  tab: "\t",
  semicolon: ";",
  comma: ",",
  space: " ",
};

const DEFAULT_WIZARD_SETTINGS: WizardSettings = {
  delimiters: { tab: false, semicolon: false, comma: true, space: false, other: false },
  otherChar: "",
  treatConsecutiveAsOne: false,
  textQualifier: "\"",
};

const AGGREGATE_LABELS: Record<AggregateChoice, string> = {
  none: "No calculation",
  sum: "Sum",
  average: "Average",
  count: "Count",
  min: "Minimum",
  max: "Maximum",
};

const STYLE_DEFAULT = 0;
const STYLE_HEADER = 1;
const STYLE_BAND = 2;
const STYLE_NUMBER = 3;
const STYLE_CURRENCY = 4;
const STYLE_PERCENT = 5;
const STYLE_DATE = 6;
const STYLE_TOTAL_LABEL = 7;
const STYLE_TOTAL_NUMBER = 8;
const STYLE_TOTAL_CURRENCY = 9;

const PRIORITY_HEADERS = [
  "First Name",
  "Last Name",
  "Full Name",
  "Email",
  "Phone",
  "Company",
  "Job Title",
  "Industry",
  "Industry Confidence",
  "Website",
  "City",
  "State",
  "Country",
  "Data Quality Notes",
];

const HEADER_ALIASES: Record<string, string> = {
  firstname: "First Name",
  first: "First Name",
  fname: "First Name",
  lastname: "Last Name",
  last: "Last Name",
  lname: "Last Name",
  fullname: "Full Name",
  name: "Full Name",
  contactname: "Full Name",
  leadname: "Full Name",
  lead: "Full Name",
  prospectname: "Full Name",
  email: "Email",
  emailaddress: "Email",
  "e-mail": "Email",
  mail: "Email",
  phone: "Phone",
  phonenumber: "Phone",
  mobile: "Phone",
  telephone: "Phone",
  tel: "Phone",
  company: "Company",
  companyname: "Company",
  account: "Company",
  accountname: "Company",
  organization: "Company",
  organisation: "Company",
  business: "Company",
  title: "Job Title",
  jobtitle: "Job Title",
  role: "Job Title",
  position: "Job Title",
  industry: "Industry",
  sector: "Industry",
  vertical: "Industry",
  category: "Industry",
  website: "Website",
  web: "Website",
  url: "Website",
  domain: "Website",
  city: "City",
  town: "City",
  state: "State",
  province: "State",
  region: "State",
  country: "Country",
};

const ACRONYMS = new Set([
  "AI",
  "API",
  "B2B",
  "B2C",
  "CEO",
  "CFO",
  "CMO",
  "COO",
  "CRM",
  "CSV",
  "HR",
  "IT",
  "LLC",
  "PPC",
  "SaaS",
  "SEO",
  "UK",
  "USA",
  "UX",
]);

const INDUSTRY_RULES = [
  {
    industry: "Software and SaaS",
    terms: ["software", "saas", "cloud", "crm", "platform", "ai", "automation", "data", "app", "tech"],
  },
  {
    industry: "Marketing and Advertising",
    terms: ["marketing", "advertising", "media", "seo", "ppc", "creative", "brand", "agency", "growth"],
  },
  {
    industry: "Financial Services",
    terms: ["bank", "finance", "fintech", "insurance", "wealth", "capital", "loan", "payments", "credit"],
  },
  {
    industry: "Healthcare",
    terms: ["health", "medical", "clinic", "pharma", "hospital", "dental", "wellness", "care"],
  },
  {
    industry: "Retail and Ecommerce",
    terms: ["retail", "shop", "store", "commerce", "ecommerce", "fashion", "consumer", "marketplace"],
  },
  {
    industry: "Education",
    terms: ["school", "university", "college", "education", "learning", "academy", "training"],
  },
  {
    industry: "Real Estate",
    terms: ["real estate", "property", "broker", "mortgage", "homes", "leasing", "realtor"],
  },
  {
    industry: "Manufacturing",
    terms: ["manufacturing", "industrial", "factory", "machinery", "engineering", "materials"],
  },
  {
    industry: "Professional Services",
    terms: ["consulting", "legal", "law", "accounting", "advisory", "services", "partner"],
  },
  {
    industry: "Hospitality and Travel",
    terms: ["hotel", "travel", "restaurant", "tourism", "hospitality", "resort", "food"],
  },
  {
    industry: "Logistics and Transportation",
    terms: ["logistics", "transport", "freight", "shipping", "fleet", "delivery", "supply chain"],
  },
  {
    industry: "Nonprofit",
    terms: ["nonprofit", "charity", "foundation", "association", "ngo", "community"],
  },
];

const SAMPLE_ROWS: RawRow[] = [
  {
    "contact name": "  aLEx   JOHNSON ",
    "E-mail": "ALEX.JOHNSON@GROWTHLOOP.COM ",
    Organization: "growthloop saas llc",
    Role: "vp demand generation",
    URL: "growthloop.com",
    City: "new york",
  },
  {
    firstname: "maria",
    lastname: "de la cruz",
    mail: " maria@northstarclinic.org",
    phone: "(415) 555-0198",
    companyname: "northstar health clinic",
    title: "chief marketing officer",
    industry: "",
  },
  {
    Name: "Jordan Smith",
    Email: "jordan@marketnest.agency",
    Company: "marketnest creative agency",
    Phone: "555.710.4420",
    Website: "https://marketnest.agency",
    State: "ca",
  },
  {
    Lead: "Priya Kapoor",
    BadColumn: "priya@homewisebroker.com",
    account: "homewise property brokers",
    jobtitle: "paid media manager",
    phone_number: "+1 212 555 0187",
  },
];

const FOOTER_PAGES: Record<FooterPageKey, FooterPage> = {
  about: {
    title: `About ${PRODUCT_NAME}`,
    description: `${PRODUCT_NAME} is a browser-based Excel, CSV, data-cleaning, automation and reconciliation engine, built by ${COMPANY_NAME}, a division of ${GROUP_NAME}.`,
    sections: [
      {
        heading: "Our Mission",
        body: [
          `${PRODUCT_NAME} helps marketing, operations, finance, compliance and data teams turn messy spreadsheets, CSV exports and PDF statements into clean, structured, decision-ready data without slow manual work.`,
          "The platform is designed for global teams that need reliable data cleaning, Excel automation, and reconciliation, including CRM-ready lead lists, formatted workbooks, and matched, exception-checked ledgers.",
        ],
      },
      {
        heading: "Product Ownership",
        body: [
          `${PRODUCT_NAME} is a product of ${COMPANY_NAME}, a division of ${GROUP_NAME}.`,
          "Vertex Stream Technologies builds practical data workflow tools that support marketers, operators, founders, agencies, finance teams, and growth teams.",
        ],
      },
      {
        heading: "Production Standards",
        body: [
          "Every module runs in the browser, validates supported file formats, flags data quality issues, and exports clean CSV, Excel, or reconciliation reports ready for Microsoft 365, CRM systems, accounting platforms, and marketing automation tools.",
          "Detection and matching logic is rule- and keyword-based rather than an opaque model, and every tool shows its output before you download, so results are reviewable rather than something you have to take on faith.",
        ],
      },
    ],
  },
  privacy: {
    title: "Privacy Policy",
    description: `This Privacy Policy explains how ${PRODUCT_NAME} handles uploaded spreadsheets, statements and datasets, browser processing, contact requests, and basic website information.`,
    sections: [
      {
        heading: "Data Processing",
        body: [
          "Uploaded CSV, Excel and PDF files are processed locally in your browser. The current version does not require files to be sent to a server for cleaning, conversion, or reconciliation.",
          "You are responsible for making sure you have the right to process any personal or financial data contained in lead lists, CRM exports, bank statements, ledgers, or other datasets you upload.",
        ],
      },
      {
        heading: "Information We May Receive",
        body: [
          "If you contact us, we may receive your name, email address, company name, and message so that we can respond to your request.",
          "This website uses Google AdSense to display advertising. Google and its advertising partners may use cookies and similar technologies to serve ads, measure ad performance, and limit repeat ads, based on your visits to this and other websites. Website hosting, analytics, and security providers may also process limited technical information such as device type, browser type, general location, page activity, and IP-derived signals.",
          "You can review how Google uses this information at https://policies.google.com/technologies/partner-sites, and manage ad personalization at https://adssettings.google.com.",
        ],
      },
      {
        heading: "Retention and Security",
        body: [
          "Contact information is retained only for a reasonable business purpose, such as support, compliance, and product communication.",
          "We use privacy-conscious product design, local browser processing, and limited data collection principles to reduce unnecessary exposure of sensitive datasets, including financial and reconciliation data.",
        ],
      },
      {
        heading: "Your Choices",
        body: [
          `You can avoid submitting personal information by using the browser-based tools in ${PRODUCT_NAME} without contacting us.`,
          "For privacy questions or deletion requests, contact support@vertexstreamtechnologies.com.",
        ],
      },
    ],
  },
  terms: {
    title: "Terms of Service",
    description: `These Terms of Service govern use of ${PRODUCT_NAME}, including CSV cleaning, Excel automation, reconciliation, PDF conversion, downloads, and website content.`,
    sections: [
      {
        heading: "Acceptable Use",
        body: [
          `You may use ${PRODUCT_NAME} to clean, convert, reconcile, and validate lawful CSV, Excel, and PDF files for marketing, operations, finance, compliance, data preparation, CRM imports, segmentation, and reporting workflows.`,
          "You must not use the platform to process data you do not have permission to use, to violate privacy or financial regulations, or to support spam, fraud, abuse, or unlawful activity.",
        ],
      },
      {
        heading: "No Professional Advice",
        body: [
          `${PRODUCT_NAME} provides automated data formatting, validation, and reconciliation assistance. It does not provide legal, accounting, audit, compliance, deliverability, or professional advice.`,
          "You should review cleaned, converted, or reconciled files before importing them into CRMs, accounting systems, email tools, advertising platforms, or other business systems.",
        ],
      },
      {
        heading: "Availability and Changes",
        body: [
          "We may improve, modify, suspend, or discontinue features as the product evolves for the global market.",
          "The service is provided on an as-is and as-available basis to the maximum extent permitted by law.",
        ],
      },
      {
        heading: "Disclaimer of Warranties",
        body: [
          `${PRODUCT_NAME} is provided "as is" and "as available," without warranties of any kind, whether express or implied, including implied warranties of merchantability, fitness for a particular purpose, or non-infringement, to the maximum extent permitted by law.`,
          "We do not warrant that the platform will be error-free, uninterrupted, or that every cleaning, conversion, matching, or reconciliation result will be complete or accurate for your specific use case. You are responsible for reviewing output before relying on it.",
        ],
      },
      {
        heading: "Limitation of Liability",
        body: [
          `To the maximum extent permitted by law, ${COMPANY_NAME} and ${GROUP_NAME} are not liable for any indirect, incidental, special, consequential, or punitive damages, or any loss of data, revenue, or business opportunity, arising from your use of or inability to use ${PRODUCT_NAME}, even if advised of the possibility of such damages.`,
          "Nothing in these Terms limits any liability that cannot be limited under applicable law.",
        ],
      },
      {
        heading: "Ownership",
        body: [
          `${PRODUCT_NAME} is operated by ${COMPANY_NAME}, a division of ${GROUP_NAME}. Product names, interface content, and platform assets belong to their respective owners.`,
        ],
      },
    ],
  },
  cookies: {
    title: "Cookie Policy",
    description: `This Cookie Policy explains how cookies and similar technologies may be used on ${PRODUCT_NAME}.`,
    sections: [
      {
        heading: "Cookie Use",
        body: [
          "The core cleaning, conversion, automation, and reconciliation experience does not require cookies to process uploaded files in your browser.",
          "Cookies or similar technologies may be used by hosting, analytics, security, or advertising providers to operate, measure, protect, or monetize the website.",
        ],
      },
      {
        heading: "Advertising and Measurement",
        body: [
          "This website uses Google AdSense. Google may use cookies to serve ads based on your prior visits to this or other websites, measure ad performance, and prevent the same ad from showing repeatedly, where permitted by law and your browser settings.",
          "You can manage ad personalization at https://adssettings.google.com, and manage cookies generally through your browser settings, device privacy controls, and any consent tools presented in your region.",
        ],
      },
      {
        heading: "Essential Technologies",
        body: [
          "Some technologies may be necessary for security, performance, session integrity, and basic site operation.",
        ],
      },
    ],
  },
  contact: {
    title: "Contact",
    description: `Contact Vertex Stream Technologies for ${PRODUCT_NAME} support, product feedback, partnerships, and business inquiries.`,
    sections: [
      {
        heading: "Support",
        body: [
          "For support, privacy requests, accessibility feedback, or business inquiries, email support@vertexstreamtechnologies.com.",
          "When reporting a file issue, include the file type, approximate row count, browser name, and a description of the workflow. Do not send sensitive lead lists, statements, or ledgers unless a secure support process has been agreed.",
        ],
      },
      {
        heading: "Company",
        body: [
          `${PRODUCT_NAME} is a product of ${COMPANY_NAME}, a division of ${GROUP_NAME}.`,
          "Business inquiries can include partnerships, enterprise licensing, agency workflows, Microsoft 365 spreadsheet automation, reconciliation projects, and marketing or finance data operations.",
        ],
      },
    ],
  },
  accessibility: {
    title: "Accessibility Statement",
    description: "MarqClean AI is designed to be accessible, responsive, keyboard-friendly, and usable across modern devices.",
    sections: [
      {
        heading: "Accessibility Commitment",
        body: [
          "We aim to provide a clear, keyboard-friendly, readable, and responsive experience for marketing, operations, finance, compliance, and data teams using desktop, tablet, and mobile browsers.",
          "The interface uses semantic headings, visible focus behavior, descriptive actions, responsive layouts, and text contrast decisions intended to support broad usability.",
        ],
      },
      {
        heading: "Feedback",
        body: [
          "If you experience an accessibility barrier, contact support@vertexstreamtechnologies.com with the page, browser, assistive technology if applicable, and a description of the issue.",
          "We review feedback as part of product quality and continuous improvement for global users.",
        ],
      },
    ],
  },
};

const SEO_TOOL_PAGES: Record<ToolPageKey, SeoToolPage> = {
  "clean-csv-file-online": {
    slug: "clean-csv-file-online",
    category: "Cleaners",
    cardTitle: "CSV File Cleaner",
    icon: "🧹",
    title: `Free Clean CSV File Online | ${PRODUCT_NAME}`,
    h1: "Free Clean CSV File Online for Marketing Lead Lists",
    description:
      "Clean CSV files online for free with MarqClean AI. Fix messy lead lists, standardize spreadsheet columns, repair capitalization, and export clean CSV or Excel files.",
    keyword: "clean CSV file online",
    h2: "A free CSV cleaner for fast, accurate marketing data cleanup",
    sections: [
      {
        h3: "Upload a messy CSV file",
        body: "Drop in a lead list from a CRM, ad platform, event export, or spreadsheet vendor. The same MarqClean AI cleaning engine validates rows, normalizes headers, and prepares a clean preview before export.",
      },
      {
        h3: "Fix common CSV formatting problems",
        body: "The free CSV cleaner repairs name casing, company formatting, email placement, phone formatting, website fields, missing data notes, and campaign-ready industry categories.",
      },
      {
        h3: "Download clean CSV or Excel",
        body: "Export your cleaned spreadsheet as a normal Excel workbook or CSV file for CRM import, segmentation, email outreach, Microsoft 365 reporting, and marketing automation.",
      },
    ],
  },
  "csv-to-excel-cleaner": {
    slug: "csv-to-excel-cleaner",
    category: "Converters",
    cardTitle: "CSV to Excel Cleaner",
    icon: "🔄",
    title: `Free CSV to Excel Cleaner | ${PRODUCT_NAME}`,
    h1: "Free CSV to Excel Cleaner for Marketers",
    description:
      "Convert CSV lead lists into clean Excel files for free. MarqClean AI formats names, columns, emails, phones, companies, and industries before XLSX export.",
    keyword: "CSV to Excel cleaner",
    h2: "Convert messy CSV lead lists into clean Excel workbooks",
    sections: [
      {
        h3: "CSV input with Excel-ready output",
        body: "Use the browser-based upload workflow to parse CSV rows, clean marketer-specific lead data, and download a professional XLSX workbook with readable columns.",
      },
      {
        h3: "Built for Microsoft 365 workflows",
        body: "Cleaned files are designed for Excel, Microsoft 365 reporting, CRM imports, spreadsheet QA, and campaign operations where consistent fields matter.",
      },
      {
        h3: "No signup and no payment required",
        body: "The core CSV to Excel cleaner is free to use and runs in your browser, helping teams avoid repetitive spreadsheet cleanup before marketing activation.",
      },
    ],
  },
  "excel-data-cleaner-for-marketers": {
    slug: "excel-data-cleaner-for-marketers",
    category: "Cleaners",
    cardTitle: "Excel Data Cleaner",
    icon: "🧹",
    title: `Free Excel Data Cleaner for Marketers | ${PRODUCT_NAME}`,
    h1: "Free Excel Data Cleaner for Marketers",
    description:
      "Clean Excel lead lists for free. Standardize names, columns, emails, phones, company fields, and industries before campaign launch.",
    keyword: "Excel data cleaner for marketers",
    h2: "Clean Excel marketing spreadsheets without manual formatting",
    sections: [
      {
        h3: "Repair marketer-specific spreadsheet fields",
        body: "MarqClean AI focuses on lead list fields that growth teams use every day, including full name, first name, last name, email, phone, company, job title, website, city, state, and industry.",
      },
      {
        h3: "Prepare outreach-ready data",
        body: "The cleaner flags missing names, companies, emails, and duplicate email records so marketers can review issues before uploading data to campaign systems.",
      },
      {
        h3: "Export a clean Excel workbook",
        body: "Download an XLSX file that is simple to open, review, share, and import into the tools your marketing team already uses.",
      },
    ],
  },
  "lead-list-cleaner": {
    slug: "lead-list-cleaner",
    category: "Cleaners",
    cardTitle: "Marketing Lead List & CSV Cleaner",
    icon: "🧹",
    title: `Free Lead List Cleaner | ${PRODUCT_NAME}`,
    h1: "Free Lead List Cleaner for CRM and Email Campaigns",
    description:
      "Clean marketing lead lists for free. Fix messy CSV and Excel files, improve CRM imports, and prepare outreach-ready spreadsheet data.",
    keyword: "lead list cleaner",
    h2: "Turn messy lead lists into CRM-ready campaign data",
    sections: [
      {
        h3: "Clean contacts before CRM import",
        body: "Format lead names, companies, job titles, emails, phones, websites, and locations before sending data into HubSpot, Salesforce, Mailchimp, or internal spreadsheets.",
      },
      {
        h3: "Reduce manual data entry",
        body: "The workflow replaces repetitive spreadsheet edits with a single upload, preview, and download process that can be reused across campaigns.",
      },
      {
        h3: "Free for global marketing teams",
        body: "Use MarqClean AI as a free lead list cleaner for agencies, startups, consultants, sales teams, nonprofit campaigns, and enterprise marketing operations.",
      },
    ],
  },
  "remove-duplicates-from-csv": {
    slug: "remove-duplicates-from-csv",
    category: "Cleaners",
    cardTitle: "CSV Duplicate Remover",
    icon: "🧹",
    title: `Free Remove Duplicates from CSV Tool | ${PRODUCT_NAME}`,
    h1: "Free Remove Duplicates from CSV Tool for Lead Lists",
    description:
      "Find duplicate email records in CSV and Excel lead lists for free. Clean outreach data and export campaign-ready files.",
    keyword: "remove duplicates from CSV",
    h2: "Detect duplicate email records while cleaning CSV files",
    sections: [
      {
        h3: "Duplicate detection for marketing records",
        body: "MarqClean AI checks repeated email addresses and marks duplicate rows in the data quality notes so marketers can review records before outreach.",
      },
      {
        h3: "Avoid campaign waste",
        body: "Cleaner lead lists help reduce repeated sends, messy CRM imports, inaccurate reporting, and poor personalization caused by duplicate contact records.",
      },
      {
        h3: "Export after review",
        body: "Download the cleaned CSV or Excel file after duplicates are flagged, formatted, and prepared for your CRM or marketing automation workflow.",
      },
    ],
  },
  "fix-csv-capitalization": {
    slug: "fix-csv-capitalization",
    category: "Formatters",
    cardTitle: "CSV Capitalization Fixer",
    icon: "🎯",
    title: `Free Fix CSV Capitalization Tool | ${PRODUCT_NAME}`,
    h1: "Free Fix CSV Capitalization Tool for Names and Companies",
    description:
      "Fix CSV capitalization for free. Standardize contact names, company names, job titles, cities, states, and countries in marketing lead lists.",
    keyword: "fix CSV capitalization",
    h2: "Fix capitalization issues in CSV and Excel lead data",
    sections: [
      {
        h3: "Clean names and companies",
        body: "The cleaner applies readable capitalization to names, job titles, company names, cities, states, and countries while preserving common marketing acronyms.",
      },
      {
        h3: "Improve personalization",
        body: "Better capitalization helps subject lines, email greetings, CRM records, and sales handoffs look more professional and less automated.",
      },
      {
        h3: "Download formatted files",
        body: "Export clean CSV or Excel output after capitalization, header mapping, field detection, and data quality notes are applied.",
      },
    ],
  },
  "categorize-company-industries-csv": {
    slug: "categorize-company-industries-csv",
    category: "Formatters",
    cardTitle: "Company Industry Categorizer",
    icon: "🎯",
    title: `Free Categorize Company Industries from CSV | ${PRODUCT_NAME}`,
    h1: "Free Categorize Company Industries from CSV",
    description:
      "Categorize company industries in CSV and Excel lead lists for free. Enrich marketing data for segmentation, personalization, and reporting.",
    keyword: "categorize company industries CSV",
    h2: "Add campaign-friendly industry categories to lead lists",
    sections: [
      {
        h3: "Industry enrichment for segmentation",
        body: "MarqClean AI reviews company, website, existing industry, and job title evidence to add practical categories like SaaS, healthcare, finance, ecommerce, real estate, and marketing services.",
      },
      {
        h3: "Confidence labels included",
        body: "Each categorized record includes an industry confidence field so teams can filter high-confidence segments or review uncertain matches.",
      },
      {
        h3: "Built for marketing campaigns",
        body: "Industry categories help marketers personalize copy, build audience segments, route leads, and improve spreadsheet reporting without manual tagging.",
      },
    ],
  },
  "crm-data-cleanup-tool": {
    slug: "crm-data-cleanup-tool",
    category: "Cleaners",
    cardTitle: "CRM Data Cleanup Tool",
    icon: "🧹",
    title: `Free CRM Data Cleanup Tool | ${PRODUCT_NAME}`,
    h1: "Free CRM Data Cleanup Tool for CSV and Excel Imports",
    description:
      "Clean CRM import files for free. Format CSV and Excel lead data, repair columns, flag missing fields, and export CRM-ready spreadsheets.",
    keyword: "CRM data cleanup tool",
    h2: "Prepare cleaner CRM import files before upload",
    sections: [
      {
        h3: "Standardize common CRM fields",
        body: "Map messy headers into practical fields such as first name, last name, full name, email, phone, company, job title, industry, website, city, state, and country.",
      },
      {
        h3: "Reduce import errors",
        body: "By fixing field placement and flagging missing outreach data, MarqClean AI helps teams catch avoidable issues before CRM import.",
      },
      {
        h3: "Export clean CRM-ready files",
        body: "Download a clean CSV or Excel file that is easier to review, enrich, segment, and import into your CRM platform.",
      },
    ],
  },
  "email-list-data-cleaner": {
    slug: "email-list-data-cleaner",
    category: "Cleaners",
    cardTitle: "Email List Cleaner",
    icon: "🧹",
    title: `Free Email List Data Cleaner | ${PRODUCT_NAME}`,
    h1: "Free Email List Data Cleaner for Marketing Campaigns",
    description:
      "Clean email list data for free. Fix contact fields, standardize names and companies, identify missing emails, and export clean CSV or Excel files.",
    keyword: "email list data cleaner",
    h2: "Clean email campaign lists before personalization and sending",
    sections: [
      {
        h3: "Fix contacts and campaign fields",
        body: "Clean common email marketing fields including contact name, email address, phone, company, job title, website, location, and industry category.",
      },
      {
        h3: "Flag data quality issues",
        body: "Missing emails, missing contact names, missing companies, and duplicate email addresses are marked in data quality notes for quick review.",
      },
      {
        h3: "Download for outreach tools",
        body: "Export clean lists for email platforms, CRM segments, outreach tools, Microsoft Excel, Google Sheets imports, and marketing reports.",
      },
    ],
  },
  "microsoft-excel-lead-cleaner": {
    slug: "microsoft-excel-lead-cleaner",
    category: "Cleaners",
    cardTitle: "Excel Lead Cleaner",
    icon: "🧹",
    title: `Free Microsoft Excel Lead Cleaner | ${PRODUCT_NAME}`,
    h1: "Free Microsoft Excel Lead Cleaner for Marketing Teams",
    description:
      "Clean Excel lead lists for free. MarqClean AI prepares Microsoft Excel-ready marketing spreadsheets from messy CSV and XLSX files.",
    keyword: "Microsoft Excel lead cleaner",
    h2: "Create cleaner Excel lead files for Microsoft 365 workflows",
    sections: [
      {
        h3: "Excel-friendly export",
        body: "Download a normal XLSX workbook with clean headers, readable columns, formatted values, industry enrichment, and data quality notes.",
      },
      {
        h3: "Built for spreadsheet operators",
        body: "The workflow supports marketing coordinators, agency teams, campaign managers, analysts, and operators who rely on Excel for lead list QA.",
      },
      {
        h3: "Keeps Excel formatting intact",
        body: "Column widths, header styling, and cell types are preserved on export, so the workbook opens looking like a normal Excel file rather than a flat data dump.",
      },
    ],
  },
  "csv-to-json-converter": {
    slug: "csv-to-json-converter",
    category: "Converters",
    cardTitle: "CSV to JSON Converter",
    icon: "🔄",
    title: `Free CSV to JSON Converter | ${PRODUCT_NAME}`,
    h1: "Free CSV to JSON Converter Online",
    description:
      "Convert CSV and Excel files to JSON for free in your browser. Upload data and download clean JSON for APIs, apps, and integrations.",
    keyword: "CSV to JSON converter",
    h2: "Turn spreadsheet rows into structured JSON",
    sections: [
      {
        h3: "Upload CSV, TXT, or Excel",
        body: "MarqClean AI reads your header row and converts each record into a JSON object with matching keys, ready for developers and integrations.",
      },
      {
        h3: "Clean before you convert",
        body: "Remove duplicates, drop empty rows, trim whitespace, and sort data before exporting so your JSON payload stays tidy.",
      },
      {
        h3: "Download JSON instantly",
        body: "Export formatted JSON in the browser with no signup, no payment, and no server upload of your data.",
      },
    ],
  },
  "csv-to-xml-converter": {
    slug: "csv-to-xml-converter",
    category: "Converters",
    cardTitle: "CSV to XML Converter",
    icon: "🔄",
    title: `Free CSV to XML Converter | ${PRODUCT_NAME}`,
    h1: "Free CSV to XML Converter Online",
    description:
      "Convert CSV and Excel files to XML for free. Generate clean XML records from spreadsheet data for legacy systems and integrations.",
    keyword: "CSV to XML converter",
    h2: "Convert spreadsheet data into portable XML",
    sections: [
      {
        h3: "Structured record output",
        body: "Each row becomes an XML record with safe element names derived from your headers, making the output easy to import into other systems.",
      },
      {
        h3: "Prepare data first",
        body: "Use the built-in cleaning tools to remove duplicates, filter rows, and standardize values before you generate XML.",
      },
      {
        h3: "Valid, well-formed XML",
        body: "Special characters are escaped and element names are sanitized automatically, so the output parses cleanly in legacy systems, ETL tools, and XML validators without manual fixes.",
      },
    ],
  },
  "ai-csv-editor-assistant": {
    slug: "ai-csv-editor-assistant",
    category: "AI Tools",
    cardTitle: "AI CSV Editor",
    icon: "🤖",
    title: `Free AI CSV Editor and Assistant | ${PRODUCT_NAME}`,
    h1: "Free AI CSV Editor and Data Assistant",
    description:
      "Talk to your CSV data in plain English. Clean, transform, summarize, and analyze spreadsheets for free with the MarqClean AI data assistant.",
    keyword: "AI CSV editor",
    h2: "Chat with your spreadsheet to clean and analyze data",
    sections: [
      {
        h3: "Natural language commands",
        body: "Type commands like remove empty rows, remove duplicates, sum of amount, sort by date, or summarize, and the assistant performs the task.",
      },
      {
        h3: "Instant transformations",
        body: "Cleaning and analysis commands update the live spreadsheet preview so you can review changes before you export.",
      },
      {
        h3: "No formulas required",
        body: "MarqClean AI makes data analysis accessible for operations, finance, compliance, and data teams without complex spreadsheet formulas.",
      },
    ],
  },
  "csv-statistics-and-charts": {
    slug: "csv-statistics-and-charts",
    category: "Data Modelling",
    cardTitle: "CSV Statistics & Charts",
    icon: "📊",
    title: `Free CSV Statistics and Charts Tool | ${PRODUCT_NAME}`,
    h1: "Free CSV Statistics and Charts Generator",
    description:
      "Generate statistics and charts from CSV and Excel files for free. Get column summaries, totals, averages, unique counts, and bar charts.",
    keyword: "CSV statistics and charts",
    h2: "Summarize and visualize spreadsheet data in seconds",
    sections: [
      {
        h3: "Automatic column statistics",
        body: "See filled, empty, and unique counts for every column, plus sum, average, minimum, maximum, and median for numeric columns.",
      },
      {
        h3: "Built-in chart builder",
        body: "Group by any column and measure counts or numeric sums to produce a quick bar chart for reporting and analysis.",
      },
      {
        h3: "Export a data report",
        body: "Download a plain-text data report summarizing rows, columns, and key statistics for sharing and documentation.",
      },
    ],
  },
  "gpt-for-spreadsheets": {
    slug: "gpt-for-spreadsheets",
    category: "AI Tools",
    cardTitle: "GPT for Spreadsheets",
    icon: "🤖",
    title: `Free GPT for Spreadsheets Bulk Tools | ${PRODUCT_NAME}`,
    h1: "Free GPT for Spreadsheets Bulk Data Tools",
    description:
      "Apply AI style bulk operations to every row of a CSV or Excel file for free. Extract, classify, summarize, and reformat columns without formulas.",
    keyword: "GPT for spreadsheets",
    h2: "Process every row in one click with bulk column operations",
    sections: [
      {
        h3: "Bulk column operations",
        body: "Extract emails, phone numbers, numbers, and domains, or reformat text with title case, uppercase, and lowercase across an entire column.",
      },
      {
        h3: "Classify and summarize rows",
        body: "Classify sentiment and length, summarize long text, and count words for every row, then add the result as a new column.",
      },
      {
        h3: "No API key required",
        body: "The bulk tools run locally in your browser using rule-based logic, so you can process data quickly and privately for free.",
      },
    ],
  },
  "ai-document-generator": {
    slug: "ai-document-generator",
    category: "AI Tools",
    cardTitle: "AI Document Generator",
    icon: "🤖",
    title: `Free AI Document Generator from CSV | ${PRODUCT_NAME}`,
    h1: "Free AI Document Generator from Spreadsheet Data",
    description:
      "Generate personalized documents from CSV and Excel data for free. Use templates with placeholders to create outreach emails, letters, and notes.",
    keyword: "AI document generator",
    h2: "Turn spreadsheet rows into personalized documents",
    sections: [
      {
        h3: "Template placeholders",
        body: "Use fields like name, company, email, and industry as placeholders, and the generator fills a document for every row of your data.",
      },
      {
        h3: "Ready-made and custom templates",
        body: "Choose outreach emails, cover letters, invoice notes, or onboarding messages, or write your own custom template.",
      },
      {
        h3: "Export in one click",
        body: "Preview the first document and download all generated documents as a single text file for review and sending.",
      },
    ],
  },
  "ai-business-analyst": {
    slug: "ai-business-analyst",
    category: "AI Tools",
    cardTitle: "AI Business Analyst",
    icon: "🤖",
    title: `Free AI Business Analyst Artefacts | ${PRODUCT_NAME}`,
    h1: "Free AI Business Analyst Document Generator",
    description:
      "Generate business analyst artefacts from your data for free. Create requirements documents, test cases, data dictionaries, and user stories.",
    keyword: "AI business analyst",
    h2: "Generate BA artefacts from your dataset structure",
    sections: [
      {
        h3: "Requirements and user stories",
        body: "Produce a business requirements document and user stories derived from your dataset fields and record counts.",
      },
      {
        h3: "Test cases and data dictionary",
        body: "Generate test cases per column and a data dictionary that documents field types, population, unique values, and samples.",
      },
      {
        h3: "Documentation you can export",
        body: "Download each artefact as a text file to support planning, QA, onboarding, and stakeholder review.",
      },
    ],
  },
  "bank-statement-converter": {
    slug: "bank-statement-converter",
    category: "Converters",
    cardTitle: "Bank Statement Converter",
    icon: "🔄",
    title: `Free Bank Statement Converter to Excel | ${PRODUCT_NAME}`,
    h1: "Free Bank Statement Converter to Excel, CSV, and QIF",
    description:
      "Convert PDF bank statements into clean Excel, CSV, and QIF files for free. Extract dates, descriptions, amounts, and balances for accounting.",
    keyword: "bank statement converter",
    h2: "Turn PDF bank statements into import-ready spreadsheets",
    sections: [
      {
        h3: "Any bank, any layout",
        body: "Upload a digital PDF statement or paste statement text and the converter detects transaction dates, descriptions, amounts, and balances automatically.",
      },
      {
        h3: "Normalized and accounting-ready",
        body: "Dates and amounts are standardized so the output imports cleanly into QuickBooks, Xero, Sage, Excel, and other tools that read Excel, CSV, or QIF.",
      },
      {
        h3: "Also exports QIF for Quicken",
        body: "Alongside Excel and CSV, the converter can output QIF for direct import into Quicken and other personal finance software, which the Excel-only and CSV-only converters do not support.",
      },
    ],
  },
  "pdf-to-excel-converter": {
    slug: "pdf-to-excel-converter",
    category: "Converters",
    cardTitle: "PDF to Excel Converter",
    icon: "🔄",
    title: `Free PDF to Excel Converter for Statements | ${PRODUCT_NAME}`,
    h1: "Free PDF to Excel Converter for Bank Statements",
    description:
      "Convert PDF statements to Excel for free. Extract every transaction into clean spreadsheet rows ready for reporting and reconciliation.",
    keyword: "PDF to Excel converter",
    h2: "Extract transactions from PDF into a clean Excel file",
    sections: [
      {
        h3: "Structured transaction rows",
        body: "The converter reads PDF text, groups it into lines, and structures each transaction into date, description, amount, and balance columns.",
      },
      {
        h3: "Clean Excel output",
        body: "Download a formatted Excel workbook with a frozen header, autofilter, and normalized values ready for accounting work.",
      },
      {
        h3: "Handles multi-page statements",
        body: "Transactions spanning multiple statement pages are combined into a single continuous table, so a 12-page PDF becomes one clean Excel sheet instead of 12 separate exports.",
      },
    ],
  },
  "pdf-bank-statement-to-csv": {
    slug: "pdf-bank-statement-to-csv",
    category: "Converters",
    cardTitle: "PDF Bank Statement to CSV",
    icon: "🔄",
    title: `Free PDF Bank Statement to CSV | ${PRODUCT_NAME}`,
    h1: "Free PDF Bank Statement to CSV Converter",
    description:
      "Convert PDF bank statements to CSV for free. Turn transactions into a clean CSV file for spreadsheets, imports, and analysis.",
    keyword: "PDF bank statement to CSV",
    h2: "Export bank statement transactions as clean CSV",
    sections: [
      {
        h3: "Simple CSV export",
        body: "Extract transactions from a PDF or pasted statement text and download a tidy CSV with standardized dates and amounts.",
      },
      {
        h3: "Ready for any tool",
        body: "Use the CSV in Excel, Google Sheets, accounting software, or your own data pipeline for reconciliation and reporting.",
      },
      {
        h3: "Minimal, script-friendly output",
        body: "CSV has no formatting overhead, which makes it the simplest format to feed into scripts, database imports, or automated pipelines that don't need Excel-specific formatting.",
      },
    ],
  },
  "data-reconciliation-tool": {
    slug: "data-reconciliation-tool",
    category: "Reconciliation",
    cardTitle: "Data Reconciliation Tool",
    icon: "⚖️",
    title: `Free Data Reconciliation Tool | ${PRODUCT_NAME}`,
    h1: "Free Data Reconciliation and Validation Tool",
    description:
      "Reconcile Excel, CSV, PDF, and Word files for free. Auto-detect fields, run reconciliation macros, and download colour-coded Excel reports with exceptions and data quality checks.",
    keyword: "data reconciliation tool",
    h2: "Reconcile and validate records across files without technical skills",
    sections: [
      {
        h3: "Any file, one dataset",
        body: "Upload Excel, CSV, PDF, or Word files and the platform extracts and normalizes them into a common structured dataset before comparison.",
      },
      {
        h3: "Exact, fuzzy, and standardized matching",
        body: "Compare selected fields with exact, fuzzy, near-address, and formatting-insensitive matching to find matches, partial matches, mismatches, and missing values.",
      },
      {
        h3: "Colour-coded Excel reports",
        body: "Download a workbook with summary, detailed reconciliation, exceptions-only, and data quality sheets, all colour-coded for fast review.",
      },
    ],
  },
  "excel-file-comparison": {
    slug: "excel-file-comparison",
    category: "Reconciliation",
    cardTitle: "Excel & CSV File Comparison",
    icon: "⚖️",
    title: `Free Excel File Comparison Tool | ${PRODUCT_NAME}`,
    h1: "Free Excel and CSV File Comparison Tool",
    description:
      "Compare two Excel or CSV files for free. Detect matches, differences, missing values, and duplicates, then export a colour-coded reconciliation report.",
    keyword: "Excel file comparison",
    h2: "Compare spreadsheets and highlight every difference",
    sections: [
      {
        h3: "Auto field mapping",
        body: "The tool detects columns, headings, and likely field mappings automatically, and lets you adjust the mapping before comparing.",
      },
      {
        h3: "Record-level results",
        body: "See field-level source and target values with a clear status for every record, including partial matches that need review.",
      },
      {
        h3: "High-volume ready",
        body: "Compare thousands of records in the browser and export professional Excel reports for audit and sign-off.",
      },
    ],
  },
  "document-validation-tool": {
    slug: "document-validation-tool",
    category: "Reconciliation",
    cardTitle: "Document Validation Tool",
    icon: "⚖️",
    title: `Free Document Validation Tool | ${PRODUCT_NAME}`,
    h1: "Free Document Validation and Data Quality Tool",
    description:
      "Validate documents and data for free. Check missing fields, invalid formats, duplicates, and incomplete records across Excel, CSV, PDF, and Word files.",
    keyword: "document validation tool",
    h2: "Validate data quality and compliance fields in seconds",
    sections: [
      {
        h3: "Compliance and KYC fields",
        body: "Validate identity, address, contact, compliance, banking, and account fields such as ID numbers, tax numbers, and postal codes.",
      },
      {
        h3: "Data quality checks",
        body: "Identify missing values, invalid emails, invalid postal codes, duplicate records, and other anomalies in a dedicated report.",
      },
      {
        h3: "Private and free",
        body: "All extraction and validation runs in your browser with no data stored on servers, so sensitive documents stay on your device.",
      },
    ],
  },
  "pdf-vs-excel-verification": {
    slug: "pdf-vs-excel-verification",
    category: "Reconciliation",
    cardTitle: "PDF vs Excel Verification",
    icon: "⚖️",
    title: `Free PDF vs Excel Verification | ${PRODUCT_NAME}`,
    h1: "Free PDF vs Excel Data Verification",
    description:
      "Verify unstructured PDF client data against structured Excel master data for free. Extract, map, and compare names, addresses, and postal codes with colour-coded reports.",
    keyword: "PDF vs Excel verification",
    h2: "Compare extracted PDF records against Excel master rows",
    sections: [
      {
        h3: "Structure the PDF automatically",
        body: "Free-text PDF records are structured into full name, title, address lines, and postal code, then mapped to your Excel columns.",
      },
      {
        h3: "Name, address, and postal matching",
        body: "Exact, case-insensitive, title-removal, and fuzzy matching handle real-world differences between PDF text and Excel fields.",
      },
      {
        h3: "Audit-ready Excel report",
        body: "Download a six-sheet workbook with executive summary, detailed verification, exceptions, PDF extraction output, master data, and data quality.",
      },
    ],
  },
  "mailing-list-verification": {
    slug: "mailing-list-verification",
    category: "Reconciliation",
    cardTitle: "Mailing List Verification",
    icon: "⚖️",
    title: `Free Mailing List Verification Tool | ${PRODUCT_NAME}`,
    h1: "Free Mailing List Verification Against Excel Master Data",
    description:
      "Verify PDF mailing lists against Excel master data for free. Check shareholder mailing records, addresses, postal codes, email, and campaign types.",
    keyword: "mailing list verification",
    h2: "Check client mailing data for financial and investor services",
    sections: [
      {
        h3: "Built for transfer secretarial teams",
        body: "Verify shareholder mailing data, issuer, addresses, and campaign types against a structured Excel master file.",
      },
      {
        h3: "Campaign and contact validation",
        body: "Flag Email campaigns with missing emails, Post campaigns with incomplete addresses, and invalid or scientific-notation cell numbers.",
      },
      {
        h3: "Colour-coded exceptions",
        body: "Green, blue, red, and orange results highlight exact matches, partial matches, mismatches, and missing data for fast review.",
      },
    ],
  },
  "merge-spreadsheets": {
    slug: "merge-spreadsheets",
    category: "Formatters",
    cardTitle: "Spreadsheet Merger",
    icon: "🎯",
    title: `Free Merge Spreadsheets Tool | ${PRODUCT_NAME}`,
    h1: "Free Merge Spreadsheets and Combine CSV Files",
    description:
      "Merge and combine multiple CSV and Excel files for free. Append rows with smart column alignment or join files on a key column, then export Excel or CSV.",
    keyword: "merge spreadsheets",
    h2: "Combine CSV and Excel files with smart column matching",
    sections: [
      { h3: "Append or join", body: "Stack rows from multiple files, automatically aligning matching columns even if the files have a different number of columns, or join a second file on a shared key column the way a VLOOKUP or database join would." },
      { h3: "Smart column matching", body: "Columns are matched by header name rather than position, so files exported in a different column order still combine correctly instead of mixing unrelated fields together." },
      { h3: "Handles mismatched schemas", body: "If one file has a column the others don't, it's kept and filled with blanks for the rows that lack it, rather than the merge failing or silently dropping data." },
    ],
  },
  "remove-duplicates-excel": {
    slug: "remove-duplicates-excel",
    category: "Cleaners",
    cardTitle: "Excel Duplicate Remover",
    icon: "🧹",
    title: `Free Remove Duplicates in Excel | ${PRODUCT_NAME}`,
    h1: "Free Remove Duplicates in Excel and CSV",
    description:
      "Remove duplicate rows in Excel and CSV files for free. Deduplicate by selected columns and find near-duplicate records with fuzzy matching.",
    keyword: "remove duplicates in excel",
    h2: "Delete exact and near-duplicate rows in seconds",
    sections: [
      { h3: "Column-based dedupe", body: "Remove duplicates based on the whole row, or only the key columns you choose, such as email or account number, so records that differ only in an unrelated field still get flagged correctly." },
      { h3: "Fuzzy duplicate finder", body: "Detect near-duplicate names and records that aren't identical character-for-character, like \"Jon Smith\" and \"John Smith\" at the same address, using similarity scoring instead of an exact match." },
      { h3: "Review before you delete", body: "Matches are shown with a similarity score before anything is removed, so you can confirm real duplicates and skip coincidental near-matches rather than deleting on blind trust." },
    ],
  },
  "fuzzy-name-matching": {
    slug: "fuzzy-name-matching",
    category: "Reconciliation",
    cardTitle: "Fuzzy Name & Address Matching",
    icon: "⚖️",
    title: `Free Fuzzy Name Matching Tool | ${PRODUCT_NAME}`,
    h1: "Free Fuzzy Name and Address Matching",
    description:
      "Similarity-based searching and matching for names and addresses. Free fuzzy matching for data cleansing, record linking, and compliance screening.",
    keyword: "fuzzy name matching",
    h2: "Match names and addresses beyond exact string metrics",
    sections: [
      { h3: "Approximate matching", body: "Find records that are highly similar but not identical, such as misspelled names, transposed initials, or abbreviated titles, improving match precision and recall over an exact-string comparison." },
      { h3: "Compliance-grade logic", body: "Standardisation and similarity scoring reduce false positives and missed matches in screening workflows, where both a missed real match and a false flag carry real cost." },
      { h3: "Tunable match threshold", body: "Adjust how strict the similarity scoring is, so you can trade off catching more potential matches against reviewing more false positives, depending on the workflow." },
    ],
  },
  "web-scraping-to-excel": {
    slug: "web-scraping-to-excel",
    category: "Converters",
    cardTitle: "Web Table Extractor",
    icon: "🔄",
    title: `Free Web Table to Excel Extractor | ${PRODUCT_NAME}`,
    h1: "Free Web Table and List to Excel Extractor",
    description:
      "Turn web tables and lists into structured spreadsheets for free. Paste HTML and export clean Excel or CSV data, no code required.",
    keyword: "web scraping to excel",
    h2: "Convert web data into structured spreadsheets",
    sections: [
      { h3: "No-code extraction", body: "Paste the HTML of a web table or list, and the tool structures it into rows and columns automatically, without needing to know CSS selectors or write a scraping script." },
      { h3: "Export anywhere", body: "Download the extracted data as Excel or CSV for analysis, CRM import, competitor tracking, or reporting, ready to open directly rather than needing further cleanup." },
      { h3: "Handles messy markup", body: "Nested tags, inconsistent spacing, and mixed formatting inside table cells are stripped down to plain text automatically, so the extracted columns are usable, not cluttered with leftover HTML." },
    ],
  },
  "timesheet-tool": {
    slug: "timesheet-tool",
    category: "Data Modelling",
    cardTitle: "Timesheet & Billable Hours Tracker",
    icon: "📊",
    title: `Free Timesheet Tool and Time Tracker | ${PRODUCT_NAME}`,
    h1: "Free Timesheet Tool and Billable Hours Tracker",
    description:
      "Track hours across projects and generate billable time reports for free. Build a timesheet and export a clean Excel report.",
    keyword: "timesheet tool",
    h2: "Log hours and export billable time reports",
    sections: [
      { h3: "Project time tracking", body: "Record date, project, task, hours, and billable status for every entry, building a running log you can add to throughout the week rather than reconstructing hours from memory at invoicing time." },
      { h3: "Instant summaries", body: "See total, billable, and non-billable hours and a breakdown by project update live as you type, so you can spot an under-logged week before it becomes a billing problem." },
      { h3: "Client-ready Excel export", body: "Download a formatted Excel timesheet with totals already calculated, ready to attach to an invoice or hand to a client without reformatting it first." },
    ],
  },
  "advanced-excel-functions": {
    slug: "advanced-excel-functions",
    category: "Excel Formulas",
    cardTitle: "Advanced Excel Functions",
    icon: "🧮",
    title: `Free Advanced Excel Functions Tool | ${PRODUCT_NAME}`,
    h1: "Free Advanced Excel Functions and Formula Engine",
    description:
      "Run 35+ advanced Excel functions online for free: XLOOKUP, VLOOKUP, INDEX MATCH, SUMIFS, COUNTIFS, AVERAGEIFS, SUMPRODUCT, FILTER, UNIQUE, SORT, IF, IFS, SWITCH, LEFT, MID, TEXTSPLIT, DATEDIF, NETWORKDAYS, RANK, PERCENTILE, LET, XNPV, XIRR, PMT, and IPMT. Get results in seconds.",
    keyword: "advanced excel functions",
    h2: "Turn a basic spreadsheet into a data analysis engine",
    sections: [
      { h3: "Pick or type an action", body: "Choose a function from a dropdown or type what you want in plain English, and the tool runs it automatically on your data." },
      { h3: "Every essential function", body: "Lookup and reference, conditional maths, statistics, text, date and time, logical, dynamic arrays, and financial modelling functions are all included." },
      { h3: "See the Excel formula", body: "Each result shows the equivalent Excel formula so you can reuse it in your own workbook, plus export derived tables." },
    ],
  },
  "xlookup-online": {
    slug: "xlookup-online",
    category: "Lookups",
    cardTitle: "XLOOKUP & INDEX MATCH Generator",
    icon: "🔎",
    title: `Free XLOOKUP and INDEX MATCH Online | ${PRODUCT_NAME}`,
    h1: "Free XLOOKUP and INDEX MATCH Online Tool",
    description:
      "Run XLOOKUP and INDEX MATCH online for free. Search columns left or right, return values from anywhere in a table, and set a default when no match is found.",
    keyword: "XLOOKUP online",
    h2: "Multidirectional lookups without writing formulas",
    sections: [
      { h3: "Left or right lookups", body: "Unlike VLOOKUP, which only searches left to right, XLOOKUP searches in either direction and returns any column you choose, so the lookup column no longer has to sit to the left of the result." },
      { h3: "INDEX and MATCH", body: "Look up values anywhere in a table regardless of column position, with a default value when no match exists instead of a raw #N/A error breaking the rest of a report." },
      { h3: "See exactly which rows matched", body: "Results show which source row each match came from, so you can verify the lookup logic worked correctly before trusting it in a larger workbook." },
    ],
  },
  "sumifs-countifs-tool": {
    slug: "sumifs-countifs-tool",
    category: "Excel Formulas",
    cardTitle: "SUMIFS, COUNTIFS & FILTER Tool",
    icon: "🧮",
    title: `Free SUMIFS and COUNTIFS Tool | ${PRODUCT_NAME}`,
    h1: "Free SUMIFS, COUNTIFS, and FILTER Tool",
    description:
      "Run SUMIFS, COUNTIFS, SUMPRODUCT, FILTER, UNIQUE, and SORT online for free. Apply multiple criteria and extract data subsets in seconds.",
    keyword: "SUMIFS COUNTIFS tool",
    h2: "Conditional calculations and dynamic arrays online",
    sections: [
      { h3: "Multi-criteria maths", body: "Sum or count across multiple columns with several conditions at once, for example total sales for one region in one month, without nesting several IF statements to get there." },
      { h3: "Dynamic arrays", body: "Use FILTER, UNIQUE, and SORT to extract, deduplicate, and order data subsets automatically, producing a live-feeling result table instead of a static one-off calculation." },
      { h3: "Weighted totals with SUMPRODUCT", body: "SUMPRODUCT multiplies columns row by row before summing, which is the standard way to calculate weighted averages, revenue from price times quantity, or other conditional sums that a plain SUM can't express." },
    ],
  },
};

const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function sanitizeCell(value: unknown) {
  return String(value ?? "")
    .replace(/\uFEFF/g, "")
    .replace(/[\t\n\r]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function aliasKey(header: string) {
  return header.toLowerCase().replace(/[^a-z0-9-]/g, "");
}

function smartTitleCase(input: string) {
  return sanitizeCell(input)
    .toLowerCase()
    .split(" ")
    .map((word) => {
      const compact = word.replace(/[^a-z0-9]/gi, "").toUpperCase();
      const acronym = Array.from(ACRONYMS).find((item) => item.toUpperCase() === compact);

      if (acronym) return word.replace(compact.toLowerCase(), acronym);
      if (["and", "or", "of", "the", "in", "for", "to", "de", "la", "von"].includes(word)) return word;

      return word
        .split("-")
        .map((segment) =>
          segment
            .split("'")
            .map((part) => (part ? part.charAt(0).toUpperCase() + part.slice(1) : part))
            .join("'")
        )
        .join("-");
    })
    .join(" ")
    .replace(/\b(Llc|Inc|Ltd|Co)\b/g, (match) => match.toUpperCase());
}

function titleCasePerson(input: string) {
  return smartTitleCase(input)
    .split(" ")
    .map((word, index) => (index === 0 ? word.charAt(0).toUpperCase() + word.slice(1) : word))
    .join(" ");
}

function normalizeHeader(header: string, index: number) {
  const cleaned = sanitizeCell(header);
  const key = aliasKey(cleaned);

  if (HEADER_ALIASES[key]) return HEADER_ALIASES[key];
  if (!cleaned) return `Column ${index + 1}`;

  return smartTitleCase(cleaned.replace(/[_-]+/g, " "));
}

function isLikelyPhone(value: string) {
  const candidate = sanitizeCell(value);
  if (!candidate || !/[0-9]/.test(candidate) || !/^[+()\d\s.-]+(?:\s*(?:ext|x)\s*\d+)?$/i.test(candidate)) return false;
  const parsed = parsePhoneNumberFromString(candidate, { defaultCountry: "ZA", extract: false });
  return Boolean(parsed?.isPossible());
}

function formatPhone(value: string) {
  const candidate = sanitizeCell(value);
  const parsed = parsePhoneNumberFromString(candidate, { defaultCountry: "ZA", extract: false });
  if (!parsed || !parsed.isPossible()) return candidate;
  return parsed.formatInternational();
}

function cleanWebsite(value: string) {
  const candidate = sanitizeCell(value);
  if (!candidate) return "";
  try {
    const url = new URL(/^https?:\/\//i.test(candidate) ? candidate : `https://${candidate}`);
    url.hash = url.hash;
    return url.toString().replace(/\/$/, "");
  } catch {
    return candidate;
  }
}

function splitName(fullName: string) {
  const parts = titleCasePerson(fullName).split(" ").filter(Boolean);

  if (parts.length === 0) return { firstName: "", lastName: "" };
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };

  return {
    firstName: parts[0],
    lastName: parts.slice(1).join(" "),
  };
}

function detectIndustry(row: CleanRow): IndustryMatch {
  const evidence = [row.Industry, row.Company, row.Website, row["Job Title"]].join(" ").toLowerCase();
  const match = INDUSTRY_RULES.find((rule) => rule.terms.some((term) => evidence.includes(term)));

  if (match) {
    return {
      industry: match.industry,
      confidence: row.Industry ? "High" : "Medium",
    };
  }

  return {
    industry: row.Industry ? smartTitleCase(row.Industry) : "General Business",
    confidence: row.Industry ? "Medium" : "Needs review",
  };
}

function addNote(row: CleanRow, note: string) {
  const existing = row["Data Quality Notes"];
  row["Data Quality Notes"] = existing ? `${existing}; ${note}` : note;
}

function createDownloadName(fileName: string, extension: "csv" | "xlsx") {
  const baseName = fileName.replace(/\.[^.]+$/, "") || "marketing-leads";
  return `${baseName}-marqclean-clean.${extension}`;
}

function isBlankRawRow(row: RawRow) {
  return Object.values(row).every((value) => !sanitizeCell(value));
}

function cleanDataset(inputRows: RawRow[], fileName: string): CleanResult {
  const usableRows = inputRows.filter((row) => !isBlankRawRow(row));
  const originalHeaders = Array.from(new Set(usableRows.flatMap((row) => Object.keys(row))));
  const stats: CleanStats = {
    totalRows: usableRows.length,
    cleanedRows: 0,
    headerFixes: 0,
    fieldRepairs: 0,
    industriesAdded: 0,
    missingFlags: 0,
    duplicates: 0,
  };

  const rows = usableRows
    .map((rawRow) => {
      const cleanRow: CleanRow = {};

      Object.entries(rawRow).forEach(([rawHeader, rawValue], index) => {
        const value = sanitizeCell(rawValue);
        const normalizedHeader = normalizeHeader(rawHeader, index);
        if (normalizedHeader !== sanitizeCell(rawHeader)) stats.headerFixes += 1;
        if (!value) return;

        let cleanedValue = value;
        switch (normalizedHeader) {
          case "Email":
            cleanedValue = value.toLowerCase();
            if (cleanedValue !== value) stats.fieldRepairs += 1;
            break;
          case "Phone":
            if (isLikelyPhone(value)) {
              cleanedValue = formatPhone(value);
              if (cleanedValue !== value) stats.fieldRepairs += 1;
            } else {
              addNote(cleanRow, "Phone needs review");
            }
            break;
          case "Website":
            cleanedValue = cleanWebsite(value);
            if (cleanedValue !== value) stats.fieldRepairs += 1;
            break;
          case "Full Name":
          case "First Name":
          case "Last Name":
            cleanedValue = titleCasePerson(value);
            if (cleanedValue !== value) stats.fieldRepairs += 1;
            break;
          case "Company":
          case "Job Title":
          case "City":
          case "State":
          case "Country":
          case "Industry":
            cleanedValue = smartTitleCase(value);
            if (cleanedValue !== value) stats.fieldRepairs += 1;
            break;
          default:
            break;
        }

        if (!cleanRow[normalizedHeader]) cleanRow[normalizedHeader] = cleanedValue;
      });

      if (cleanRow["Full Name"]) {
        const fullName = titleCasePerson(cleanRow["Full Name"]);
        const split = splitName(fullName);
        cleanRow["Full Name"] = fullName;
        cleanRow["First Name"] = cleanRow["First Name"] || split.firstName;
        cleanRow["Last Name"] = cleanRow["Last Name"] || split.lastName;
      } else if (cleanRow["First Name"] || cleanRow["Last Name"]) {
        cleanRow["First Name"] = titleCasePerson(cleanRow["First Name"] ?? "");
        cleanRow["Last Name"] = titleCasePerson(cleanRow["Last Name"] ?? "");
        cleanRow["Full Name"] = `${cleanRow["First Name"]} ${cleanRow["Last Name"]}`.trim();
      }

      const industry = detectIndustry(cleanRow);
      if (!cleanRow.Industry) stats.industriesAdded += 1;
      cleanRow.Industry = industry.industry;
      cleanRow["Industry Confidence"] = industry.confidence;

      if (!cleanRow.Email) {
        addNote(cleanRow, "Missing email");
        stats.missingFlags += 1;
      } else if (!emailPattern.test(cleanRow.Email)) {
        addNote(cleanRow, "Email needs review");
        stats.missingFlags += 1;
      }
      if (!cleanRow["Full Name"]) {
        addNote(cleanRow, "Missing contact name");
        stats.missingFlags += 1;
      }
      if (!cleanRow.Company) {
        addNote(cleanRow, "Missing company");
        stats.missingFlags += 1;
      }
      if (!cleanRow["Data Quality Notes"]) cleanRow["Data Quality Notes"] = "Ready for outreach";
      return cleanRow;
    })
    .filter((row) => Object.values(row).some(Boolean));

  const emailCounts = rows.reduce<Record<string, number>>((accumulator, row) => {
    if (row.Email) accumulator[row.Email] = (accumulator[row.Email] ?? 0) + 1;
    return accumulator;
  }, {});

  rows.forEach((row) => {
    if (row.Email && emailCounts[row.Email] > 1) {
      addNote(row, "Duplicate email");
      stats.duplicates += 1;
    }
  });

  stats.cleanedRows = rows.length;
  const discoveredHeaders = Array.from(new Set(rows.flatMap((row) => Object.keys(row))));
  const headers = [
    ...PRIORITY_HEADERS.filter((header) => discoveredHeaders.includes(header)),
    ...discoveredHeaders.filter((header) => !PRIORITY_HEADERS.includes(header)),
  ];
  return { fileName, rows, headers, stats, originalHeaders };
}
function decodeDelimitedText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.slice(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.slice(2));
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

async function parseCsvFile(file: File) {
  return new Promise<RawRow[]>((resolve, reject) => {
    file.arrayBuffer().then((buffer) => {
      const text = decodeDelimitedText(new Uint8Array(buffer));
      Papa.parse<RawRow>(text, {
        header: true,
        skipEmptyLines: "greedy",
        transformHeader: (header) => sanitizeCell(header),
        complete: (result) => {
          const blockingError = result.errors.find((parseError) => parseError.code !== "UndetectableDelimiter");
          if (blockingError) {
            reject(new Error(`CSV parsing failed: ${blockingError.message}`));
            return;
          }
          if (!result.data.length) {
            reject(new Error("The CSV file did not contain readable rows."));
            return;
          }
          resolve(result.data);
        },
        error: (error: Error) => reject(error),
      });
    }).catch(reject);
  });
}

async function readWorkbookMatrix(file: File): Promise<string[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    dense: true,
    cellDates: true,
    raw: false,
    dateNF: "yyyy-mm-dd",
  });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The Excel workbook does not contain a readable worksheet.");
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  return matrix.map((row) => row.map((value) => {
    if (value instanceof Date) {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, "0");
      const day = String(value.getDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
    return sanitizeCell(value);
  }));
}

async function parseExcelFile(file: File) {
  const table = await readWorkbookMatrix(file);
  const headerRowIndex = table.findIndex((row) => row.some((value) => sanitizeCell(value)));
  if (headerRowIndex === -1) throw new Error("The Excel worksheet is empty.");

  const maxColumns = table.reduce((max, row) => Math.max(max, row.length), 0);
  const headers = Array.from({ length: maxColumns || 1 }, (_, index) => sanitizeCell(table[headerRowIndex][index]) || `Column ${index + 1}`);

  return table.slice(headerRowIndex + 1).map((row) =>
    headers.reduce<RawRow>((record, header, index) => {
      record[header] = sanitizeCell(row[index]);
      return record;
    }, {})
  );
}

async function readFileMatrix(file: File): Promise<string[][]> {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "xls") return readWorkbookMatrix(file);
  const text = decodeDelimitedText(new Uint8Array(await file.arrayBuffer()));
  const parsed = Papa.parse<string[]>(text, { skipEmptyLines: "greedy" });
  return (parsed.data as string[][]).filter((row) => row.some((value) => sanitizeCell(value)));
}

function createConvertedFileName(fileName: string, extension: string) {
  const baseName = fileName.replace(/\.[^.]+$/, "") || "converted-data";
  return `${baseName}-marqclean.${extension}`;
}

// Bank statement conversion (BankPDF style): PDF, scan text, or pasted text into clean transactions
type BankTransaction = {
  date: string;
  description: string;
  amount: string;
  balance: string;
};

const MONTHS: Record<string, string> = {
  jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
  jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12",
};

function normalizeStatementDate(raw: string): string {
  const value = raw.trim();

  let match = value.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;

  match = value.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (match) {
    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    let year = match[3];
    if (year.length === 2) year = Number(year) > 60 ? `19${year}` : `20${year}`;
    return `${year}-${month}-${day}`;
  }

  match = value.match(/^(\d{1,2})[\s-]([A-Za-z]{3,})[\s-](\d{2,4})$/);
  if (match) {
    const month = MONTHS[match[2].slice(0, 3).toLowerCase()];
    if (month) {
      let year = match[3];
      if (year.length === 2) year = `20${year}`;
      return `${year}-${month}-${match[1].padStart(2, "0")}`;
    }
  }

  match = value.match(/^([A-Za-z]{3,})[\s-](\d{1,2}),?[\s-](\d{2,4})$/);
  if (match) {
    const month = MONTHS[match[1].slice(0, 3).toLowerCase()];
    if (month) {
      let year = match[3];
      if (year.length === 2) year = `20${year}`;
      return `${year}-${month}-${match[2].padStart(2, "0")}`;
    }
  }

  return value;
}

function normalizeStatementAmount(raw: string): string {
  let value = raw.trim();
  if (!value) return "";

  let negative = false;
  if (/^\(.*\)$/.test(value)) {
    negative = true;
    value = value.slice(1, -1);
  }
  if (/-\s*$/.test(value) || /^-/.test(value) || /\bDR\b/i.test(value)) negative = true;
  if (/\bCR\b/i.test(value)) negative = false;

  value = value.replace(/[^\d.,-]/g, "").replace(/(?!^)-/g, "");
  if (!value) return "";

  const comma = value.lastIndexOf(",");
  const dot = value.lastIndexOf(".");
  const hasComma = comma >= 0;
  const hasDot = dot >= 0;

  if (hasComma && hasDot) {
    // The rightmost separator is the decimal separator.
    if (comma > dot) value = value.replace(/\./g, "").replace(",", ".");
    else value = value.replace(/,/g, "");
  } else if (hasComma) {
    const decimals = value.length - comma - 1;
    value = decimals === 1 || decimals === 2 ? value.replace(",", ".") : value.replace(/,/g, "");
  } else if (hasDot) {
    const decimals = value.length - dot - 1;
    value = decimals === 1 || decimals === 2 ? value : value.replace(/\./g, "");
  }

  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  return (negative ? -Math.abs(number) : number).toFixed(2);
}
const STATEMENT_DATE_PATTERN = /(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{1,2}[\s-][A-Za-z]{3,}[\s-]\d{2,4}|[A-Za-z]{3,}[\s-]\d{1,2},?[\s-]\d{2,4})/;
const AMOUNT_TOKEN = /-?\(?[\d.,]+\)?(?:\s?(?:CR|DR))?/gi;

function extractBankTransactions(text: string): BankTransaction[] {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  const transactions: BankTransaction[] = [];

  for (const line of lines) {
    const dateMatch = line.match(STATEMENT_DATE_PATTERN);
    if (!dateMatch) continue;

    const date = normalizeStatementDate(dateMatch[0]);
    const remainder = line.slice((dateMatch.index ?? 0) + dateMatch[0].length).trim();
    const tokenMatches = [...remainder.matchAll(AMOUNT_TOKEN)].map((match) => ({
      token: match[0],
      index: match.index ?? 0,
      end: (match.index ?? 0) + match[0].length,
    }));
    if (!tokenMatches.length) continue;

    const lastToken = tokenMatches[tokenMatches.length - 1];
    const trailing = remainder.slice(lastToken.end).trim();
    // Amounts/balances must be in the trailing numeric portion of the statement.
    // A number embedded in a description (e.g. "POS PURCHASE 1234 WOOLWORTHS")
    // is therefore not treated as a transaction amount.
    if (trailing && !/^(?:CR|DR)$/i.test(trailing)) continue;

    const candidates = tokenMatches.slice(-2);
    const firstCandidate = candidates[0];
    const amountToken = candidates.length > 1 ? firstCandidate : lastToken;
    const balanceToken = candidates.length > 1 ? candidates[1] : undefined;

    const descriptionEnd = amountToken.index;
    const description = remainder.slice(0, descriptionEnd).replace(/[.\-\s]+$/, "").trim();
    const amount = normalizeStatementAmount(amountToken.token);
    const balance = balanceToken ? normalizeStatementAmount(balanceToken.token) : "";

    if (!amount && !balance) continue;
    transactions.push({
      date,
      description: description || "Transaction",
      amount,
      balance,
    });
  }

  return transactions;
}

function bankTransactionsToMatrix(transactions: BankTransaction[]): string[][] {
  const header = ["Date", "Description", "Amount", "Balance"];
  const rows = transactions.map((transaction) => [transaction.date, transaction.description, transaction.amount, transaction.balance]);
  return [header, ...rows];
}

function bankTransactionsToQif(transactions: BankTransaction[]): string {
  const lines: string[] = ["!Type:Bank"];

  transactions.forEach((transaction) => {
    lines.push(`D${transaction.date}`);
    lines.push(`T${transaction.amount || "0.00"}`);
    lines.push(`P${transaction.description}`);
    lines.push("^");
  });

  return lines.join("\n");
}

function columnIndexToReference(index: number) {
  let dividend = index + 1;
  let reference = "";
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    reference = String.fromCharCode(65 + modulo) + reference;
    dividend = Math.floor((dividend - 1) / 26);
  }
  return reference;
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getActiveDelimiterChars(settings: WizardSettings): string[] {
  const chars: string[] = [];

  (Object.keys(settings.delimiters) as DelimiterOption[]).forEach((key) => {
    if (!settings.delimiters[key]) return;

    if (key === "other") {
      if (settings.otherChar) chars.push(settings.otherChar);
    } else {
      chars.push(DELIMITER_CHAR_MAP[key]);
    }
  });

  return chars;
}

function parseSingleDelimited(text: string, delimiter: string, qualifier: TextQualifier, treatConsecutiveAsOne: boolean): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;

  const pushField = () => {
    row.push(field);
    field = "";
  };

  const pushRow = () => {
    pushField();
    if (treatConsecutiveAsOne) row = row.filter((value) => value !== "");
    rows.push(row);
    row = [];
  };

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (inQuotes) {
      if (qualifier && char === qualifier) {
        if (next === qualifier) {
          field += qualifier;
          index += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }

    if (qualifier && char === qualifier) {
      inQuotes = true;
      continue;
    }

    if (char === delimiter) {
      pushField();
      continue;
    }

    if (char === "\r") {
      if (next === "\n") index += 1;
      pushRow();
      continue;
    }

    if (char === "\n") {
      pushRow();
      continue;
    }

    field += char;
  }

  if (field.length > 0 || row.length > 0) pushRow();

  return rows.filter((cells) => cells.some((value) => value.length > 0));
}

function parseMultiDelimited(text: string, chars: string[], treatConsecutiveAsOne: boolean): string[][] {
  const escaped = chars.map(escapeRegExp).join("");
  const pattern = treatConsecutiveAsOne ? `[${escaped}]+` : `[${escaped}]`;
  const regex = new RegExp(pattern);

  return text
    .split(/\r\n|\r|\n/)
    .map((line) => line.split(regex))
    .filter((cells) => cells.some((value) => value.trim().length > 0));
}

function parseWithWizardSettings(text: string, settings: WizardSettings): string[][] {
  const chars = getActiveDelimiterChars(settings);

  if (!chars.length) {
    return text
      .split(/\r\n|\r|\n/)
      .filter((line) => line.length > 0)
      .map((line) => [line]);
  }

  if (chars.length === 1) {
    return parseSingleDelimited(text, chars[0], settings.textQualifier, settings.treatConsecutiveAsOne);
  }

  return parseMultiDelimited(text, chars, settings.treatConsecutiveAsOne);
}

function matrixToDelimitedText(matrix: string[][], delimiter: string, qualifier: TextQualifier) {
  const quote = qualifier || '"';

  return matrix
    .map((row) =>
      row
        .map((value) => {
          const needsQuoting = Boolean(qualifier) && (value.includes(delimiter) || value.includes(quote) || /[\r\n]/.test(value));

          if (!needsQuoting) return value;

          return `${quote}${value.split(quote).join(quote + quote)}${quote}`;
        })
        .join(delimiter)
    )
    .join("\r\n");
}

function detectColumnType(values: string[]): ColumnType {
  const nonEmpty = values.map(sanitizeCell).filter(Boolean);

  if (!nonEmpty.length) return "text";

  const currencyCount = nonEmpty.filter((value) => /^[$€£]\s?-?[\d,]+(\.\d+)?$/.test(value)).length;
  const percentCount = nonEmpty.filter((value) => /^-?[\d,]+(\.\d+)?\s?%$/.test(value)).length;
  const dateCount = nonEmpty.filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value) || /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(value)).length;
  const numericCount = nonEmpty.filter((value) => /^-?[\d,]+(\.\d+)?$/.test(value.replace(/[$€£%]/g, ""))).length;
  const ratio = (count: number) => count / nonEmpty.length;

  if (ratio(currencyCount) >= 0.7) return "currency";
  if (ratio(percentCount) >= 0.7) return "percent";
  if (ratio(dateCount) >= 0.7) return "date";
  if (ratio(numericCount) >= 0.7) return "numeric";

  return "text";
}

function parseNumericValue(raw: string) {
  const cleaned = raw.replace(/[^0-9.-]/g, "");
  const value = parseFloat(cleaned);

  return Number.isFinite(value) ? value : 0;
}

function parseDateSerial(raw: string) {
  const parsed = Date.parse(raw);

  if (Number.isNaN(parsed)) return null;

  const excelEpoch = Date.UTC(1899, 11, 30);

  return Math.round((parsed - excelEpoch) / 86400000);
}

function computeAggregateValue(aggregate: AggregateType, values: number[]) {
  const finiteValues = values.filter((value) => Number.isFinite(value));

  if (!finiteValues.length) return 0;

  switch (aggregate) {
    case "sum":
      return finiteValues.reduce((total, value) => total + value, 0);
    case "average":
      return finiteValues.reduce((total, value) => total + value, 0) / finiteValues.length;
    case "count":
      return finiteValues.length;
    case "min":
      return Math.min(...finiteValues);
    case "max":
      return finiteValues.reduce((max, value) => Math.max(max, value), -Infinity);
    default:
      return 0;
  }
}

function numericStyleFor(type: ColumnType) {
  if (type === "currency") return STYLE_CURRENCY;
  if (type === "percent") return STYLE_PERCENT;
  if (type === "date") return STYLE_DATE;
  if (type === "numeric") return STYLE_NUMBER;

  return STYLE_DEFAULT;
}

function totalsStyleFor(type: ColumnType) {
  return type === "currency" ? STYLE_TOTAL_CURRENCY : STYLE_TOTAL_NUMBER;
}

function cellXml(ref: string, options: { text?: string; number?: number; formula?: string; style: number }) {
  if (options.formula !== undefined) {
    const cachedValue = options.number !== undefined ? `<v>${options.number}</v>` : "";
    return `<c r="${ref}" s="${options.style}"><f>${escapeXml(options.formula)}</f>${cachedValue}</c>`;
  }

  if (options.number !== undefined) {
    return `<c r="${ref}" s="${options.style}"><v>${options.number}</v></c>`;
  }

  const text = options.text ?? "";

  if (!text) return `<c r="${ref}" s="${options.style}"/>`;

  return `<c r="${ref}" t="inlineStr" s="${options.style}"><is><t>${escapeXml(text)}</t></is></c>`;
}

const STYLED_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/><color rgb="FF0F172A"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF0F172A"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFCBD5E1"/></bottom><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FF94A3B8"/></top><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="10"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="4" fontId="2" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="2" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs></styleSheet>`;

function buildStyledWorkbook(matrix: string[][], options?: { aggregates?: Record<number, AggregateChoice>; sheetName?: string }) {
  const headers = matrix[0] ?? [];
  const dataRows = matrix.slice(1);
  const columnCount = dataRows.reduce((max, row) => Math.max(max, row.length), Math.max(headers.length, 1));
  const columnTypes = Array.from({ length: columnCount }, (_, columnIndex) =>
    detectColumnType(dataRows.map((row) => row[columnIndex] ?? ""))
  );
  const aggregates = options?.aggregates ?? {};
  const sheetName = (options?.sheetName ?? "Sheet1").slice(0, 31);

  const headerCells = Array.from({ length: columnCount }, (_, columnIndex) =>
    cellXml(`${columnIndexToReference(columnIndex)}1`, { text: headers[columnIndex] ?? `Column ${columnIndex + 1}`, style: STYLE_HEADER })
  ).join("");
  const rowsXml = [`<row r="1">${headerCells}</row>`];

  dataRows.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + 2;
    const isBand = rowIndex % 2 === 1;
    const cells = Array.from({ length: columnCount }, (_, columnIndex) => {
      const ref = `${columnIndexToReference(columnIndex)}${rowNumber}`;
      const rawValue = sanitizeCell(row[columnIndex] ?? "");
      const type = columnTypes[columnIndex];

      if (type === "text" || !rawValue) {
        return cellXml(ref, { text: rawValue, style: isBand ? STYLE_BAND : STYLE_DEFAULT });
      }

      if (type === "date") {
        const serial = parseDateSerial(rawValue);
        if (serial === null) return cellXml(ref, { text: rawValue, style: isBand ? STYLE_BAND : STYLE_DEFAULT });
        return cellXml(ref, { number: serial, style: STYLE_DATE });
      }

      const numericValue = parseNumericValue(rawValue) / (type === "percent" ? 100 : 1);
      return cellXml(ref, { number: numericValue, style: numericStyleFor(type) });
    }).join("");

    rowsXml.push(`<row r="${rowNumber}">${cells}</row>`);
  });

  const hasAggregates = Object.values(aggregates).some((value) => value && value !== "none");
  let totalsRowXml = "";

  if (hasAggregates) {
    const totalsRowNumber = dataRows.length + 2;
    const totalCells = Array.from({ length: columnCount }, (_, columnIndex) => {
      const ref = `${columnIndexToReference(columnIndex)}${totalsRowNumber}`;
      const aggregate = aggregates[columnIndex];

      if (!aggregate || aggregate === "none") {
        return columnIndex === 0
          ? cellXml(ref, { text: "Total", style: STYLE_TOTAL_LABEL })
          : cellXml(ref, { text: "", style: STYLE_TOTAL_LABEL });
      }

      const columnLetter = columnIndexToReference(columnIndex);
      const formulaName = { sum: "SUM", average: "AVERAGE", count: "COUNT", min: "MIN", max: "MAX" }[aggregate];
      const formula = `${formulaName}(${columnLetter}2:${columnLetter}${dataRows.length + 1})`;
      const values = dataRows.map((row) => parseNumericValue(sanitizeCell(row[columnIndex] ?? "")));
      const cached = computeAggregateValue(aggregate, values);

      return cellXml(ref, { formula, number: cached, style: totalsStyleFor(columnTypes[columnIndex]) });
    }).join("");

    totalsRowXml = `<row r="${totalsRowNumber}">${totalCells}</row>`;
  }

  const totalRowCount = dataRows.length + 1 + (hasAggregates ? 1 : 0);
  const lastColumnRef = columnIndexToReference(columnCount - 1);
  const dimensionRef = `A1:${lastColumnRef}${totalRowCount}`;
  const columnWidthXml = Array.from({ length: columnCount }, (_, index) => {
    const width = Math.min(Math.max((headers[index] ?? "").length + 4, 12), 40);
    return `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
  }).join("");

  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${dimensionRef}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${columnWidthXml}</cols><sheetData>${rowsXml.join("") + totalsRowXml}</sheetData><autoFilter ref="A1:${lastColumnRef}1"/></worksheet>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><calcPr calcId="0" fullCalcOnLoad="1"/><sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/styles.xml": strToU8(STYLED_STYLES_XML),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
  });
}

function createExcelWorkbook(rows: CleanRow[], headers: string[]) {
  const allRows = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  const sheetRows = allRows
    .map((row, rowIndex) => {
      const rowNumber = rowIndex + 1;
      const cells = row
        .map((value, columnIndex) => {
          const cellReference = `${columnIndexToReference(columnIndex)}${rowNumber}`;
          return `<c r="${cellReference}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
        })
        .join("");

      return `<row r="${rowNumber}">${cells}</row>`;
    })
    .join("");

  const columnWidthXml = headers
    .map((header, index) => `<col min="${index + 1}" max="${index + 1}" width="${Math.min(Math.max(header.length + 3, 14), 34)}" customWidth="1"/>`)
    .join("");
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${columnWidthXml}</cols><sheetData>${sheetRows}</sheetData></worksheet>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Clean Leads" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/styles.xml": strToU8(stylesXml),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
  });
}

function downloadBlob(content: string | Uint8Array, fileName: string, type: string) {
  const blob = new Blob([content as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 250);
}

function isFooterPageKey(value: string): value is FooterPageKey {
  return value in FOOTER_PAGES;
}

function isToolPageKey(value: string): value is ToolPageKey {
  return value in SEO_TOOL_PAGES;
}

let pendingWorkspaceTab: "leads" | null = null;

function getCurrentPage(): AppPageKey {
  const hashRoute = window.location.hash.startsWith("#/") ? window.location.hash.replace("#/", "") : "";
  const pathRoute = window.location.pathname.replace(/^\//, "").replace(/\/$/, "");
  const route = hashRoute || (pathRoute && pathRoute !== "index.html" ? pathRoute : "");

  // Deep links to the workspace: open the home workspace on the cleaner tab.
  const WORKSPACE_SLUGS: Record<string, "leads"> = {
    "data-cleaner": "leads",
    "cleaner": "leads",
    "quick-data-csv-cleaner": "leads",
    "csv-studio": "leads",
  };
  if (WORKSPACE_SLUGS[route]) {
    pendingWorkspaceTab = WORKSPACE_SLUGS[route];
    if (typeof window !== "undefined") window.history.replaceState({}, "", "/");
    return "home";
  }

  // Legacy slug redirects (301-style client-side): map old URLs to new ones.
  const LEGACY: Record<string, AppPageKey> = {
    "sort-subtotal-automation": "excel-automation",
    "client-funds": "excel-automation",
    "bank-ledger-x": "reconciliation-hub",
  };
  if (LEGACY[route]) {
    const target = LEGACY[route];
    if (typeof window !== "undefined") {
      const newPath = target === "home" ? "/" : `/${target}`;
      window.history.replaceState({}, "", newPath);
    }
    return target;
  }

  if (route === "data-engine") return "not-found";
  if (route === "reconciliation-hub") return "reconciliation-hub";
  if (route === "data-toolbox") return "data-toolbox";
  if (route === "excel-automation") return "excel-automation";
  if (route === "excel-academy" || route === "free-excel-academy") return "excel-academy";
  if (isFooterPageKey(route) || isToolPageKey(route)) return route;

  return "not-found";
}

function updateMetaTag(selector: string, attribute: "content" | "href", value: string) {
  const element = document.head.querySelector(selector);

  if (element) element.setAttribute(attribute, value);
}

function updatePageMetadata(page: AppPageKey) {
  const toolPage = isToolPageKey(page) ? SEO_TOOL_PAGES[page] : null;
  const footerPage = isFooterPageKey(page) ? FOOTER_PAGES[page] : null;
  const isHub = page === "reconciliation-hub";
  const isToolbox = page === "data-toolbox";
  const isExcelAuto = page === "excel-automation";
  const isAcademy = page === "excel-academy";
  const title = isHub
    ? "Reconciliation Hub | Data Matching, Verification and Exception Reporting"
    : isToolbox
    ? "Data Toolbox | Excel, CSV, Validation and Data Quality Tools"
    : isExcelAuto
    ? "Excel Automation | Sorting, Subtotals, Formulas and Spreadsheet Processing"
    : isAcademy
    ? "Free Excel Academy | Learn Excel Formulas, Functions and Shortcuts"
    : toolPage?.title ?? (footerPage ? `${footerPage.title} | ${PRODUCT_NAME}` : "MarqClean AI | AI Data, Excel, CSV and Reconciliation Automation Platform");
  const description = isHub
    ? "AI-powered reconciliation, record matching and validation platform. Reconcile client, investor, account and banking datasets across Excel, CSV and PDF files, including the Bank Ledger X bank statement and ledger reconciliation module, with exception reporting."
    : isToolbox
    ? "Use advanced Excel, CSV and data utilities for validation, formatting, conversion, cleansing and data quality checks across structured datasets."
    : isAcademy
    ? "A free, structured Excel course built into MarqClean AI: guided lessons across every major topic, a searchable A-Z function reference, and a printable cheat sheet."
    : isExcelAuto
    ? "Automate Excel sorting, subtotaling, number formatting, formula handling and workbook generation for cleaner spreadsheet workflows."
    : toolPage?.description ?? footerPage?.description ?? BASE_META_DESCRIPTION;
  const keywords = isHub
    ? "reconciliation hub, data reconciliation, client reconciliation, investor reconciliation, account matching, record matching, Excel reconciliation, CSV reconciliation, exception reporting, data verification"
    : isToolbox
    ? "data toolbox, data validation tools, data formatting tools, Excel utilities, CSV utilities, data transformation, data quality tools, spreadsheet utilities, file conversion tools, advanced excel functions"
    : isExcelAuto
    ? "Excel automation, Excel sorting automation, Excel subtotal automation, spreadsheet automation, Excel formulas, Excel formatting, workbook automation, Excel data processing, automate Excel files"
    : isAcademy
    ? "learn Excel, free Excel course, Excel tutorial, Excel functions, Excel formulas, Excel shortcuts, Excel training, VLOOKUP tutorial, XLOOKUP tutorial, Excel for beginners, Excel cheat sheet"
    : toolPage?.keyword
    ? `${toolPage.keyword}, AI data automation, Excel automation, CSV processing, data reconciliation, data validation, data quality`
    : "AI data automation, data cleaning software, Excel automation, CSV processing, data reconciliation, spreadsheet automation, data validation, data quality, financial reconciliation, Excel data cleaning, CSV to Excel conversion, data transformation, record matching, exception reporting, PDF to Excel, bank statement reconciliation";
  const canonicalUrl = page === "home" ? `${SITE_URL}/` : `${SITE_URL}/${page}`;
  const robots = page === "not-found" ? "noindex, nofollow" : "index, follow";

  document.title = title;
  updateMetaTag('meta[name="description"]', "content", description);
  updateMetaTag('meta[name="robots"]', "content", robots);
  updateMetaTag('meta[name="keywords"]', "content", keywords);
  updateMetaTag('meta[property="og:title"]', "content", title);
  updateMetaTag('meta[property="og:description"]', "content", description);
  updateMetaTag('meta[property="og:url"]', "content", canonicalUrl);
  updateMetaTag('link[rel="canonical"]', "href", canonicalUrl);
}

function navigateToPage(page: AppPageKey) {
  const path = page === "home" ? "/" : `/${page}`;
  window.history.pushState({}, "", path);
  window.dispatchEvent(new PopStateEvent("popstate"));
}

function PageLink({ page, children, className }: { page: AppPageKey; children: React.ReactNode; className?: string }) {
  return (
    <a
      className={className}
      href={page === "home" ? "/" : `/${page}`}
      onClick={(event) => {
        event.preventDefault();
        navigateToPage(page);
      }}
    >
      {children}
    </a>
  );
}

const TOOL_CATEGORY_ORDER = ["All", "Cleaners", "Formatters", "Converters", "Excel Formulas", "Lookups", "Finance", "Data Modelling", "Reconciliation", "AI Tools"];

function ToolsDirectory() {
  const [query, setQuery] = useState("");
  const [category, setCategory] = useState("All");
  const allTools = useMemo(() => Object.values(SEO_TOOL_PAGES), []);
  const categories = useMemo(() => {
    const used = new Set(allTools.map((t) => t.category));
    return TOOL_CATEGORY_ORDER.filter((c) => c === "All" || used.has(c));
  }, [allTools]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return allTools.filter((tool) => {
      if (category !== "All" && tool.category !== category) return false;
      if (!q) return true;
      return (
        tool.cardTitle.toLowerCase().includes(q) ||
        tool.keyword.toLowerCase().includes(q) ||
        tool.category.toLowerCase().includes(q)
      );
    });
  }, [allTools, query, category]);

  return (
    <div>
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div className="relative w-full sm:max-w-sm">
          <svg className="pointer-events-none absolute left-3.5 top-1/2 h-4 w-4 -translate-y-1/2 text-ink-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <circle cx="11" cy="11" r="7" /><path d="m21 21-4.3-4.3" />
          </svg>
          <input
            type="search"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search data and Excel tools..."
            aria-label="Search data and Excel tools"
            className="w-full rounded-md border border-line bg-canvas py-2.5 pl-10 pr-4 text-sm text-ink placeholder:text-ink-2 outline-none transition focus-visible:border-accent focus-visible:ring-2 focus-visible:ring-accent/40"
          />
        </div>
        <p className="text-xs text-ink-2">{filtered.length} of {allTools.length} tools</p>
      </div>

      <div className="mt-4 flex flex-wrap gap-2" role="group" aria-label="Filter tools by category">
        {categories.map((c) => (
          <button
            key={c}
            type="button"
            aria-pressed={category === c}
            onClick={() => setCategory(c)}
            className={`rounded-md border px-3.5 py-1.5 text-xs font-medium transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50 ${
              category === c
                ? "border-[rgb(var(--accent))] bg-[rgb(var(--accent))]/20 text-ink"
                : "border-line text-ink-3 hover:border-line hover:text-ink"
            }`}
          >
            {c}
          </button>
        ))}
      </div>

      {filtered.length === 0 ? (
        <div className="mt-10 rounded-lg border border-dashed border-line bg-canvas p-10 text-center">
          <p className="text-sm font-medium text-ink">No tools match "{query}"{category !== "All" ? ` in ${category}` : ""}.</p>
          <p className="mt-2 text-sm text-ink-3">Try a different search term, or clear the category filter.</p>
          <button
            onClick={() => { setQuery(""); setCategory("All"); }}
            className="mt-4 rounded-md border border-line px-4 py-2 text-xs font-medium text-ink hover:border-line"
          >
            Clear search and filters
          </button>
        </div>
      ) : (
        <div className="mt-8 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {filtered.map((tool) => (
            <PageLink
              key={tool.slug}
              page={tool.slug}
              className="group mc-card-interactive flex flex-col rounded-lg border border-line bg-canvas p-5 text-sm hover:border-accent hover:bg-surface  focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50"
            >
              <div className="flex items-center justify-between gap-2">
                <span className="text-2xl" aria-hidden="true">{tool.icon}</span>
                <span className="rounded-md border border-line bg-surface px-2 py-0.5 text-[10px] font-medium uppercase tracking-wide text-ink-2">Free Tool</span>
              </div>
              <span className="mc-mono mt-3 block text-[10px] uppercase tracking-[0.18em] text-[rgb(var(--accent))]/80">{tool.category}</span>
              <span className="mt-1.5 block text-lg font-medium leading-snug tracking-[-0.01em] text-ink">{tool.cardTitle}</span>
              <span className="mt-2 flex-1 text-ink-3">{tool.description.replace(/^Free /, "").slice(0, 96)}...</span>
              <span className="mt-4 inline-flex items-center gap-1.5 text-sm font-medium text-[rgb(var(--accent-hover))] transition group-hover:gap-2.5 group-hover:text-ink">
                Open Tool
                <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M5 12h14" /><path d="m12 5 7 7-7 7" /></svg>
              </span>
            </PageLink>
          ))}
        </div>
      )}
    </div>
  );
}

function FooterContentPage({ page }: { page: FooterPage }) {
  return (
    <main id="main-content" className="min-h-screen bg-[rgb(var(--canvas))] px-5 pb-20 pt-32 text-ink lg:px-8">
      <motion.article
        className="mx-auto max-w-4xl"
        initial={{ opacity: 0, y: 18 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, ease: "easeOut" }}
      >
        <PageLink className="text-sm font-medium text-[rgb(var(--accent-hover))] transition hover:text-ink" page="home">
          Back to product
        </PageLink>
        <p className="mt-8 text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">{COMPANY_NAME}</p>
        <h1 className="mt-4 text-5xl font-medium tracking-[-0.05em] text-ink sm:text-6xl">{page.title}</h1>
        <p className="mt-6 text-lg leading-8 text-ink-2">{page.description}</p>
        <p className="mt-4 text-sm text-ink-2">Last updated: June 29, 2026</p>

        <div className="mt-12 space-y-10">
          {page.sections.map((section) => (
            <section key={section.heading} className="border-t border-line pt-8">
              <h2 className="text-3xl font-medium tracking-[-0.04em] text-ink">{section.heading}</h2>
              <div className="mt-4 space-y-4">
                {section.body.map((paragraph) => (
                  <p className="text-justify leading-7 text-ink-2 [hyphens:auto]" key={paragraph}>
                    {paragraph}
                  </p>
                ))}
              </div>
            </section>
          ))}
        </div>
      </motion.article>
    </main>
  );
}

function ToolContentPage({ page }: { page: SeoToolPage }) {
  return (
    <main id="main-content" className="min-h-screen bg-[rgb(var(--canvas))] px-5 pb-20 pt-32 text-ink lg:px-8">
      <motion.article
        className="mx-auto max-w-5xl"
        initial={{ opacity: 0, y: 18 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, ease: "easeOut" }}
      >
        <PageLink className="text-sm font-medium text-[rgb(var(--accent-hover))] transition hover:text-ink" page="home">
          Back to free CSV cleaner
        </PageLink>
        <p className="mt-8 text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Free browser-based tool</p>
        <h1 className="mt-4 max-w-4xl text-5xl font-medium tracking-[-0.05em] text-ink sm:text-6xl">{page.h1}</h1>
        <p className="mt-6 max-w-3xl text-lg leading-8 text-ink-2">{page.description}</p>
        <div className="mt-8 flex flex-col gap-3 sm:flex-row">
          <a
            className="mc-cta-cyan rounded-md px-6 py-3 text-center text-sm font-medium"
            href="/#cleaner"
            onClick={(event) => {
              event.preventDefault();
              navigateToPage("home");
              window.setTimeout(() => document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" }), 80);
            }}
          >
            Use the free cleaner
          </a>
          <PageLink className="rounded-md border border-line px-6 py-3 text-center text-sm font-medium text-ink transition hover:border-line" page="lead-list-cleaner">
            Explore lead list cleanup
          </PageLink>
        </div>

        <section className="mt-16 border-t border-line pt-10">
          <p className="text-sm font-medium uppercase tracking-[0.24em] text-[rgb(var(--accent))]">{page.keyword}</p>
          <h2 className="mt-4 text-4xl font-medium tracking-[-0.04em] text-ink">{page.h2}</h2>
          <p className="mt-5 max-w-3xl text-justify leading-7 text-ink-2 [hyphens:auto]">
            This page is powered by the same modular MarqClean AI cleaning engine used across every free CSV cleaner,
            Excel converter, lead list cleaner, CRM data cleanup tool, duplicate checker, and spreadsheet formatting workflow.
          </p>
        </section>

        <div className="mt-12 grid gap-6 md:grid-cols-3">
          {page.sections.map((section) => (
            <section key={section.h3} className="mc-glass rounded-lg p-6">
              <h3 className="text-xl font-medium tracking-[-0.02em] text-ink">{section.h3}</h3>
              <p className="mt-3 text-[15px] leading-6 text-ink-2">{section.body}</p>
            </section>
          ))}
        </div>
      </motion.article>
    </main>
  );
}

type WorkspaceSideWidgetProps = {
  tab: "leads" | "converter" | "formulas" | "bank";
  processing?: boolean;
  hasResult?: boolean;
};

function WorkspaceSideWidget({ tab, processing = false, hasResult = false }: WorkspaceSideWidgetProps) {
  const config = {
    leads: {
      eyebrow: "Cleaner workflow",
      title: "Clean → review → export",
      description: "A focused workspace for preparing structured lead and spreadsheet data.",
      steps: ["Choose CSV/XLSX", "Review detected fixes", "Export clean data"],
      tip: "Use Sample CSV first if you want to see the complete workflow before uploading your own file.",
    },
    converter: {
      eyebrow: "Conversion workflow",
      title: "Split → preview → export",
      description: "Keep the conversion settings and output format visible while you work.",
      steps: ["Upload CSV/TXT/XLSX", "Confirm delimiter rules", "Export Excel or text"],
      tip: "Match the delimiter and text qualifier to the source file before exporting.",
    },
    formulas: {
      eyebrow: "Excel workflow",
      title: "Inspect → format → calculate",
      description: "Detect column types, choose calculations, then create a structured workbook.",
      steps: ["Upload CSV/XLSX", "Review detected types", "Export formatted workbook"],
      tip: "Use the calculation selector only on columns where a summary is meaningful.",
    },
    bank: {
      eyebrow: "Ledger workflow",
      title: "Extract → verify → export",
      description: "Turn statement data into structured transactions while keeping review steps visible.",
      steps: ["Upload or paste statement", "Review transactions", "Export Excel, CSV or QIF"],
      tip: "Always review extracted values before using them for accounting or reconciliation.",
    },
  }[tab];

  const currentStep = processing ? 1 : hasResult ? 3 : 1;

  return (
    <aside className="ws-side-widget" aria-label="Workspace guidance">
      <div className="ws-side-widget__header">
        <span className="ws-side-widget__eyebrow">{config.eyebrow}</span>
        <span className={`ws-side-widget__status ${processing ? "is-processing" : hasResult ? "is-complete" : ""}`}>
          {processing ? "Working" : hasResult ? "Ready" : "Ready"}
        </span>
      </div>
      <h3>{config.title}</h3>
      <p className="ws-side-widget__description">{config.description}</p>

      <div className="ws-side-widget__section">
        <div className="ws-side-widget__section-title">Workflow</div>
        <ol className="ws-side-widget__steps">
          {config.steps.map((step, index) => (
            <li key={step} className={index + 1 <= currentStep ? "is-current" : ""}>
              <span>{index + 1}</span>
              <strong>{step}</strong>
            </li>
          ))}
        </ol>
      </div>

      <div className="ws-side-widget__section">
        <div className="ws-side-widget__section-title">Workspace tip</div>
        <p className="ws-side-widget__tip">{config.tip}</p>
      </div>

      <div className="ws-side-widget__privacy">
        <span aria-hidden="true">✓</span>
        <div>
          <strong>Browser-local processing</strong>
          <small>Your files stay on this device while the workspace runs.</small>
        </div>
      </div>
    </aside>
  );
}

export default function App() {
  const [result, setResult] = useState<CleanResult | null>(null);
  const [error, setError] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [currentPage, setCurrentPage] = useState<AppPageKey>("home");
  const [workspaceTab, setWorkspaceTab] = useState<"leads" | "converter" | "formulas" | "bank">("leads");
  const [isMobileNavOpen, setIsMobileNavOpen] = useState(false);
  const [isWorkspacesOpen, setIsWorkspacesOpen] = useState(false);
  const [isMobileWorkspacesOpen, setIsMobileWorkspacesOpen] = useState(false);
  const productsMenuRef = useRef<HTMLDivElement>(null);
  const [showBackToTop, setShowBackToTop] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const workspaceRef = useRef<HTMLElement>(null);

  const [converterFileName, setConverterFileName] = useState("");
  const [converterMode, setConverterMode] = useState<"text-to-columns" | "excel-to-text" | null>(null);
  const [converterRawText, setConverterRawText] = useState("");
  const [converterMatrix, setConverterMatrix] = useState<string[][] | null>(null);
  const [wizardSettings, setWizardSettings] = useState<WizardSettings>(DEFAULT_WIZARD_SETTINGS);
  const [converterError, setConverterError] = useState("");
  const [converterProcessing, setConverterProcessing] = useState(false);
  const converterInputRef = useRef<HTMLInputElement>(null);

  const [formulaFileName, setFormulaFileName] = useState("");
  const [formulaMatrix, setFormulaMatrix] = useState<string[][] | null>(null);
  const [formulaColumnTypes, setFormulaColumnTypes] = useState<ColumnType[]>([]);
  const [aggregateChoices, setAggregateChoices] = useState<Record<number, AggregateChoice>>({});
  const [formulaError, setFormulaError] = useState("");
  const [formulaProcessing, setFormulaProcessing] = useState(false);
  const formulaInputRef = useRef<HTMLInputElement>(null);

  const [bankFileName, setBankFileName] = useState("");
  const [bankTransactions, setBankTransactions] = useState<BankTransaction[] | null>(null);
  const [bankPastedText, setBankPastedText] = useState("");
  const [bankError, setBankError] = useState("");
  const [bankProcessing, setBankProcessing] = useState(false);
  const bankInputRef = useRef<HTMLInputElement>(null);

  const previewRows = useMemo(() => result?.rows.slice(0, 6) ?? [], [result]);
  const visibleHeaders = useMemo(() => result?.headers.slice(0, 9) ?? [], [result]);
  const activeFooterPage = isFooterPageKey(currentPage) ? FOOTER_PAGES[currentPage] : null;
  const activeToolPage = isToolPageKey(currentPage) ? SEO_TOOL_PAGES[currentPage] : null;

  const converterPreviewMatrix = useMemo(() => {
    if (converterMode === "text-to-columns") return parseWithWizardSettings(converterRawText, wizardSettings);
    if (converterMode === "excel-to-text") return converterMatrix ?? [];
    return [];
  }, [converterMode, converterRawText, wizardSettings, converterMatrix]);

  useEffect(() => {
    const syncPageFromHash = () => {
      setCurrentPage(getCurrentPage());
    };

    const openBank = () => setWorkspaceTab("bank");

    if (pendingWorkspaceTab) {
      setWorkspaceTab(pendingWorkspaceTab);
      const target = pendingWorkspaceTab;
      pendingWorkspaceTab = null;
      window.setTimeout(() => {
        if (target === "leads") {
          workspaceRef.current?.scrollIntoView({ behavior: "smooth" });
        }
      }, 300);
    }

    syncPageFromHash();
    window.addEventListener("hashchange", syncPageFromHash);
    window.addEventListener("popstate", syncPageFromHash);
    window.addEventListener("marqclean:open-bank", openBank);

    return () => {
      window.removeEventListener("hashchange", syncPageFromHash);
      window.removeEventListener("popstate", syncPageFromHash);
      window.removeEventListener("marqclean:open-bank", openBank);
    };
  }, []);

  useEffect(() => {
    if (!isWorkspacesOpen) return;
    function handleOutside(event: MouseEvent) {
      if (productsMenuRef.current && !productsMenuRef.current.contains(event.target as Node)) {
        setIsWorkspacesOpen(false);
      }
    }
    function handleKey(event: KeyboardEvent) {
      if (event.key === "Escape") setIsWorkspacesOpen(false);
    }
    document.addEventListener("mousedown", handleOutside);
    document.addEventListener("keydown", handleKey);
    return () => {
      document.removeEventListener("mousedown", handleOutside);
      document.removeEventListener("keydown", handleKey);
    };
  }, [isWorkspacesOpen]);

  useEffect(() => {
    let ticking = false;
    const onScroll = () => {
      if (ticking) return;
      ticking = true;
      window.requestAnimationFrame(() => {
        setShowBackToTop(window.scrollY > 480);
        ticking = false;
      });
    };
    window.addEventListener("scroll", onScroll, { passive: true });
    onScroll();
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  useEffect(() => {
    updatePageMetadata(currentPage);
    setIsMobileNavOpen(false);
    setIsWorkspacesOpen(false);
    setIsMobileWorkspacesOpen(false);
  }, [currentPage]);

  useEffect(() => {
    if (currentPage !== "home") {
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    const targetId = window.location.hash.replace("#", "");
    if (targetId && !targetId.startsWith("/")) {
      window.setTimeout(() => document.getElementById(targetId)?.scrollIntoView({ behavior: "smooth" }), 60);
    } else {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }, [currentPage]);

  async function handleFile(file: File, pipelineActions: SmartPipelineAction[] = []) {
    setError("");
    setIsProcessing(true);

    try {
      const fileExtension = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";
      const mimeExtension = file.type.includes("csv") ? "csv" : file.type.includes("spreadsheet") ? "xlsx" : "";
      const extension = fileExtension || mimeExtension;

      if (!extension || !SUPPORTED_EXTENSIONS.has(extension)) {
        throw new Error("Unsupported file type. Upload a CSV or XLSX file.");
      }

      if (file.size > MAX_FILE_SIZE_BYTES) {
        throw new Error(`This file is larger than ${MAX_FILE_SIZE_LABEL}. Try a smaller export or split the file first.`);
      }

      const rawRows = extension === "xlsx" ? await parseExcelFile(file) : await parseCsvFile(file);
      const usableRows = rawRows.filter((row) => !isBlankRawRow(row));

      if (!usableRows.length) {
        throw new Error("The file did not contain readable rows. Try a CSV or Excel workbook with a header row.");
      }

      let cleaned = cleanDataset(usableRows, file.name);

      if (pipelineActions.some((action) => action.id === "trim")) {
        cleaned = {
          ...cleaned,
          rows: cleaned.rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, value.trim()]))),
        };
      }

      if (pipelineActions.some((action) => action.id === "group-reference")) {
        const referenceHeader = cleaned.headers.find((header) => /reference|transaction|account|client|customer/i.test(header));
        if (referenceHeader) {
          const grouped = [...cleaned.rows].sort((a, b) => String(a[referenceHeader] ?? "").localeCompare(String(b[referenceHeader] ?? "")));
          cleaned = { ...cleaned, rows: grouped };
        }
      }

      if (pipelineActions.some((action) => action.id === "subtotals")) {
        const amountHeader = cleaned.headers.find((header) => /amount|value|total|debit|credit/i.test(header));
        if (amountHeader) {
          const numericTotal = cleaned.rows.reduce((sum, row) => {
            const value = Number(String(row[amountHeader] ?? "").replace(/[^0-9.-]/g, ""));
            return Number.isFinite(value) ? sum + value : sum;
          }, 0);
          cleaned = {
            ...cleaned,
            rows: [
              ...cleaned.rows,
              Object.fromEntries(cleaned.headers.map((header, index) => [header, index === 0 ? "Pipeline subtotal" : header === amountHeader ? numericTotal.toFixed(2) : ""])),
            ],
          };
        }
      }

      setResult(cleaned);
      if (pipelineActions.some((action) => action.id === "export-xlsx")) {
        window.setTimeout(() => {
          const workbook = createExcelWorkbook(cleaned.rows, cleaned.headers);
          downloadBlob(workbook, createDownloadName(cleaned.fileName, "xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }, 0);
      }
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "The file could not be cleaned.");
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  function handleSmartRecommendation(recommendation: SmartRecommendation) {
    if (recommendation.id === "bank") {
      setWorkspaceTab("bank");
      window.setTimeout(() => document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" }), 50);
      return;
    }
    if (recommendation.id === "reconciliation") {
      navigateToPage("reconciliation-hub");
      return;
    }
    if (recommendation.id === "cleaner") {
      setWorkspaceTab("leads");
      window.setTimeout(() => document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" }), 50);
      return;
    }
    if (recommendation.id === "formulas") {
      setWorkspaceTab("formulas");
      window.setTimeout(() => document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" }), 50);
      return;
    }
    navigateToPage("data-toolbox");
  }

  function handleFileInput(files: FileList | null, pipelineActions: SmartPipelineAction[] = []) {
    const file = files?.[0];

    if (file) void handleFile(file, pipelineActions);
  }

  function loadSampleData() {
    setResult(cleanDataset(SAMPLE_ROWS, "sample-marketing-leads.csv"));
    setError("");
    workspaceRef.current?.scrollIntoView({ behavior: "smooth" });
  }

  function downloadCsv() {
    if (!result) return;

    const csv = Papa.unparse(result.rows, { columns: result.headers });
    downloadBlob(csv, createDownloadName(result.fileName, "csv"), "text/csv;charset=utf-8;");
  }

  function downloadXlsx() {
    if (!result) return;

    const workbook = createExcelWorkbook(result.rows, result.headers);
    downloadBlob(
      workbook,
      createDownloadName(result.fileName, "xlsx"),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
  }

  function navigateToHomeSection(sectionId: "top" | "cleaner" | "features" | "workflow") {
    navigateToPage("home");
    window.setTimeout(() => document.getElementById(sectionId)?.scrollIntoView({ behavior: "smooth" }), 80);
  }

  function openWorkspaceTab(tab: "leads" | "converter" | "formulas" | "bank") {
    const slug = tab === "leads" ? "/data-cleaner" : "/";
    window.history.pushState({}, "", slug);
    setCurrentPage("home");
    setWorkspaceTab(tab);
    window.setTimeout(() => document.getElementById("cleaner")?.scrollIntoView({ behavior: "smooth" }), 80);
  }

  function resetConverterState() {
    setConverterFileName("");
    setConverterMode(null);
    setConverterRawText("");
    setConverterMatrix(null);
    setConverterError("");
  }

  async function handleConverterFile(file: File) {
    setConverterError("");
    setConverterProcessing(true);

    try {
      const extension = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";

      if (!extension || !CONVERTER_SUPPORTED_EXTENSIONS.has(extension)) {
        throw new Error("Unsupported file type. Upload a CSV, TXT, or XLSX file.");
      }

      if (file.size > MAX_FILE_SIZE_BYTES) {
        throw new Error(`This file is larger than ${MAX_FILE_SIZE_LABEL}. Try a smaller export or split the file first.`);
      }

      if (extension === "xlsx") {
        const matrix = await readWorkbookMatrix(file);
        setConverterMatrix(matrix);
        setConverterMode("excel-to-text");
        setConverterRawText("");
      } else {
        const text = await file.text();
        setConverterRawText(text);
        setConverterMode("text-to-columns");
        setConverterMatrix(null);
      }

      setConverterFileName(file.name);
    } catch (caughtError) {
      resetConverterState();
      setConverterError(caughtError instanceof Error ? caughtError.message : "The file could not be loaded.");
    } finally {
      setConverterProcessing(false);
      if (converterInputRef.current) converterInputRef.current.value = "";
    }
  }

  function handleConverterFileInput(files: FileList | null) {
    const file = files?.[0];
    if (file) void handleConverterFile(file);
  }

  function toggleWizardDelimiter(option: DelimiterOption) {
    setWizardSettings((previous) => ({
      ...previous,
      delimiters: { ...previous.delimiters, [option]: !previous.delimiters[option] },
    }));
  }

  function downloadConverterExcel() {
    if (!converterFileName) return;

    const matrix = converterMode === "excel-to-text" ? converterMatrix ?? [] : parseWithWizardSettings(converterRawText, wizardSettings);
    const workbook = buildStyledWorkbook(matrix, { sheetName: "Converted Data" });

    downloadBlob(workbook, createConvertedFileName(converterFileName, "xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  function downloadConverterText() {
    if (!converterFileName) return;

    const outputChars = getActiveDelimiterChars(wizardSettings);
    const delimiter = outputChars[0] ?? ",";
    const matrix = converterMode === "excel-to-text" ? converterMatrix ?? [] : parseWithWizardSettings(converterRawText, wizardSettings);
    const text = matrixToDelimitedText(matrix, delimiter, wizardSettings.textQualifier);

    downloadBlob(text, createConvertedFileName(converterFileName, "csv"), "text/csv;charset=utf-8;");
  }

  async function handleFormulaFile(file: File) {
    setFormulaError("");
    setFormulaProcessing(true);

    try {
      const extension = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";

      if (!extension || !SUPPORTED_EXTENSIONS.has(extension)) {
        throw new Error("Unsupported file type. Upload a CSV or XLSX file.");
      }

      if (file.size > MAX_FILE_SIZE_BYTES) {
        throw new Error(`This file is larger than ${MAX_FILE_SIZE_LABEL}. Try a smaller export or split the file first.`);
      }

      const matrix = await readFileMatrix(file);

      if (matrix.length < 2) {
        throw new Error("Add a header row and at least one data row before generating formulas.");
      }

      const headers = matrix[0];
      const dataRows = matrix.slice(1);
      const columnTypes = headers.map((_, columnIndex) => detectColumnType(dataRows.map((row) => row[columnIndex] ?? "")));
      const defaultAggregates: Record<number, AggregateChoice> = {};

      columnTypes.forEach((type, columnIndex) => {
        defaultAggregates[columnIndex] = type === "numeric" || type === "currency" || type === "percent" ? "sum" : "none";
      });

      setFormulaMatrix(matrix);
      setFormulaColumnTypes(columnTypes);
      setAggregateChoices(defaultAggregates);
      setFormulaFileName(file.name);
    } catch (caughtError) {
      setFormulaMatrix(null);
      setFormulaColumnTypes([]);
      setAggregateChoices({});
      setFormulaFileName("");
      setFormulaError(caughtError instanceof Error ? caughtError.message : "The file could not be loaded.");
    } finally {
      setFormulaProcessing(false);
      if (formulaInputRef.current) formulaInputRef.current.value = "";
    }
  }

  function handleFormulaFileInput(files: FileList | null) {
    const file = files?.[0];
    if (file) void handleFormulaFile(file);
  }

  function updateAggregateChoice(columnIndex: number, choice: AggregateChoice) {
    setAggregateChoices((previous) => ({ ...previous, [columnIndex]: choice }));
  }

  function downloadFormulaWorkbook() {
    if (!formulaMatrix || !formulaFileName) return;

    const workbook = buildStyledWorkbook(formulaMatrix, { aggregates: aggregateChoices, sheetName: "Formatted Data" });
    downloadBlob(workbook, createConvertedFileName(formulaFileName, "xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }

  async function handleBankFile(file: File) {
    setBankError("");
    setBankProcessing(true);

    try {
      const extension = file.name.includes(".") ? file.name.split(".").pop()?.toLowerCase() : "";

      if (!extension || !["pdf", "txt", "csv"].includes(extension)) {
        throw new Error("Unsupported file type. Upload a PDF or a text based statement (PDF, TXT, or CSV).");
      }

      if (file.size > MAX_FILE_SIZE_BYTES) {
        throw new Error(`This file is larger than ${MAX_FILE_SIZE_LABEL}. Try a smaller statement.`);
      }

      const text = extension === "pdf" ? await extractPdfTextInWorker(file) : await file.text();
      const transactions = extractBankTransactions(text);

      if (!transactions.length) {
        throw new Error("No transactions were detected. Try a clearer statement, or paste the transaction text manually below.");
      }

      setBankFileName(file.name);
      setBankTransactions(transactions);
    } catch (caughtError) {
      setBankTransactions(null);
      setBankError(caughtError instanceof Error ? caughtError.message : "The statement could not be converted.");
    } finally {
      setBankProcessing(false);
      if (bankInputRef.current) bankInputRef.current.value = "";
    }
  }

  function handleBankFileInput(files: FileList | null) {
    const file = files?.[0];
    if (file) void handleBankFile(file);
  }

  function convertPastedStatement() {
    setBankError("");
    const transactions = extractBankTransactions(bankPastedText);

    if (!transactions.length) {
      setBankError("No transactions were detected in the pasted text. Make sure each line includes a date and an amount.");
      setBankTransactions(null);
      return;
    }

    setBankFileName(bankFileName || "pasted-statement.txt");
    setBankTransactions(transactions);
  }

  function downloadBankStatement(format: "xlsx" | "csv" | "qif") {
    if (!bankTransactions) return;
    const baseName = bankFileName || "bank-statement";

    if (format === "xlsx") {
      const workbook = buildStyledWorkbook(bankTransactionsToMatrix(bankTransactions), { sheetName: "Transactions" });
      downloadBlob(workbook, createConvertedFileName(baseName, "xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } else if (format === "csv") {
      const csv = Papa.unparse(bankTransactionsToMatrix(bankTransactions));
      downloadBlob(csv, createConvertedFileName(baseName, "csv"), "text/csv;charset=utf-8;");
    } else {
      downloadBlob(bankTransactionsToQif(bankTransactions), createConvertedFileName(baseName, "qif"), "application/qif");
    }
  }

  function renderBackToTop() {
    return (
      <button
        type="button"
        onClick={() => {
          const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
          window.scrollTo({ top: 0, behavior: reduceMotion ? "auto" : "smooth" });
        }}
        aria-label="Back to top"
        className={`fixed bottom-5 right-4 z-40 inline-flex h-11 w-11 items-center justify-center rounded-md
          border border-line bg-accent text-ink
           transition-all duration-300  
          focus:outline-none focus-visible:ring-2 focus-visible:ring-accent focus-visible:ring-offset-2 focus-visible:ring-offset-canvas
          sm:bottom-7 sm:right-6 sm:h-12 sm:w-12
          ${showBackToTop ? "translate-y-0 scale-100 opacity-100" : "pointer-events-none translate-y-3 scale-90 opacity-0"}`}
      >
        <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2.25" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
          <path d="M12 19V5M5 12l7-7 7 7" />
        </svg>
      </button>
    );
  }

  if (currentPage === "not-found") {
    return (
      <>
        <main id="main-content" className="min-h-screen bg-[rgb(var(--canvas))] px-5 pb-20 pt-40 text-ink lg:px-8">
          <div className="mx-auto max-w-3xl text-center">
            <p className="mc-mono text-sm uppercase tracking-[0.24em] text-[rgb(var(--accent))]">404 · Page not found</p>
            <h1 className="mt-4 text-5xl font-medium tracking-[-0.04em] sm:text-6xl">That page does not exist.</h1>
            <p className="mx-auto mt-5 max-w-xl text-lg leading-8 text-ink-2">The address may be outdated or incorrect. Return to MarqClean AI and choose a workspace or tool.</p>
            <button type="button" className="mt-8 rounded-md bg-accent px-6 py-3 text-sm font-medium text-[rgb(var(--accent-ink))]" onClick={() => navigateToPage("home")}>Return home</button>
          </div>
        </main>
        {renderBackToTop()}
      </>
    );
  }

  if (currentPage === "reconciliation-hub") {
    return (
      <>
        <Suspense fallback={<div className="flex min-h-screen items-center justify-center bg-[rgb(var(--canvas))] text-ink-2">Loading Reconciliation Hub...</div>}>
          <ReconciliationHub onExit={() => navigateToPage("home")} />
        </Suspense>
        {renderBackToTop()}
      </>
    );
  }

  if (currentPage === "data-toolbox") {
    return (
      <>
        <Suspense fallback={<div className="flex min-h-screen items-center justify-center bg-[rgb(var(--canvas))] text-ink-2">Loading Data Toolbox...</div>}>
          <DataToolbox onExit={() => navigateToPage("home")} />
        </Suspense>
        {renderBackToTop()}
      </>
    );
  }

  if (currentPage === "excel-automation") {
    return (
      <>
        <Suspense fallback={<div className="flex min-h-screen items-center justify-center bg-[rgb(var(--surface))] text-[rgb(var(--ink-2))]">Loading Excel Automation...</div>}>
          <ClientFunds onExit={() => navigateToPage("home")} />
        </Suspense>
        {renderBackToTop()}
      </>
    );
  }

  if (currentPage === "excel-academy") {
    return (
      <>
        <Suspense fallback={<div className="flex min-h-screen items-center justify-center bg-[rgb(var(--surface))] text-[rgb(var(--ink-2))]">Loading Free Excel Academy...</div>}>
          <AcademyHub onExit={() => navigateToPage("home")} />
        </Suspense>
        {renderBackToTop()}
      </>
    );
  }

  const productsMenu: Array<{ label: string; href: string; onClick: () => void }> = [
    { label: "Quick Data & CSV Cleaner", href: "/data-cleaner", onClick: () => openWorkspaceTab("leads") },
    { label: "Excel Automation", href: "/excel-automation", onClick: () => navigateToPage("excel-automation") },
    { label: "Reconciliation Hub", href: "/reconciliation-hub", onClick: () => navigateToPage("reconciliation-hub") },
    { label: "Data Toolbox", href: "/data-toolbox", onClick: () => navigateToPage("data-toolbox") },
  ];

  return (
    <div className="amzigo-revamp min-h-screen bg-[rgb(var(--canvas))] text-ink">
      <a
        className="sr-only focus:not-sr-only focus:fixed focus:left-4 focus:top-4 focus:z-50 focus:rounded-md focus:bg-canvas focus:px-4 focus:py-2 focus:text-sm focus:font-medium focus:text-ink"
        href="#main-content"
      >
        Skip to content
      </a>
      <header className="mc-glass fixed left-0 right-0 top-0 z-30 border-b border-line/80 bg-canvas/95 text-ink -xl">
        <nav className="mx-auto flex max-w-[100rem] items-center justify-between gap-4 px-5 py-3.5 lg:px-8" aria-label="Primary navigation">
          <a
            href="/#top"
            className="mc-display flex items-center gap-2 text-lg font-medium tracking-tight"
            onClick={(event) => {
              event.preventDefault();
              navigateToHomeSection("top");
            }}
          >
            <Logo className="shrink-0" />
          </a>
          <div className="hidden items-center gap-6 text-[13px] font-medium text-ink-2 xl:flex">
            <a href="/#top" className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); navigateToHomeSection("top"); }}>Home</a>
            <a href="/#features" className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); navigateToHomeSection("features"); }}>Platform</a>
            <a href="/#workflow" className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); navigateToHomeSection("workflow"); }}>How it works</a>
            <div className="relative" ref={productsMenuRef}>
              <button
                type="button"
                className="flex items-center gap-1 transition hover:text-[rgb(var(--accent))] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50"
                aria-haspopup="true"
                aria-expanded={isWorkspacesOpen}
                aria-controls="workspaces-menu"
                onClick={() => setIsWorkspacesOpen((open) => !open)}
              >Workspaces<svg className={`h-3.5 w-3.5 transition ${isWorkspacesOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
              </button>
              {isWorkspacesOpen ? (
                <div
                  id="workspaces-menu"
                  role="menu"
                  aria-label="Workspaces"
                  className="absolute left-1/2 top-full mt-3 w-72 -translate-x-1/2 rounded-lg border border-line bg-[rgb(var(--surface))] p-2 text-sm"
                >
                  {productsMenu.map((item) => (
                    <a
                      key={item.label}
                      role="menuitem"
                      href={item.href}
                      className="block rounded-lg px-3 py-2.5 font-medium text-ink-2 transition hover:bg-canvas/5 hover:text-ink focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50"
                      onClick={(e) => { e.preventDefault(); setIsWorkspacesOpen(false); item.onClick(); }}
                    >
                      {item.label}
                    </a>
                  ))}
                </div>
              ) : null}
            </div>
            <a href="/contact" className="transition hover:text-[rgb(var(--accent))]" onClick={(e) => { e.preventDefault(); navigateToPage("contact"); }}>Contact</a>
          </div>
          <div className="flex items-center gap-2">
            <ThemeToggle compact />
            <button
              className="mc-cta-cyan hidden rounded-md px-5 py-2 text-sm font-medium sm:inline-flex"
              onClick={() => navigateToHomeSection("cleaner")}
            >
              Start Free
            </button>
            <button
              type="button"
              className="inline-flex h-10 w-10 items-center justify-center rounded-md border border-line text-ink transition hover:border-[rgb(var(--accent))]/60 hover:text-[rgb(var(--accent))] xl:hidden"
              aria-label={isMobileNavOpen ? "Close navigation menu" : "Open navigation menu"}
              aria-expanded={isMobileNavOpen}
              aria-controls="mobile-nav-panel"
              onClick={() => setIsMobileNavOpen((open) => !open)}
            >
              {isMobileNavOpen ? (
                <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
                  <path d="M6 6l12 12M18 6L6 18" />
                </svg>
              ) : (
                <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true">
                  <path d="M4 6h16M4 12h16M4 18h16" />
                </svg>
              )}
            </button>
          </div>
        </nav>
        {isMobileNavOpen ? (
          <div id="mobile-nav-panel" className="border-t border-line bg-canvas px-5 py-4 xl:hidden">
            <div className="flex flex-col gap-1 text-[15px] font-medium text-ink-2">
              <a className="rounded-lg px-3 py-2.5 transition hover:bg-surface hover:text-[rgb(var(--accent))]" href="/#top" onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); navigateToHomeSection("top"); }}>Home</a>
              <a className="rounded-lg px-3 py-2.5 transition hover:bg-canvas/5 hover:text-[rgb(var(--accent))]" href="/#features" onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); navigateToHomeSection("features"); }}>Platform</a>
              <button
                type="button"
                className="flex items-center justify-between rounded-lg px-3 py-2.5 text-left transition hover:bg-surface hover:text-[rgb(var(--accent))]"
                aria-expanded={isMobileWorkspacesOpen}
                aria-controls="mobile-workspaces-menu"
                onClick={() => setIsMobileWorkspacesOpen((open) => !open)}
              >Workspaces<svg className={`h-4 w-4 transition ${isMobileWorkspacesOpen ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6" /></svg>
              </button>
              {isMobileWorkspacesOpen ? (
                <div id="mobile-workspaces-menu" className="ml-3 flex flex-col gap-1 border-l border-line pl-3">
                  {productsMenu.map((item) => (
                    <a
                      key={item.label}
                      href={item.href}
                      className="rounded-lg px-3 py-2 text-sm text-ink-2 transition hover:bg-surface hover:text-[rgb(var(--accent))]"
                      onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); setIsMobileWorkspacesOpen(false); item.onClick(); }}
                    >
                      {item.label}
                    </a>
                  ))}
                </div>
              ) : null}
              <a className="rounded-lg px-3 py-2.5 transition hover:bg-canvas/5 hover:text-[rgb(var(--accent))]" href="/contact" onClick={(e) => { e.preventDefault(); setIsMobileNavOpen(false); navigateToPage("contact"); }}>Contact</a>
              <button
                className="mc-cta-cyan mt-2 rounded-md px-5 py-2.5 text-sm font-medium sm:hidden"
                onClick={() => { setIsMobileNavOpen(false); navigateToHomeSection("cleaner"); }}
              >
                Start Free
              </button>
            </div>
          </div>
        ) : null}
      </header>

      {activeFooterPage ? (
        <FooterContentPage page={activeFooterPage} />
      ) : activeToolPage ? (
        <ToolContentPage page={activeToolPage} />
      ) : (
      <main id="main-content">
        <span id="top" className="sr-only">Top</span>
        <Suspense fallback={<div className="min-h-[80vh] bg-[rgb(var(--canvas))]" />}>
          <HeroCarousel
            onCleanFile={() => {
              navigateToHomeSection("cleaner");
              window.setTimeout(() => fileInputRef.current?.click(), 400);
            }}
            onSampleLeads={() => {
              navigateToHomeSection("cleaner");
              window.setTimeout(() => loadSampleData(), 400);
            }}
            onOpenReconciliation={() => navigateToPage("reconciliation-hub")}
          />
        </Suspense>

        <section className="border-y border-line bg-[rgb(var(--canvas))] py-10 text-ink">
          <div className="mx-auto flex max-w-6xl items-center justify-center gap-2 px-5 pb-6 text-xs text-ink-2 lg:px-8">
            <span className="relative flex h-2 w-2">
              <span className="mc-anim-glow-pulse absolute inline-flex h-full w-full rounded-md bg-[rgb(var(--accent))]" />
              <span className="relative inline-flex h-2 w-2 rounded-md bg-[rgb(var(--accent))]" />
            </span>
            Engine live and processing in your browser right now
          </div>
          <motion.div
            className="mx-auto grid max-w-6xl grid-cols-2 gap-6 px-5 sm:grid-cols-4 lg:px-8"
            initial="hidden"
            whileInView="show"
            viewport={{ once: true, margin: "-40px" }}
            variants={{ hidden: {}, show: { transition: { staggerChildren: 0.1 } } }}
          >
            {[
              { value: "40+", label: "Excel functions built in" },
              { value: "33", label: "Free data & Excel tools" },
              { value: "100%", label: "Runs in your browser" },
              { value: "0", label: "Files ever sent to a server" },
            ].map((stat) => (
              <motion.div
                key={stat.label}
                variants={{ hidden: { opacity: 0, y: 16 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.4, ease: "easeOut" }}
                className="text-center"
              >
                <p className="mc-display text-accent text-3xl font-medium tracking-tight sm:text-4xl">{stat.value}</p>
                <p className="mt-1 text-xs text-ink-3 sm:text-sm">{stat.label}</p>
              </motion.div>
            ))}
          </motion.div>
        </section>

        <section className="bg-canvas px-5 py-14 lg:px-8" aria-labelledby="smart-drop-heading">
          <div id="smart-drop-heading" className="sr-only">Smart Drop data ingestion</div>
          <SmartDropzone onFile={(file, actions) => void handleFile(file, actions)} onRecommendation={handleSmartRecommendation} />
        </section>

        <section ref={workspaceRef} id="cleaner" className="ws-root ws-scope relative">
          <div className="ws-grid pointer-events-none absolute inset-x-0 top-0 h-[30rem]" aria-hidden="true" />
          <div className="relative mx-auto max-w-[100rem] px-5 py-16 lg:px-8">
            <div className="ws-module-header ws-surface mb-5 rounded-lg p-5 sm:p-6">
              <div className="ws-module-header__identity">
                <span className="ws-module-header__icon" aria-hidden="true">{WORKSPACE_MODULES[workspaceTab].icon}</span>
                <div className="min-w-0">
                  <div className="ws-module-header__eyebrow">Workspace module</div>
                  <h2 className="mc-display text-2xl font-medium tracking-tight text-ink sm:text-3xl">{WORKSPACE_MODULES[workspaceTab].title}</h2>
                  <p className="mc-prose mt-1 max-w-3xl text-sm leading-6 text-[rgb(var(--ink-2))] [hyphens:auto]">{WORKSPACE_MODULES[workspaceTab].description}</p>
                </div>
              </div>
              <div className="ws-module-header__formats" aria-label="Supported formats">
                <span className="ws-module-header__formats-label">Supports</span>
                <div className="flex flex-wrap gap-2">
                  {WORKSPACE_MODULES[workspaceTab].formats.map((f) => (
                    <span key={f} className="ws-badge rounded-md px-2.5 py-1 text-[11px] font-medium uppercase tracking-wide">{f}</span>
                  ))}
                </div>
              </div>
            </div>

            <div className="ws-module-tabs">
              <div className="ws-module-tabs__label">Modules available <span>04</span></div>
              <div className="ws-module-tabs__list" role="tablist" aria-label="Data tools">
                {([
                  ["leads", "Quick Data & CSV Cleaner"],
                  ["converter", "CSV to Excel Converter"],
                  ["formulas", "Excel Formulas"],
                  ["bank", "Bank Ledger X"],
                ] as const).map(([key, label]) => (
                  <button
                    key={key}
                    type="button"
                    role="tab"
                    aria-selected={workspaceTab === key}
                    className={`ws-module-tab ${workspaceTab === key ? "is-active" : ""}`}
                    onClick={() => setWorkspaceTab(key)}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>

            <div className="ws-workspace-layout">
              <div className="ws-workspace-main">
          {workspaceTab === "leads" ? (
          <motion.div
            className="ws-cleaner-content grid gap-8 lg:grid-cols-[0.9fr_1.1fr] lg:items-stretch"
            initial={{ opacity: 0, y: 28 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, margin: "-80px" }}
            transition={{ duration: 0.7, ease: "easeOut" }}
          >
            <div className="ws-cleaner-copy flex flex-col justify-center">
              <p className="text-xs font-medium uppercase tracking-[0.22em] text-accent">Automated CSV cleaner</p>
              <h3 className="mt-3 text-[2rem] font-medium leading-tight tracking-[-0.03em] text-ink">
                Upload messy marketing data for free. Download a clean Excel workbook.
              </h3>
              <p className="mt-4 max-w-xl text-[15px] leading-6 text-ink-2">
                {PRODUCT_NAME} runs local browser-based cleaning for CSV and XLSX lead lists, giving marketing teams a
                fast way to standardize outreach data without payment, signup, or repetitive spreadsheet edits.
              </p>
            </div>

            <div className="ws-upload-card rounded-lg border border-line bg-canvas p-3 shadow-hairline sm:p-4" aria-busy={isProcessing}>
              <div
                className="ws-upload-zone rounded-lg border border-dashed border-line bg-surface p-6 text-center transition hover:border-accent hover:bg-accent-tint"
                onDragOver={(event) => event.preventDefault()}
                onDrop={(event) => {
                  event.preventDefault();
                  handleFileInput(event.dataTransfer.files);
                }}
              >
                <input
                  ref={fileInputRef}
                  className="sr-only"
                  type="file"
                  accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  aria-label="Upload CSV or Excel file for cleaning"
                  onChange={(event) => handleFileInput(event.target.files)}
                />
                <div className="ws-upload-icon mx-auto flex h-14 w-14 items-center justify-center rounded-lg bg-accent-tint text-accent">
                  <svg className="h-7 w-7" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
                    <path d="M12 16V4" strokeLinecap="round" />
                    <path d="m7 9 5-5 5 5" strokeLinecap="round" strokeLinejoin="round" />
                    <path d="M20 16.5v1.75A1.75 1.75 0 0 1 18.25 20H5.75A1.75 1.75 0 0 1 4 18.25V16.5" strokeLinecap="round" />
                  </svg>
                </div>
                <h4 className="mt-5 text-2xl font-medium tracking-[-0.03em] text-ink">CSV and Excel upload</h4>
                <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-ink-2">
                  Drop a .csv or .xlsx file here, or choose a file to automatically clean names, companies,
                  industries, emails, phones, and spreadsheet headers. Maximum file size is {MAX_FILE_SIZE_LABEL}.
                </p>
                <div className="mt-6 flex flex-col justify-center gap-3 sm:flex-row">
                  <button
                    className="rounded-md bg-accent px-5 py-3 text-sm font-medium text-[rgb(var(--accent-ink))] transition hover:bg-accent-hover"
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isProcessing}
                  >
                    {isProcessing ? "Processing" : "Choose File"}
                  </button>
                  <button
                    className="rounded-md border border-line px-5 py-3 text-sm font-medium text-ink transition hover:border-ink"
                    onClick={loadSampleData}
                    disabled={isProcessing}
                  >
                    Use Sample CSV
                  </button>
                </div>
              </div>

              <AnimatePresence>
                {isProcessing ? (
                  <motion.div
                    className="mt-5 overflow-hidden rounded-lg bg-surface"
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: "auto" }}
                    exit={{ opacity: 0, height: 0 }}
                  >
                    <div className="p-4 text-sm font-medium text-ink-2" role="status">Cleaning fields and preparing exports...</div>
                    <motion.div
                      className="h-1 bg-accent"
                      initial={{ width: "12%" }}
                      animate={{ width: "100%" }}
                      transition={{ duration: 1.1, repeat: Infinity, repeatType: "reverse" }}
                    />
                  </motion.div>
                ) : null}
              </AnimatePresence>

              {error ? <p className="mt-4 rounded-lg bg-surface p-4 text-sm font-medium text-error" role="alert">{error}</p> : null}

              <AnimatePresence>
                {result ? (
                  <motion.div
                    className="mt-6"
                    initial={{ opacity: 0, y: 18 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: 18 }}
                    transition={{ duration: 0.45 }}
                  >
                    <div className="grid gap-3 sm:grid-cols-3">
                      <div className="rounded-lg bg-ink p-4 text-ink">
                        <p className="text-3xl font-medium">{result.stats.cleanedRows}</p>
                        <p className="mt-1 text-xs uppercase tracking-[0.18em] text-ink-2">Rows cleaned</p>
                      </div>
                      <div className="rounded-lg bg-accent-tint p-4 text-ink">
                        <p className="text-3xl font-medium">{result.stats.fieldRepairs + result.stats.headerFixes}</p>
                        <p className="mt-1 text-xs uppercase tracking-[0.18em] text-ink-2">Fixes applied</p>
                      </div>
                      <div className="rounded-lg bg-surface p-4 text-ink">
                        <p className="text-3xl font-medium">{result.stats.industriesAdded}</p>
                        <p className="mt-1 text-xs uppercase tracking-[0.18em] text-ink-2">Industries added</p>
                      </div>
                    </div>

                    <div className="mt-5 flex flex-col gap-3 sm:flex-row">
                      <motion.button
                        className="rounded-md bg-accent px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint"
                        onClick={downloadXlsx}
                        whileTap={{ scale: 0.98 }}
                      >
                        Download Clean Excel
                      </motion.button>
                      <motion.button
                        className="rounded-md border border-line px-5 py-3 text-sm font-medium text-ink transition hover:border-ink"
                        onClick={downloadCsv}
                        whileTap={{ scale: 0.98 }}
                      >
                        Download Clean CSV
                      </motion.button>
                    </div>

                    <div className="mt-6 overflow-hidden rounded-lg border border-line">
                      <div className="flex items-center justify-between border-b border-line bg-surface px-4 py-3">
                        <h4 className="text-sm font-medium text-ink">Cleaned lead preview</h4>
                        <p className="text-xs text-ink-3">Showing {previewRows.length} rows</p>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="min-w-full text-left text-sm">
                          <thead className="bg-canvas text-xs uppercase tracking-[0.12em] text-ink-3">
                            <tr>
                              {visibleHeaders.map((header) => (
                                <th className="whitespace-nowrap px-4 py-3 font-medium" key={header}>
                                  {header}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-line bg-canvas">
                            {previewRows.map((row, rowIndex) => (
                              <motion.tr
                                key={`${row.Email}-${rowIndex}`}
                                initial={{ opacity: 0, y: 8 }}
                                animate={{ opacity: 1, y: 0 }}
                                transition={{ delay: rowIndex * 0.04 }}
                              >
                                {visibleHeaders.map((header) => (
                                  <td className="max-w-64 truncate px-4 py-3 text-ink-2" key={header}>
                                    {row[header] || "-"}
                                  </td>
                                ))}
                              </motion.tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    <p className="mt-4 text-xs leading-5 text-ink-3">
                      Source headers detected: {result.originalHeaders.join(", ") || "None"}. Files are processed in your
                      browser and are not uploaded to a server.
                    </p>
                  </motion.div>
                ) : null}
              </AnimatePresence>
            </div>
          </motion.div>
          ) : workspaceTab === "converter" ? (
            <motion.div
              className="grid gap-10 lg:grid-cols-[0.82fr_1.18fr] lg:items-start"
              initial={{ opacity: 0, y: 28 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, ease: "easeOut" }}
            >
              <div>
                <p className="text-sm font-medium uppercase tracking-[0.28em] text-accent">CSV to Excel converter</p>
                <h3 className="mt-4 text-3xl font-medium tracking-[-0.03em] text-ink">
                  Split delimited text into columns, just like the Excel wizard.
                </h3>
                <p className="mt-5 text-lg leading-8 text-ink-2">
                  Upload a CSV or TXT file, choose the same delimiter and text qualifier options as the Excel Data tab
                  Text to Columns wizard, preview the split result, and export a properly formatted Excel workbook.
                  Upload an XLSX file instead to reverse the conversion into delimited CSV or TXT.
                </p>
                <p className="mt-4 text-sm text-ink-3">Maximum file size is {MAX_FILE_SIZE_LABEL}.</p>
              </div>

              <div className="rounded-lg border border-line bg-canvas p-4 shadow-hairline sm:p-6" aria-busy={converterProcessing}>
                <div
                  className="rounded-lg border border-dashed border-line bg-surface p-6 text-center transition hover:border-accent hover:bg-accent-tint"
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={(event) => {
                    event.preventDefault();
                    handleConverterFileInput(event.dataTransfer.files);
                  }}
                >
                  <input
                    ref={converterInputRef}
                    className="sr-only"
                    type="file"
                    accept=".csv,.txt,.xlsx,text/csv,text/plain,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    aria-label="Upload CSV, TXT, or Excel file to convert"
                    onChange={(event) => handleConverterFileInput(event.target.files)}
                  />
                  <h4 className="text-2xl font-medium tracking-[-0.03em] text-ink">CSV, TXT, or Excel upload</h4>
                  <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-ink-2">
                    Drop a delimited text file to split it into columns, or drop an XLSX file to convert it back into
                    delimited text.
                  </p>
                  <div className="mt-6 flex flex-col justify-center gap-3 sm:flex-row">
                    <button
                      className="rounded-md bg-ink px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent"
                      onClick={() => converterInputRef.current?.click()}
                      disabled={converterProcessing}
                    >
                      {converterProcessing ? "Loading" : "Choose File"}
                    </button>
                  </div>
                </div>

                {converterError ? <p className="mt-4 rounded-lg bg-surface p-4 text-sm font-medium text-error" role="alert">{converterError}</p> : null}

                {converterMode ? (
                  <div className="mt-6">
                    <h4 className="text-sm font-medium uppercase tracking-[0.18em] text-ink-3">Delimiters</h4>
                    <div className="mt-3 flex flex-wrap gap-4 text-sm text-ink-2">
                      {(["tab", "semicolon", "comma", "space", "other"] as DelimiterOption[]).map((option) => (
                        <label className="flex items-center gap-2" key={option}>
                          <input
                            type="checkbox"
                            checked={wizardSettings.delimiters[option]}
                            onChange={() => toggleWizardDelimiter(option)}
                          />
                          <span className="capitalize">{option}</span>
                        </label>
                      ))}
                      {wizardSettings.delimiters.other ? (
                        <input
                          className="w-16 rounded border border-line px-2 py-1 text-sm"
                          maxLength={1}
                          value={wizardSettings.otherChar}
                          onChange={(event) => setWizardSettings((previous) => ({ ...previous, otherChar: event.target.value }))}
                          aria-label="Other delimiter character"
                        />
                      ) : null}
                    </div>

                    <div className="mt-4 flex flex-wrap items-center gap-6 text-sm text-ink-2">
                      <label className="flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={wizardSettings.treatConsecutiveAsOne}
                          onChange={() => setWizardSettings((previous) => ({ ...previous, treatConsecutiveAsOne: !previous.treatConsecutiveAsOne }))}
                        />
                        Treat consecutive delimiters as one
                      </label>
                      <label className="flex items-center gap-2">
                        Text qualifier
                        <select
                          className="rounded border border-line px-2 py-1"
                          value={wizardSettings.textQualifier}
                          onChange={(event) => setWizardSettings((previous) => ({ ...previous, textQualifier: event.target.value as TextQualifier }))}
                        >
                          <option value="&quot;">&quot;</option>
                          <option value="'">&apos;</option>
                          <option value="">None</option>
                        </select>
                      </label>
                    </div>

                    <div className="mt-6 overflow-hidden rounded-lg border border-line">
                      <div className="border-b border-line bg-surface px-4 py-3">
                        <h4 className="text-sm font-medium text-ink">Data preview</h4>
                      </div>
                      <div className="max-h-72 overflow-auto">
                        <table className="min-w-full text-left text-sm">
                          <tbody className="divide-y divide-line bg-canvas">
                            {converterPreviewMatrix.slice(0, 8).map((row, rowIndex) => (
                              <tr key={rowIndex}>
                                {row.map((value, columnIndex) => (
                                  <td className="max-w-56 truncate px-4 py-2 text-ink-2" key={columnIndex}>
                                    {value || "-"}
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    <div className="mt-5 flex flex-col gap-3 sm:flex-row">
                      <motion.button
                        className="rounded-md bg-accent px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint"
                        onClick={downloadConverterExcel}
                        whileTap={{ scale: 0.98 }}
                      >
                        Download as Excel
                      </motion.button>
                      <motion.button
                        className="rounded-md border border-line px-5 py-3 text-sm font-medium text-ink transition hover:border-ink"
                        onClick={downloadConverterText}
                        whileTap={{ scale: 0.98 }}
                      >
                        {converterMode === "excel-to-text" ? "Download as CSV or TXT" : "Download as Normalized CSV"}
                      </motion.button>
                    </div>
                    <p className="mt-4 text-xs leading-5 text-ink-3">
                      Files are processed in your browser and are not uploaded to a server.
                    </p>
                  </div>
                ) : null}
              </div>
            </motion.div>
          ) : workspaceTab === "formulas" ? (
            <motion.div
              className="grid gap-10 lg:grid-cols-[0.82fr_1.18fr] lg:items-start"
              initial={{ opacity: 0, y: 28 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, ease: "easeOut" }}
            >
              <div>
                <p className="text-sm font-medium uppercase tracking-[0.28em] text-accent">Excel formulas and formatting</p>
                <h3 className="mt-4 text-3xl font-medium tracking-[-0.03em] text-ink">
                  Add totals, currency, percentage, and date formatting to your Excel export.
                </h3>
                <p className="mt-5 text-lg leading-8 text-ink-2">
                  Upload a CSV or Excel file, choose a calculation for each numeric column such as sum, average, count,
                  minimum, or maximum, and download a formatted Excel workbook with real formulas, a frozen header row,
                  autofilter, and a totals row.
                </p>
                <p className="mt-4 text-sm text-ink-3">Maximum file size is {MAX_FILE_SIZE_LABEL}.</p>
              </div>

              <div className="rounded-lg border border-line bg-canvas p-4 shadow-hairline sm:p-6" aria-busy={formulaProcessing}>
                <div
                  className="rounded-lg border border-dashed border-line bg-surface p-6 text-center transition hover:border-accent hover:bg-accent-tint"
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={(event) => {
                    event.preventDefault();
                    handleFormulaFileInput(event.dataTransfer.files);
                  }}
                >
                  <input
                    ref={formulaInputRef}
                    className="sr-only"
                    type="file"
                    accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    aria-label="Upload CSV or Excel file for formulas and formatting"
                    onChange={(event) => handleFormulaFileInput(event.target.files)}
                  />
                  <h4 className="text-2xl font-medium tracking-[-0.03em] text-ink">CSV or Excel upload</h4>
                  <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-ink-2">
                    Drop a spreadsheet with a header row. Numeric, currency, percentage, and date columns are detected
                    automatically.
                  </p>
                  <div className="mt-6 flex flex-col justify-center gap-3 sm:flex-row">
                    <button
                      className="rounded-md bg-ink px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent"
                      onClick={() => formulaInputRef.current?.click()}
                      disabled={formulaProcessing}
                    >
                      {formulaProcessing ? "Loading" : "Choose File"}
                    </button>
                  </div>
                </div>

                {formulaError ? <p className="mt-4 rounded-lg bg-surface p-4 text-sm font-medium text-error" role="alert">{formulaError}</p> : null}

                {formulaMatrix ? (
                  <div className="mt-6">
                    <h4 className="text-sm font-medium uppercase tracking-[0.18em] text-ink-3">Column calculations</h4>
                    <div className="mt-3 max-h-64 space-y-3 overflow-auto pr-1">
                      {formulaMatrix[0].map((header, columnIndex) => {
                        const type = formulaColumnTypes[columnIndex] ?? "text";
                        const isCalculable = type === "numeric" || type === "currency" || type === "percent";

                        return (
                          <div className="flex items-center justify-between gap-3 rounded-lg border border-line px-3 py-2" key={columnIndex}>
                            <div className="min-w-0">
                              <p className="truncate text-sm font-medium text-ink">{header || `Column ${columnIndex + 1}`}</p>
                              <p className="text-xs uppercase tracking-[0.12em] text-ink-3">{type}</p>
                            </div>
                            <select
                              className="rounded border border-line px-2 py-1 text-sm disabled:opacity-40"
                              value={aggregateChoices[columnIndex] ?? "none"}
                              disabled={!isCalculable}
                              onChange={(event) => updateAggregateChoice(columnIndex, event.target.value as AggregateChoice)}
                            >
                              {(["none", "sum", "average", "count", "min", "max"] as AggregateChoice[]).map((choice) => (
                                <option key={choice} value={choice}>
                                  {AGGREGATE_LABELS[choice]}
                                </option>
                              ))}
                            </select>
                          </div>
                        );
                      })}
                    </div>

                    <div className="mt-5">
                      <motion.button
                        className="rounded-md bg-accent px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent-tint"
                        onClick={downloadFormulaWorkbook}
                        whileTap={{ scale: 0.98 }}
                      >
                        Download Formatted Excel
                      </motion.button>
                    </div>
                    <p className="mt-4 text-xs leading-5 text-ink-3">
                      Formulas and number formatting are saved in the Excel file. CSV does not support formulas, so this
                      tool exports Excel only. Files are processed in your browser and are not uploaded to a server.
                    </p>
                  </div>
                ) : null}
              </div>
            </motion.div>
          ) : (
            <motion.div
              className="grid gap-10 lg:grid-cols-[0.82fr_1.18fr] lg:items-start"
              initial={{ opacity: 0, y: 28 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, ease: "easeOut" }}
            >
              <div>
                <p className="text-sm font-medium uppercase tracking-[0.28em] text-accent">Bank Ledger X</p>
                <h3 className="mt-4 text-3xl font-medium tracking-[-0.03em] text-ink">
                  Turn PDF bank statements into clean Excel, CSV, and QIF files.
                </h3>
                <p className="mt-5 text-lg leading-8 text-ink-2">
                  Upload a PDF statement from any bank, or paste statement text, and {PRODUCT_NAME} extracts every
                  transaction into a clean structure with normalized dates, descriptions, amounts, and balances. Export
                  import-ready files for QuickBooks, Xero, Sage, Excel, and any tool that reads Excel, CSV, or QIF.
                </p>
                <ul className="mt-5 space-y-2 text-sm text-ink-2">
                  <li>Works with digital PDF statements and text based exports</li>
                  <li>Normalizes dates, amounts, debits, credits, and balances</li>
                  <li>Exports Excel, CSV, and QIF for accounting software</li>
                  <li>Processed locally in your browser, up to {MAX_FILE_SIZE_LABEL}</li>
                </ul>
              </div>

              <div className="rounded-lg border border-line bg-canvas p-4 shadow-hairline sm:p-6" aria-busy={bankProcessing}>
                <div
                  className="rounded-lg border border-dashed border-line bg-surface p-6 text-center transition hover:border-accent hover:bg-accent-tint"
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={(event) => {
                    event.preventDefault();
                    handleBankFileInput(event.dataTransfer.files);
                  }}
                >
                  <input
                    ref={bankInputRef}
                    className="sr-only"
                    type="file"
                    accept=".pdf,.txt,.csv,application/pdf,text/plain,text/csv"
                    aria-label="Upload a bank statement PDF or text file"
                    onChange={(event) => handleBankFileInput(event.target.files)}
                  />
                  <h4 className="text-2xl font-medium tracking-[-0.03em] text-ink">Upload a bank statement</h4>
                  <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-ink-2">
                    Drop a PDF, TXT, or CSV statement, or choose a file. Scanned image-only PDFs may need to be pasted as
                    text below.
                  </p>
                  <button
                    className="mt-6 rounded-md bg-ink px-5 py-3 text-sm font-medium text-ink transition hover:bg-accent"
                    onClick={() => bankInputRef.current?.click()}
                    disabled={bankProcessing}
                  >
                    {bankProcessing ? "Extracting" : "Choose Statement"}
                  </button>
                </div>

                <div className="mt-4">
                  <label className="text-xs font-medium uppercase tracking-[0.14em] text-ink-3">
                    Or paste statement text
                  </label>
                  <textarea
                    className="mt-2 h-24 w-full rounded-lg border border-line px-3 py-2 text-sm"
                    placeholder="2026-06-03 Client payment 2,500.00 12,480.40"
                    value={bankPastedText}
                    onChange={(event) => setBankPastedText(event.target.value)}
                    aria-label="Paste bank statement text"
                  />
                  <button
                    className="mt-2 rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-ink"
                    onClick={convertPastedStatement}
                  >
                    Convert pasted text
                  </button>
                </div>

                {bankError ? <p className="mt-4 rounded-lg bg-surface p-4 text-sm font-medium text-error" role="alert">{bankError}</p> : null}

                {bankTransactions ? (
                  <div className="mt-6">
                    <div className="rounded-lg bg-ink p-4 text-ink">
                      <p className="text-3xl font-medium">{bankTransactions.length}</p>
                      <p className="mt-1 text-xs uppercase tracking-[0.18em] text-ink-2">Transactions extracted</p>
                    </div>

                    <div className="mt-5 flex flex-wrap gap-2">
                      <button className="rounded-md bg-accent px-4 py-2 text-xs font-medium text-ink transition hover:bg-accent-tint" onClick={() => downloadBankStatement("xlsx")}>Download Excel</button>
                      <button className="rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-ink" onClick={() => downloadBankStatement("csv")}>Download CSV</button>
                      <button className="rounded-md border border-line px-4 py-2 text-xs font-medium text-ink-2 transition hover:border-ink" onClick={() => downloadBankStatement("qif")}>Download QIF</button>
                    </div>

                    <div className="mt-6 overflow-hidden rounded-lg border border-line">
                      <div className="border-b border-line bg-surface px-4 py-3">
                        <h4 className="text-sm font-medium text-ink">Transaction preview</h4>
                      </div>
                      <div className="max-h-72 overflow-auto">
                        <table className="min-w-full text-left text-sm">
                          <thead className="bg-canvas text-xs uppercase tracking-[0.1em] text-ink-3">
                            <tr>
                              <th className="px-4 py-2 font-medium">Date</th>
                              <th className="px-4 py-2 font-medium">Description</th>
                              <th className="px-4 py-2 font-medium">Amount</th>
                              <th className="px-4 py-2 font-medium">Balance</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-line bg-canvas">
                            {bankTransactions.slice(0, 12).map((transaction, index) => (
                              <tr key={index}>
                                <td className="whitespace-nowrap px-4 py-2 text-ink-2">{transaction.date}</td>
                                <td className="max-w-56 truncate px-4 py-2 text-ink-2">{transaction.description}</td>
                                <td className="whitespace-nowrap px-4 py-2 text-ink-2">{transaction.amount || "-"}</td>
                                <td className="whitespace-nowrap px-4 py-2 text-ink-2">{transaction.balance || "-"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                    <p className="mt-4 text-xs leading-5 text-ink-3">
                      Statements are processed in your browser and are not uploaded to a server. Always review extracted
                      values before using them for accounting.
                    </p>
                  </div>
                ) : null}
              </div>
            </motion.div>
          )}
              </div>
              <WorkspaceSideWidget
                tab={workspaceTab}
                processing={isProcessing || converterProcessing || formulaProcessing || bankProcessing}
                hasResult={Boolean(result || converterFileName || formulaFileName || bankTransactions)}
              />
            </div>
          </div>
        </section>

        <section id="features" className="mc-hero-root relative overflow-hidden border-y border-line py-20 text-ink">
          <div className="mc-grid-overlay pointer-events-none absolute inset-0 opacity-40" aria-hidden="true" />
          <div className="relative mx-auto max-w-7xl px-5 lg:px-8">
            <motion.div
              className="max-w-3xl"
              initial={{ opacity: 0, y: 24 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-80px" }}
              transition={{ duration: 0.6, ease: "easeOut" }}
            >
              <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Platform capabilities</p>
              <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">
                Built for data cleaning, CSV processing, Excel automation and reconciliation.
              </h2>
              <p className="mc-prose mt-5 text-lg leading-relaxed text-ink-2">
                MarqClean AI helps operations, compliance, finance and data teams clean, validate, transform and
                reconcile Excel, CSV, PDF and financial datasets, all in the browser.
              </p>
            </motion.div>

            <motion.div
              className="mt-14 grid gap-4 lg:grid-cols-4 lg:grid-rows-3"
              initial="hidden"
              whileInView="show"
              viewport={{ once: true, margin: "-80px" }}
              variants={{ hidden: {}, show: { transition: { staggerChildren: 0.08 } } }}
            >
              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group relative overflow-hidden rounded-lg p-7 hover:border-accent  lg:col-span-2 lg:row-span-2"
              >
                <div className=" pointer-events-none absolute -right-10 -top-10 h-40 w-40 rounded-md bg-[rgb(var(--accent))]/20  transition group-hover:bg-[rgb(var(--accent))]/30" aria-hidden="true" />
                <span className="mc-mono relative inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">DATA CLEANING</span>
                <h3 className="mc-display relative mt-5 text-2xl font-medium tracking-[-0.02em] text-ink lg:text-3xl">Name formatting and capitalization</h3>
                <p className="relative mt-4 max-w-md text-[15px] leading-6 text-ink-2">
                  Converts inconsistent contact names, job titles, company names, cities, states, and countries into clean,
                  readable values for CRM imports and email personalization.
                </p>
                <div className="relative mt-6 flex flex-wrap gap-2 font-mono text-xs text-ink-2">
                  <span className="rounded-md border border-line bg-surface px-2 py-1 line-through decoration-error">john SMITH</span>
                  <span className="self-center text-[rgb(var(--accent-hover))]">→</span>
                  <span className="rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-2 py-1 text-[rgb(var(--accent-hover))]">John Smith</span>
                </div>
              </motion.div>

              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group rounded-lg p-6 hover:border-[rgb(var(--accent))]/40 lg:col-span-2"
              >
                <span className="mc-mono inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">DATA CLEANING</span>
                <h3 className="mc-display mt-4 text-xl font-medium tracking-[-0.02em] text-ink">Column repair and field detection</h3>
                <p className="mt-3 text-[15px] leading-6 text-ink-2">
                  Detects emails, phone numbers, websites, and common CRM field aliases even when marketers receive exports
                  with mixed-up or poorly named columns.
                </p>
              </motion.div>

              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group rounded-lg p-6 hover:border-[rgb(var(--accent))]/40"
              >
                <span className="mc-mono inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">DATA CLEANING</span>
                <h3 className="mc-display mt-4 text-xl font-medium tracking-[-0.02em] text-ink">Industry categorization</h3>
                <p className="mt-3 text-[15px] leading-6 text-ink-2">
                  Enriches company records with campaign-friendly industries such as SaaS, healthcare, ecommerce, and finance.
                </p>
              </motion.div>

              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group rounded-lg p-6 hover:border-[rgb(var(--accent))]/40"
              >
                <span className="mc-mono inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">EXCEL AUTOMATION</span>
                <h3 className="mc-display mt-4 text-xl font-medium tracking-[-0.02em] text-ink">CSV to Excel conversion</h3>
                <p className="mt-3 text-[15px] leading-6 text-ink-2">
                  Splits delimited text into columns using Excel's Text to Columns logic, then exports a clean XLSX workbook.
                </p>
              </motion.div>

              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group rounded-lg p-6 hover:border-[rgb(var(--accent))]/40 lg:col-span-2"
              >
                <span className="mc-mono inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">EXCEL AUTOMATION</span>
                <h3 className="mc-display mt-4 text-xl font-medium tracking-[-0.02em] text-ink">Excel formulas and number formatting</h3>
                <p className="mt-3 text-[15px] leading-6 text-ink-2">
                  Detects numeric, currency, percentage, and date columns, then generates a formatted Excel workbook with
                  a frozen header row, autofilter, banded rows, and a totals row powered by real Excel formulas.
                </p>
              </motion.div>

              <motion.div
                variants={{ hidden: { opacity: 0, y: 24 }, show: { opacity: 1, y: 0 } }}
                transition={{ duration: 0.5, ease: "easeOut" }}
                className="mc-glass mc-card-interactive group rounded-lg p-6 hover:border-[rgb(var(--accent))]/40 lg:col-span-2"
              >
                <span className="mc-mono inline-flex rounded-md border border-[rgb(var(--accent))]/30 bg-[rgb(var(--accent))]/10 px-3 py-1 text-[11px] font-medium text-[rgb(var(--accent-hover))]">EXCEL AUTOMATION</span>
                <h3 className="mc-display mt-4 text-xl font-medium tracking-[-0.02em] text-ink">Advanced Excel functions engine</h3>
                <p className="mt-3 text-[15px] leading-6 text-ink-2">
                  Run 40+ Excel functions without writing formulas, including IF, XLOOKUP, INDEX/MATCH, SUMIFS, COUNTIFS,
                  and text functions like TEXTJOIN and SUBSTITUTE, in Data Toolbox.
                </p>
              </motion.div>
            </motion.div>
          </div>
        </section>

        <section id="reconciliation" className="border-y border-line bg-[rgb(var(--surface))] py-20 text-ink">
          <div className="mx-auto max-w-7xl px-5 lg:px-8">
            <div className="grid gap-10 lg:grid-cols-[1.1fr_0.9fr] lg:items-center">
              <div>
                <p className="text-sm font-medium uppercase tracking-[0.28em] text-accent">Verification and reconciliation hub</p>
                <h2 className="mt-4 text-4xl font-medium tracking-[-0.04em] sm:text-5xl">
                  Verify PDF client mailing data against your Excel master file.
                </h2>
                <p className="mt-5 text-lg leading-8 text-ink-2">
                  A comprehensive document verification, extraction, and reconciliation hub for investor services,
                  transfer secretarial, compliance, and data quality teams. Upload an unstructured PDF mailing list and a
                  structured Excel master, auto-detect fields, and download colour-coded Excel reports with executive
                  summary, detailed verification, exceptions, PDF extraction output, master data, and data quality
                  sheets. Everything runs in your browser with no data stored on servers.
                </p>
                <div className="mt-8 flex flex-wrap gap-3">
                  <button
                    className="rounded-md bg-accent-tint px-6 py-3 text-sm font-medium text-ink transition hover:bg-canvas"
                    onClick={() => navigateToPage("reconciliation-hub")}
                  >
                    Open Reconciliation Hub
                  </button>
                  <PageLink className="rounded-md border border-line px-6 py-3 text-sm font-medium text-ink transition hover:border-line hover:bg-canvas hover:text-ink" page="data-reconciliation-tool">
                    Learn about reconciliation
                  </PageLink>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                {[
                  { h: "Data Verification", p: "Structure free-text PDF records into name, address lines, and postal code, then verify against Excel." },
                  { h: "File Comparison Centre", p: "Excel vs Excel, PDF vs Excel, CSV vs PDF, Word vs Excel, and multi-file." },
                  { h: "Auto field mapping", p: "Detects columns and maps them to mailing fields automatically." },
                  { h: "Colour-coded reports", p: "Green, blue, red, and orange statuses across six report sheets." },
                ].map((card) => (
                  <div key={card.h} className="rounded-lg border border-line bg-canvas/5 p-5">
                    <h3 className="text-lg font-medium">{card.h}</h3>
                    <p className="mt-2 text-sm text-ink-2">{card.p}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>

        <section id="data-toolbox" className="relative overflow-hidden bg-[rgb(var(--canvas))] py-20 text-ink">
          <div className="mc-grid-overlay pointer-events-none absolute inset-0 opacity-30" aria-hidden="true" />
          <div className="relative mx-auto max-w-7xl px-5 lg:px-8">
            <div className="max-w-3xl">
              <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Free data toolbox</p>
              <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">
                Advanced Excel functions plus merge, match, score, predict, and extract.
              </h2>
              <p className="mt-5 text-lg leading-8 text-ink-2">
                A consolidated toolbox of best-in-class data utilities, all unlocked and free. Run advanced Excel functions
                such as XLOOKUP, INDEX and MATCH, SUMIFS, COUNTIFS, SUMPRODUCT, FILTER, UNIQUE, SORT, LET, XNPV, XIRR, PMT,
                and IPMT from a dropdown or by typing a plain-English instruction. Then merge spreadsheets, remove exact and
                fuzzy duplicates, match names and addresses, score data quality, explore distributions, make one-click
                predictions, extract web tables, and build billable timesheets. Everything runs in your browser and handles
                large datasets without a server.
              </p>
              <button
                className="mc-cta-cyan mt-8 rounded-md px-6 py-3 text-sm font-medium"
                onClick={() => navigateToPage("data-toolbox")}
              >
                Open Data Toolbox
              </button>
            </div>
            <div className="mt-10 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              {[
                { h: "Advanced Excel Functions", p: "XLOOKUP, SUMIFS, FILTER, XIRR, PMT and more, by dropdown or plain English." },
                { h: "Merge & Combine", p: "Append or join CSV and Excel files with smart column matching." },
                { h: "Remove Duplicates", p: "Exact and fuzzy duplicate detection and removal." },
                { h: "Fuzzy Match", p: "Similarity search for names and addresses." },
                { h: "Clarity Score", p: "Deduplicate, gap-fill, autoformat, and score quality." },
                { h: "Data Explorer", p: "Filter, query, and view distributions." },
                { h: "Smart Predict", p: "One-click regression and classification." },
                { h: "Web Table Extractor", p: "Turn HTML tables and lists into spreadsheets." },
              ].map((card) => (
                <div key={card.h} className="mc-glass mc-card-interactive rounded-lg p-5 hover:border-[rgb(var(--accent))]/40">
                  <h3 className="mc-display text-lg font-medium text-ink">{card.h}</h3>
                  <p className="mt-2 text-sm text-ink-2">{card.p}</p>
                </div>
              ))}
            </div>
          </div>
        </section>

        <section id="excel-automation" className="relative overflow-hidden bg-[rgb(var(--surface))] py-20 text-ink">
          <div className="mx-auto max-w-7xl px-5 lg:px-8">
            <div className="excel-automation-layout">
              <div className="excel-automation-main">
                <div className="grid gap-10 lg:grid-cols-[1.1fr_0.9fr] lg:items-center">
                  <div>
                    <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Excel automation</p>
                    <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">
                      Excel Automation for sorting, subtotals and workbooks.
                    </h2>
                    <p className="mt-5 text-lg leading-8 text-ink-2">
                      Automate Excel sorting, subtotaling, number formatting, formula handling and workbook generation.
                      The built-in Sort &amp; Subtotal Engine groups related records by the root SRN or reference before the
                      first slash, subtotals each group, inserts a blank spacer row, and produces a clean workbook with a
                      Sorted_Output sheet, a Formula_Guide, and an Error_Report when needed. The original worksheet is
                      preserved and everything runs in your browser.
                    </p>
                    <button
                      className="mc-cta-cyan mt-8 rounded-md px-6 py-3 text-sm font-medium"
                      onClick={() => navigateToPage("excel-automation")}
                    >
                      Open Excel Automation
                    </button>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-2">
                    {[
                      { h: "Group by root reference", p: "Groups rows by the SRN or reference before the first slash." },
                      { h: "Subtotal each group", p: "Adds a bold SUM subtotal and a blank spacer row per group." },
                      { h: "Preview and map", p: "Preview detected columns and start row, then adjust the mapping." },
                      { h: "Preserves the original", p: "Outputs Sorted_Output, Formula_Guide, and Error_Report sheets." },
                    ].map((card) => (
                      <div key={card.h} className="mc-glass mc-card-interactive rounded-lg p-5 hover:border-[rgb(var(--accent))]/40">
                        <h3 className="mc-display text-lg font-medium text-ink">{card.h}</h3>
                        <p className="mt-2 text-sm text-ink-2">{card.p}</p>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <aside className="ws-side-widget excel-automation-widget" aria-label="Excel Automation guidance">
                <div className="ws-side-widget__header">
                  <span className="ws-side-widget__eyebrow">Automation workflow</span>
                  <span className="ws-side-widget__status">Ready</span>
                </div>
                <h3>Prepare → automate → review</h3>
                <p className="ws-side-widget__description">Keep the Excel Automation process structured from source workbook to final output.</p>
                <div className="ws-side-widget__section">
                  <div className="ws-side-widget__section-title">Workflow</div>
                  <ol className="ws-side-widget__steps">
                    <li className="is-current"><span>1</span><strong>Choose the workbook</strong></li>
                    <li className="is-current"><span>2</span><strong>Map and run automation</strong></li>
                    <li className="is-current"><span>3</span><strong>Review and export</strong></li>
                  </ol>
                </div>
                <div className="ws-side-widget__section">
                  <div className="ws-side-widget__section-title">Workspace tip</div>
                  <p className="ws-side-widget__tip">Review the detected grouping and column mapping before generating the final workbook.</p>
                </div>
                <div className="ws-side-widget__privacy">
                  <span aria-hidden="true">✓</span>
                  <div><strong>Browser-local processing</strong><small>Your workbook remains on this device while it is processed.</small></div>
                </div>
              </aside>
            </div>
          </div>
        </section>

        <section id="free-tools" className="bg-[rgb(var(--surface))] py-20 text-ink">
          <div className="mx-auto max-w-7xl px-5 lg:px-8">
            <div className="max-w-3xl">
              <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Free data tools</p>
              <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">
                One engine for CSV, Excel, data cleaning and reconciliation workflows.
              </h2>
              <p className="mc-prose mt-5 text-lg leading-relaxed text-ink-2">
                MarqClean AI uses a modular data architecture, so teams can solve specific spreadsheet problems with
                focused tools for CSV cleanup, Excel conversion, formatting, duplicates, validation, and reconciliation.
              </p>
            </div>

            <div className="mt-10">
              <ToolsDirectory />
            </div>
          </div>
        </section>

        <section id="workflow" className="relative overflow-hidden bg-[rgb(var(--canvas))] py-20 text-ink">
          <div className="relative mx-auto max-w-7xl px-5 lg:px-8">
            <div className="mx-auto max-w-3xl text-center">
              <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">How it works</p>
              <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">A clear path from raw files to decision-ready data.</h2>
              <p className="mt-5 text-lg leading-8 text-ink-2">Every workflow follows the same simple operating model: bring your file into the workspace, review what MarqClean AI identifies, then export the structured result you need.</p>
            </div>
            <div className="mt-12 grid gap-5 lg:grid-cols-3">
              {[
                { number: "01", title: "Bring your data", text: "Upload CSV, Excel or PDF files directly into the relevant workspace. Supported formats and the next action remain visible." },
                { number: "02", title: "Review the work", text: "Cleaning, conversion, formulas and reconciliation workflows show processing state and output previews before download." },
                { number: "03", title: "Export and move forward", text: "Download structured Excel, CSV or reconciliation output and continue your normal CRM, reporting, finance or operations workflow." },
              ].map((step) => (
                <article key={step.number} className="mc-glass rounded-lg p-6">
                  <span className="mc-mono inline-flex h-8 w-8 items-center justify-center rounded-md border border-[rgb(var(--accent))]/40 bg-[rgb(var(--accent))]/10 text-xs font-medium text-[rgb(var(--accent-hover))]">{step.number}</span>
                  <h3 className="mc-display mt-5 text-xl font-medium tracking-[-0.02em] text-ink">{step.title}</h3>
                  <p className="mt-3 text-[15px] leading-7 text-ink-2">{step.text}</p>
                </article>
              ))}
            </div>
            <div className="mt-8 grid gap-5 lg:grid-cols-2">
              <div className="rounded-lg border border-line bg-surface p-6">
                <p className="mc-mono text-xs uppercase tracking-[0.18em] text-[rgb(var(--accent))]">Workspace principle</p>
                <h3 className="mc-display mt-3 text-xl font-medium text-ink">One task, one focused workspace.</h3>
                <p className="mt-3 text-[15px] leading-7 text-ink-2">Each workspace is organized around a specific business outcome, with contextual guidance beside the working area so users always know what to do next.</p>
              </div>
              <div className="rounded-lg border border-line bg-surface p-6">
                <p className="mc-mono text-xs uppercase tracking-[0.18em] text-[rgb(var(--accent))]">Review principle</p>
                <h3 className="mc-display mt-3 text-xl font-medium text-ink">Inspect before you rely on the output.</h3>
                <p className="mt-3 text-[15px] leading-7 text-ink-2">Results remain reviewable before export, helping teams catch exceptions and confirm that the generated workbook or dataset matches the intended workflow.</p>
              </div>
            </div>
          </div>
        </section>

        <section id="faq" className="bg-[rgb(var(--canvas))] py-20 text-ink">
          <div className="mx-auto max-w-3xl px-5 lg:px-8">
            <p className="mc-mono text-sm font-medium uppercase tracking-[0.28em] text-[rgb(var(--accent))]">Frequently asked questions</p>
            <h2 className="mc-display mt-4 text-4xl font-medium tracking-[-0.04em] text-ink sm:text-5xl">Questions, answered honestly.</h2>
            <div className="mt-10 divide-y divide-white/10 border-y border-line">
              {[
                {
                  q: "Is MarqClean AI actually free?",
                  a: "Yes. Every tool listed in the Free Data Tools directory, the Quick Data & CSV Cleaner, Excel Automation, Data Toolbox, and Free Excel Academy run at no cost, with no account or card required.",
                },
                {
                  q: "Do my files ever leave my browser?",
                  a: "No. Cleaning, conversion, reconciliation, and formula processing all run locally in your browser using JavaScript. Nothing is uploaded to a server, so your data never leaves your device during processing.",
                },
                {
                  q: "What file types are supported?",
                  a: "CSV, XLSX, and XLS across most tools. Reconciliation Hub and Bank Ledger X also accept PDF statements and mailing lists, which are parsed directly in the browser.",
                },
                {
                  q: "How large a file can I upload?",
                  a: "The browser-safe upload cap is 100 MB. Actual processing capacity still depends on available device memory, so very large workbooks may take longer or need to be split before processing.",
                },
                {
                  q: "Does this replace Excel, Power Query, or Power BI?",
                  a: "No. MarqClean AI automates the repetitive cleaning, formatting, and reconciliation work that normally happens before or after using Excel. The Power Query-style workflow and Power Pivot-style data modelling tools in Data Toolbox reproduce the result of those workflows, not the native Microsoft engines — see each tool's in-product notice for exactly what is and isn't preserved.",
                },
                {
                  q: "How accurate is the automatic name, industry, or field detection?",
                  a: "Detection is rule- and keyword-based (not a black-box model), which makes it predictable and reviewable. Every tool shows its output before you download, so you can check and adjust results rather than trusting an opaque process.",
                },
                {
                  q: "Can I use this for client or investor reconciliation work?",
                  a: "Yes — Reconciliation Hub is built for exactly that: matching client, investor, account, and banking records across Excel, CSV, and PDF sources, with exception reporting and a colour-coded output.",
                },
              ].map((item) => (
                <details key={item.q} className="group py-5">
                  <summary className="flex cursor-pointer list-none items-center justify-between gap-4 text-left text-base font-medium text-ink marker:content-none">
                    {item.q}
                    <svg className="h-5 w-5 shrink-0 text-ink-2 transition group-open:rotate-45" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <path d="M12 5v14M5 12h14" />
                    </svg>
                  </summary>
                  <p className="mt-3 max-w-2xl text-[15px] leading-7 text-ink-2">{item.a}</p>
                </details>
              ))}
            </div>
          </div>
        </section>
      </main>
      )}

      <footer className="mc-site-footer" aria-labelledby="footer-heading">
        <div className="mc-site-footer__inner">
          <div className="mc-site-footer__top">
            <div className="mc-site-footer__brand">
              <a href="/" aria-label="MarqClean AI home" onClick={(event)=>{event.preventDefault();navigateToPage("home");}}>
                <Logo className="shrink-0" />
              </a>
              <p>Browser-local tools for cleaning, validating, transforming and reconciling Excel, CSV and PDF data.</p>
              <div className="mc-site-footer__trust"><span aria-hidden="true">●</span> Files stay in your browser during processing</div>
            </div>

            <div className="mc-site-footer__column">
              <h3>Platform</h3>
              <PageLink page="about">About MarqClean AI</PageLink>
              <a href="/data-cleaner" onClick={(event)=>{event.preventDefault();openWorkspaceTab("leads");}}>Quick Data &amp; CSV Cleaner</a>
              <a href="/excel-automation" onClick={(event)=>{event.preventDefault();navigateToPage("excel-automation");}}>Excel Automation</a>
              <a href="/reconciliation-hub" onClick={(event)=>{event.preventDefault();navigateToPage("reconciliation-hub");}}>Reconciliation Hub</a>
              <a href="/data-toolbox" onClick={(event)=>{event.preventDefault();navigateToPage("data-toolbox");}}>Data Toolbox</a>
            </div>

            <div className="mc-site-footer__column">
              <h3>Tools</h3>
              <PageLink page="clean-csv-file-online">CSV Cleaner</PageLink>
              <PageLink page="csv-to-excel-cleaner">CSV to Excel Converter</PageLink>
              <PageLink page="fix-csv-capitalization">Spreadsheet Formatter</PageLink>
              <PageLink page="remove-duplicates-excel">Duplicate Removal</PageLink>
              <PageLink page="document-validation-tool">Data Validation</PageLink>
              <PageLink page="advanced-excel-functions">Advanced Excel Functions</PageLink>
            </div>

            <div className="mc-site-footer__column">
              <h3>Trust &amp; legal</h3>
              <PageLink page="privacy">Privacy Policy</PageLink>
              <PageLink page="terms">Terms of Service</PageLink>
              <PageLink page="cookies">Cookie Policy</PageLink>
              <PageLink page="accessibility">Accessibility</PageLink>
              <PageLink page="contact">Contact &amp; Support</PageLink>
              <a href="mailto:support@vertexstreamtechnologies.com">support@vertexstreamtechnologies.com</a>
            </div>
          </div>

          <div className="mc-site-footer__bottom">
            <p>© {new Date().getFullYear()} MarqClean AI. All rights reserved.</p>
            <p>A product of <a href="https://www.vertexsg.co.za" target="_blank" rel="noopener noreferrer">Vertex Stream Group</a> · Vertex Stream Technologies · Johannesburg, South Africa</p>
          </div>
        </div>
      </footer>

      {renderBackToTop()}
    </div>
  );
}