/// <reference types="vite/client" />

declare module "*?url" {
  const src: string;
  export default src;
}

declare module "pdfjs-dist/build/pdf.worker.min.mjs";
